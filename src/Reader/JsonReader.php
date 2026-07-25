<?php

declare(strict_types=1);

namespace Mnb\PHPExcel\Reader;

use Mnb\PHPExcel\Support\Arr;
use Mnb\PHPExcel\Support\ErrorCode;
use Mnb\PHPExcel\Support\JsonArrayNormalizer;
use Mnb\PHPExcel\Support\MnbExcelException;

final class JsonReader implements IterableReaderInterface
{
    /** @return list<list<mixed>> */
    public function readSheet(string $path, int|string $sheet = 1, array $options = []): array
    {
        return array_values(iterator_to_array($this->iterateSheet($path, $sheet, $options), true));
    }

    /**
     * Streams NDJSON directly and streams standard top-level JSON arrays item by
     * item. Workbook-shaped JSON objects intentionally fall back to materialized
     * parsing because selecting a named sheet requires document-level structure.
     *
     * @return \Generator<int,list<mixed>>
     */
    public function iterateSheet(string $path, int|string $sheet = 1, array $options = []): iterable
    {
        $this->assertFileSize($path, $options);
        $projection = ColumnProjection::fromOptions($options);

        if ($this->shouldReadNdjson($path, $options)) {
            $this->assertSingleSheet($sheet);
            yield from $this->sliceIterable($this->tableRowsFromRecords($this->iterateNdjsonRecords($path, $options), $projection, $options), $options);
            return;
        }

        $streamingParser = new StreamingJsonArrayParser();
        $streamJsonArray = (bool) ($options['stream_json_array'] ?? true);
        if ($streamJsonArray && $streamingParser->isTopLevelArray($path, $options)) {
            $this->assertSingleSheet($sheet);
            yield from $this->sliceIterable($this->tableRowsFromRecords($streamingParser->parse($path, $options), $projection, $options), $options);
            return;
        }

        // Complex workbook objects retain the broad legacy shape support.
        $normalizer = new JsonArrayNormalizer();
        $workbook = $this->readWorkbook($path, array_replace($options, ['stream_json_array' => false]));
        $rows = $this->selectSheet($workbook, $sheet);
        $table = $normalizer->rowsToTable($rows, $options);
        foreach ($this->sliceTable($table, $options) as $index => $row) {
            yield $index => array_values($projection->project($row));
        }
    }

    /** @return array<string,list<array<int|string,mixed>>> */
    public function readWorkbook(string $path, array $options = []): array
    {
        $this->assertFileSize($path, $options);
        if ($this->shouldReadNdjson($path, $options)) {
            return $this->readNdjsonWorkbook($path, $options);
        }

        $workbook = (new JsonArrayNormalizer())->readWorkbook($path, $options);
        $maxSourceRows = isset($options['max_source_rows']) ? max(0, (int) $options['max_source_rows']) : null;
        if ($maxSourceRows !== null) {
            foreach ($workbook as $sheetName => $rows) {
                if (count($rows) > $maxSourceRows) {
                    throw MnbExcelException::withCode(
                        'JSON row limit exceeded in sheet ' . $sheetName . '. Rows: ' . count($rows) . ', max_source_rows: ' . $maxSourceRows,
                        ErrorCode::FILE_READ_FAILED,
                        ['path' => $path, 'sheet' => $sheetName, 'rows' => count($rows), 'max_source_rows' => $maxSourceRows]
                    );
                }
            }
        }
        return $workbook;
    }

    /** @return list<string> */
    public function sheetNames(string $path, array $options = []): array
    {
        if ($this->shouldReadNdjson($path, $options) || (new StreamingJsonArrayParser())->isTopLevelArray($path, $options)) {
            return [trim((string) ($options['sheet_name'] ?? 'Sheet1')) ?: 'Sheet1'];
        }
        return array_keys($this->readWorkbook($path, $options));
    }

    /** @param iterable<int,mixed> $records @return \Generator<int,list<mixed>> */
    private function tableRowsFromRecords(iterable $records, ColumnProjection $projection, array $options): \Generator
    {
        $normalizer = new JsonArrayNormalizer();
        $includeHeader = (bool) ($options['json_header_row'] ?? $options['include_header_row'] ?? true);
        $columns = isset($options['json_columns']) ? array_values(array_map('strval', (array) $options['json_columns'])) : [];
        $headerYielded = false;
        $output = 0;
        $extraKeys = strtolower((string) ($options['streaming_extra_keys'] ?? 'ignore'));
        if (!in_array($extraKeys, ['ignore', 'error'], true)) {
            throw new MnbExcelException('streaming_extra_keys must be "ignore" or "error".');
        }

        foreach ($records as $recordIndex => $decoded) {
            $single = $normalizer->workbookFromDecoded([$decoded], array_replace($options, ['json_mode' => 'rows', 'normalize_columns' => false]));
            $sheetRows = reset($single);
            $row = is_array($sheetRows) && isset($sheetRows[0]) && is_array($sheetRows[0]) ? $sheetRows[0] : [];
            $row = $projection->project($row);

            if (Arr::isAssoc($row)) {
                if ($columns === []) {
                    $columns = array_map('strval', array_keys($row));
                } elseif ($extraKeys === 'error') {
                    $unknown = array_values(array_diff(array_map('strval', array_keys($row)), $columns));
                    if ($unknown !== []) {
                        throw MnbExcelException::withCode(
                            'JSON item ' . ((int) $recordIndex + 1) . ' contains columns not present in the streaming schema: ' . implode(', ', $unknown),
                            ErrorCode::FILE_READ_FAILED,
                            ['item' => (int) $recordIndex + 1, 'columns' => $unknown]
                        );
                    }
                }

                if ($includeHeader && !$headerYielded) {
                    yield $output++ => $columns;
                    $headerYielded = true;
                }
                $line = [];
                foreach ($columns as $column) {
                    $line[] = $row[$column] ?? null;
                }
                yield $output++ => $line;
                continue;
            }

            if ($includeHeader && $columns !== [] && !$headerYielded) {
                yield $output++ => $columns;
                $headerYielded = true;
            }
            yield $output++ => array_values($row);
        }
    }

    /** @return \Generator<int,mixed> */
    private function iterateNdjsonRecords(string $path, array $options): \Generator
    {
        $handle = fopen($path, 'rb');
        if ($handle === false) {
            throw MnbExcelException::withCode('Unable to open JSON Lines file: ' . $path, ErrorCode::FILE_OPEN_FAILED, ['path' => $path]);
        }

        $lineNumber = 0;
        $records = 0;
        $maxSourceRows = isset($options['max_source_rows']) ? max(0, (int) $options['max_source_rows']) : null;
        $maxRecordBytes = max(1, (int) ($options['max_json_record_bytes'] ?? 16 * 1024 * 1024));
        $ignoreBlankLines = (bool) ($options['ignore_blank_lines'] ?? true);
        try {
            while (($line = fgets($handle)) !== false) {
                $lineNumber++;
                if (strlen($line) > $maxRecordBytes) {
                    throw MnbExcelException::withCode('JSON Lines record exceeds max_json_record_bytes at line ' . $lineNumber . '.', ErrorCode::FILE_READ_FAILED, ['line' => $lineNumber]);
                }
                if ($lineNumber === 1) {
                    $line = preg_replace('/^\xEF\xBB\xBF/', '', $line) ?? $line;
                }
                if (trim($line) === '') {
                    if ($ignoreBlankLines) {
                        continue;
                    }
                    throw MnbExcelException::withCode('Blank JSON Lines record at line ' . $lineNumber . '.', ErrorCode::JSON_INVALID, ['path' => $path, 'line' => $lineNumber]);
                }

                try {
                    $decoded = json_decode($line, true, (int) ($options['depth'] ?? 512), JSON_THROW_ON_ERROR | JSON_BIGINT_AS_STRING);
                } catch (\JsonException $e) {
                    throw MnbExcelException::withCode(
                        'Invalid JSON Lines record at line ' . $lineNumber . ': ' . $e->getMessage(),
                        ErrorCode::JSON_INVALID,
                        ['path' => $path, 'line' => $lineNumber],
                        $e
                    );
                }

                $records++;
                if ($maxSourceRows !== null && $records > $maxSourceRows) {
                    throw MnbExcelException::withCode(
                        'JSON Lines row limit exceeded. Rows: ' . $records . ', max_source_rows: ' . $maxSourceRows,
                        ErrorCode::FILE_READ_FAILED,
                        ['path' => $path, 'rows' => $records, 'max_source_rows' => $maxSourceRows]
                    );
                }
                yield $lineNumber - 1 => $decoded;
            }
        } finally {
            fclose($handle);
        }
    }

    /** @param iterable<int,list<mixed>> $rows @return \Generator<int,list<mixed>> */
    private function sliceIterable(iterable $rows, array $options): \Generator
    {
        $startRow = max(1, (int) ($options['start_row'] ?? 1));
        $endRow = isset($options['end_row']) ? max(1, (int) $options['end_row']) : null;
        $sourceSkipRows = max(0, (int) ($options['source_skip_rows'] ?? 0));
        $sourceLimitRows = isset($options['source_limit_rows']) ? max(0, (int) $options['source_limit_rows']) : null;
        $seen = 0;
        $yielded = 0;
        foreach ($rows as $row) {
            $seen++;
            if ($seen < $startRow || $seen <= $sourceSkipRows) {
                continue;
            }
            if ($endRow !== null && $seen > $endRow) {
                break;
            }
            if ($sourceLimitRows !== null && $yielded >= $sourceLimitRows) {
                break;
            }
            yield $seen - 1 => $row;
            $yielded++;
        }
    }

    /** @param array<string,list<array<int|string,mixed>>> $workbook @return list<array<int|string,mixed>> */
    private function selectSheet(array $workbook, int|string $sheet): array
    {
        if (is_string($sheet) && !ctype_digit($sheet)) {
            if (!array_key_exists($sheet, $workbook)) {
                throw new MnbExcelException('JSON sheet not found: ' . $sheet);
            }
            return $workbook[$sheet];
        }
        $index = max(1, (int) $sheet) - 1;
        $names = array_keys($workbook);
        if (!isset($names[$index])) {
            throw new MnbExcelException('JSON sheet index does not exist: ' . ((int) $sheet));
        }
        return $workbook[$names[$index]];
    }

    private function assertSingleSheet(int|string $sheet): void
    {
        if ($sheet !== 1 && $sheet !== '1' && $sheet !== 'Sheet1') {
            throw new MnbExcelException('A top-level JSON row array supports only one sheet.');
        }
    }

    private function shouldReadNdjson(string $path, array $options): bool
    {
        $legacyMode = strtolower(trim((string) ($options['json_mode'] ?? $options['mode'] ?? 'auto')));
        if (in_array($legacyMode, ['ndjson', 'jsonl', 'lines'], true)) {
            return true;
        }
        $documentMode = strtolower(trim((string) ($options['json_document_mode'] ?? 'auto')));
        if (in_array($documentMode, ['ndjson', 'jsonl', 'lines'], true)) {
            return true;
        }
        if (in_array($documentMode, ['json', 'document', 'standard'], true)) {
            return false;
        }
        if (!in_array($documentMode, ['', 'auto'], true)) {
            throw new MnbExcelException('json_document_mode must be "auto", "json", or "ndjson".');
        }
        $extension = strtolower((string) pathinfo($path, PATHINFO_EXTENSION));
        if (in_array($extension, ['jsonl', 'ndjson'], true)) {
            return true;
        }
        if (($options['detect_ndjson'] ?? true) === false) {
            return false;
        }

        $handle = fopen($path, 'rb');
        if ($handle === false) {
            throw MnbExcelException::withCode('Unable to open JSON file: ' . $path, ErrorCode::FILE_OPEN_FAILED, ['path' => $path]);
        }
        $lines = [];
        $maxLines = max(2, (int) ($options['ndjson_detection_lines'] ?? 20));
        try {
            while (count($lines) < $maxLines && ($line = fgets($handle)) !== false) {
                if ($lines === []) {
                    $line = preg_replace('/^\xEF\xBB\xBF/', '', $line) ?? $line;
                }
                $line = trim($line);
                if ($line !== '') {
                    $lines[] = $line;
                }
            }
        } finally {
            fclose($handle);
        }
        if (count($lines) < 2) {
            return false;
        }
        foreach ($lines as $line) {
            try {
                json_decode($line, true, (int) ($options['depth'] ?? 512), JSON_THROW_ON_ERROR | JSON_BIGINT_AS_STRING);
            } catch (\JsonException) {
                return false;
            }
        }
        try {
            json_decode(implode("\n", $lines), true, (int) ($options['depth'] ?? 512), JSON_THROW_ON_ERROR | JSON_BIGINT_AS_STRING);
            return false;
        } catch (\JsonException) {
            return true;
        }
    }

    /** @return array<string,list<array<int|string,mixed>>> */
    private function readNdjsonWorkbook(string $path, array $options): array
    {
        $normalizer = new JsonArrayNormalizer();
        $rows = [];
        foreach ($this->iterateNdjsonRecords($path, $options) as $decoded) {
            $single = $normalizer->workbookFromDecoded([$decoded], array_replace($options, ['json_mode' => 'rows']));
            $sheetRows = reset($single);
            foreach (is_array($sheetRows) ? $sheetRows : [] as $row) {
                $rows[] = $row;
            }
        }
        $sheetName = trim((string) ($options['sheet_name'] ?? 'Sheet1')) ?: 'Sheet1';
        return [$sheetName => $rows];
    }

    /** @param list<list<mixed>> $table @return list<list<mixed>> */
    private function sliceTable(array $table, array $options): array
    {
        $startRow = max(1, (int) ($options['start_row'] ?? 1));
        $endRow = isset($options['end_row']) ? max(1, (int) $options['end_row']) : null;
        $sourceSkipRows = max(0, (int) ($options['source_skip_rows'] ?? 0));
        $sourceLimitRows = isset($options['source_limit_rows']) ? max(0, (int) $options['source_limit_rows']) : null;
        $offset = max($startRow - 1, $sourceSkipRows);
        $length = $endRow !== null ? max(0, $endRow - $offset) : null;
        if ($sourceLimitRows !== null) {
            $length = $length === null ? $sourceLimitRows : min($length, $sourceLimitRows);
        }
        return array_values(array_slice($table, $offset, $length));
    }

    private function assertFileSize(string $path, array $options): void
    {
        if (!is_file($path)) {
            throw MnbExcelException::withCode('JSON file not found: ' . $path, ErrorCode::FILE_NOT_FOUND, ['path' => $path]);
        }
        $maxBytes = isset($options['max_file_bytes']) ? max(0, (int) $options['max_file_bytes']) : null;
        $size = filesize($path);
        if ($maxBytes !== null && $size !== false && $size > $maxBytes) {
            throw MnbExcelException::withCode(
                'JSON file exceeds max_file_bytes. Size: ' . $size . ', max_file_bytes: ' . $maxBytes,
                ErrorCode::FILE_READ_FAILED,
                ['path' => $path, 'size_bytes' => $size, 'max_file_bytes' => $maxBytes]
            );
        }
    }
}

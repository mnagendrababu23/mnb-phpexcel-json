<?php

declare(strict_types=1);

namespace Mnb\PHPExcel\Writer;

use Mnb\PHPExcel\Core\CellValue;
use Mnb\PHPExcel\Core\WorkbookData;
use Mnb\PHPExcel\Core\WorksheetData;
use Mnb\PHPExcel\Support\AtomicFileWriter;
use Mnb\PHPExcel\Support\ErrorCode;
use Mnb\PHPExcel\Support\MnbExcelException;

final class JsonWriter
{
    /**
     * @param array<string,mixed> $options
     */
    public function writeWorkbook(WorkbookData $workbook, string $path, array $options = []): void
    {
        $this->writeString($path, $this->workbookToString($workbook, $options));
    }

    /**
     * @param array<string,mixed> $options
     */
    public function workbookToString(WorkbookData $workbook, array $options = []): string
    {
        $payload = $this->workbookPayload($workbook, $options);
        return $this->encode($payload, $options);
    }

    /**
     * @param list<array<string,mixed>|list<mixed>> $rows
     * @param array<string,mixed> $options
     */
    public function writeRows(array $rows, string $path, array $options = []): void
    {
        $this->writeString($path, $this->rowsToString($rows, $options));
    }

    /**
     * @param list<array<string,mixed>|list<mixed>> $rows
     * @param array<string,mixed> $options
     */
    public function rowsToString(array $rows, array $options = []): string
    {
        return $this->encode($this->normalizeRows($rows), $options);
    }

    /** @param array<string,mixed> $options */
    private function workbookPayload(WorkbookData $workbook, array $options): mixed
    {
        $mode = (string) ($options['mode'] ?? 'auto');
        $includeMetadata = (bool) ($options['include_metadata'] ?? false);
        $includeSheetNames = (bool) ($options['include_sheet_names'] ?? false);

        if ($mode === 'rows' || ($mode === 'auto' && count($workbook->sheets) === 1 && !$includeMetadata)) {
            return $this->sheetRows($workbook->sheets[0], $options);
        }

        if ($mode === 'auto' && count($workbook->sheets) === 1) {
            $sheet = $workbook->sheets[0];
            $payload = [];
            if ($includeMetadata && $workbook->metadata !== []) {
                $payload['metadata'] = $this->normalizeValue($workbook->metadata);
            }
            if ($includeSheetNames) {
                $payload['sheet'] = $sheet->name;
            }
            $payload['rows'] = $this->sheetRows($sheet, $options);
            return $payload;
        }

        $sheets = [];
        foreach ($workbook->sheets as $sheet) {
            if ($includeSheetNames) {
                $sheets[] = [
                    'name' => $sheet->name,
                    'rows' => $this->sheetRows($sheet, $options),
                ];
            } else {
                $sheets[$sheet->name] = $this->sheetRows($sheet, $options);
            }
        }

        $payload = ['sheets' => $sheets];
        if ($includeMetadata && $workbook->metadata !== []) {
            $payload = ['metadata' => $this->normalizeValue($workbook->metadata)] + $payload;
        }

        return $payload;
    }

    /** @param array<string,mixed> $options @return list<array<string,mixed>|list<mixed>> */
    private function sheetRows(WorksheetData $sheet, array $options): array
    {
        return $this->normalizeRows($sheet->rowsForStructuredExport(
            preserveAssociative: (bool) ($options['preserve_associative_rows'] ?? true),
            dataOnly: (bool) ($options['data_only'] ?? false)
        ));
    }

    /**
     * @param list<array<string,mixed>|list<mixed>|array<int,mixed>> $rows
     * @return list<array<string,mixed>|list<mixed>>
     */
    private function normalizeRows(array $rows): array
    {
        $normalized = [];
        foreach ($rows as $row) {
            if (!is_array($row)) {
                $normalized[] = [$this->normalizeValue($row)];
                continue;
            }

            $item = [];
            foreach ($row as $key => $value) {
                if (is_int($key)) {
                    $item[$key] = $this->normalizeValue($value);
                } else {
                    $item[(string) $key] = $this->normalizeValue($value);
                }
            }
            $normalized[] = $item;
        }

        return $normalized;
    }

    private function normalizeValue(mixed $value): mixed
    {
        if ($value instanceof CellValue) {
            return $value->displayValue();
        }
        if ($value instanceof \DateTimeInterface) {
            return $value->format('Y-m-d H:i:s');
        }
        if (is_array($value)) {
            $out = [];
            foreach ($value as $key => $item) {
                $out[$key] = $this->normalizeValue($item);
            }
            return $out;
        }
        if (is_scalar($value) || $value === null) {
            return $value;
        }

        return (string) $value;
    }

    /**
     * Encode any JSON-safe payload and return the JSON string.
     *
     * This is useful for structured workbook outputs where the payload
     * is already prepared by a reader/session and should be returned
     * directly to a controller, API response, or variable without saving.
     *
     * @param array<string,mixed> $options
     */
    public function payloadToString(mixed $payload, array $options = []): string
    {
        return $this->encode($payload, $options);
    }

    /** @param array<string,mixed> $options */
    private function encode(mixed $payload, array $options): string
    {
        $flags = JSON_UNESCAPED_UNICODE | JSON_UNESCAPED_SLASHES;
        if ((bool) ($options['pretty'] ?? true)) {
            $flags |= JSON_PRETTY_PRINT;
        }
        if ((bool) ($options['preserve_zero_fraction'] ?? true)) {
            $flags |= JSON_PRESERVE_ZERO_FRACTION;
        }
        if ((bool) ($options['throw_on_error'] ?? true)) {
            $flags |= JSON_THROW_ON_ERROR;
        }

        try {
            $json = json_encode($payload, $flags);
        } catch (\JsonException $e) {
            throw MnbExcelException::withCode('Unable to encode JSON: ' . $e->getMessage(), ErrorCode::JSON_ENCODE_FAILED, [], $e);
        }

        if (!is_string($json)) {
            throw MnbExcelException::withCode('Unable to encode JSON.', ErrorCode::JSON_WRITE_FAILED);
        }

        return $json . ((bool) ($options['trailing_newline'] ?? true) ? "\n" : '');
    }

    private function writeString(string $path, string $contents): void
    {
        AtomicFileWriter::writeString($path, $contents, ErrorCode::JSON_WRITE_FAILED);
    }
}

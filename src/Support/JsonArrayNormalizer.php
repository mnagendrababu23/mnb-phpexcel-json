<?php

declare(strict_types=1);

namespace Mnb\PHPExcel\Support;

final class JsonArrayNormalizer
{
    /**
     * Read JSON from disk and convert it to a workbook array suitable for WorkbookBuilder.
     *
     * Supported JSON shapes:
     * - list of objects: [{"name":"Ravi"}]
     * - single object row: {"name":"Ravi"}
     * - object with rows key: {"rows":[...]}
     * - workbook map: {"Students":[...], "Teachers":[...]}
     * - workbook object: {"sheets":{"Students":[...]}}
     * - workbook list: {"sheets":[{"name":"Students", "rows":[...]}]}
     *
     * @param array<string,mixed> $options
     * @return array<string, list<array<int|string,mixed>>>
     */
    public function readWorkbook(string $path, array $options = []): array
    {
        if (!is_file($path)) {
            throw new MnbExcelException('JSON file not found: ' . $path);
        }

        $contents = file_get_contents($path);
        if ($contents === false) {
            throw new MnbExcelException('Unable to read JSON file: ' . $path);
        }

        return $this->workbookFromJsonString($contents, $options + ['path' => $path]);
    }

    /**
     * @param array<string,mixed> $options
     * @return array<string, list<array<int|string,mixed>>>
     */
    public function workbookFromJsonString(string $json, array $options = []): array
    {
        $json = $this->stripBom($json);
        try {
            $decoded = json_decode($json, true, (int) ($options['depth'] ?? 512), JSON_THROW_ON_ERROR | JSON_BIGINT_AS_STRING);
        } catch (\JsonException $e) {
            $path = isset($options['path']) ? ' in ' . $options['path'] : '';
            throw new MnbExcelException('Invalid JSON' . $path . ': ' . $e->getMessage());
        }

        return $this->workbookFromDecoded($decoded, $options);
    }

    /**
     * @param array<string,mixed> $options
     * @return array<string, list<array<int|string,mixed>>>
     */
    public function workbookFromDecoded(mixed $decoded, array $options = []): array
    {
        $sheetName = (string) ($options['sheet_name'] ?? 'Sheet1');
        $mode = strtolower((string) ($options['json_mode'] ?? $options['mode'] ?? 'auto'));

        if ($mode === 'rows') {
            return [$this->safeSheetName($sheetName) => $this->normalizeRows($decoded, $options)];
        }

        if (is_array($decoded) && Arr::isAssoc($decoded)) {
            if (isset($decoded['sheets'])) {
                return $this->normalizeSheetsValue($decoded['sheets'], $options);
            }

            if (isset($decoded['rows']) && is_array($decoded['rows'])) {
                return [$this->safeSheetName((string) ($decoded['sheet'] ?? $sheetName)) => $this->normalizeRows($decoded['rows'], $options)];
            }

            if ($mode === 'workbook' || $this->looksLikeSheetMap($decoded)) {
                return $this->normalizeSheetsValue($decoded, $options);
            }

            return [$this->safeSheetName($sheetName) => $this->normalizeRows([$decoded], $options)];
        }

        return [$this->safeSheetName($sheetName) => $this->normalizeRows($decoded, $options)];
    }

    /**
     * Convert workbook arrays into raw table rows, optionally adding headers for associative rows.
     *
     * @param list<array<int|string,mixed>> $rows
     * @return list<list<mixed>>
     */
    public function rowsToTable(array $rows, array $options = []): array
    {
        $includeHeader = (bool) ($options['json_header_row'] ?? $options['include_header_row'] ?? true);
        if ($rows === []) {
            return [];
        }

        $first = reset($rows);
        if (!is_array($first) || !Arr::isAssoc($first)) {
            return array_map(static fn(array $row): array => array_values($row), $rows);
        }

        $columns = $this->collectColumns($rows);
        $table = [];
        if ($includeHeader) {
            $table[] = $columns;
        }

        foreach ($rows as $row) {
            $line = [];
            foreach ($columns as $column) {
                $line[] = $row[$column] ?? null;
            }
            $table[] = $line;
        }

        return $table;
    }

    /**
     * @param mixed $value
     * @return list<array<int|string,mixed>>
     */
    private function normalizeRows(mixed $value, array $options): array
    {
        if ($value === null) {
            return [];
        }

        if (!is_array($value)) {
            return [[$value]];
        }

        if ($value === []) {
            return [];
        }

        $rows = Arr::isAssoc($value) ? [$value] : $value;
        $normalized = [];

        foreach ($rows as $row) {
            if (!is_array($row)) {
                $normalized[] = [$row];
                continue;
            }

            /** @var array<int|string,mixed> $row */
            $normalized[] = $this->normalizeRow($row, $options);
        }

        if ((bool) ($options['normalize_columns'] ?? true)) {
            $normalized = $this->normalizeColumns($normalized);
        }

        return $normalized;
    }

    /** @param array<int|string,mixed> $row @return array<int|string,mixed> */
    private function normalizeRow(array $row, array $options): array
    {
        if (!Arr::isAssoc($row)) {
            $out = [];
            foreach (array_values($row) as $value) {
                $out[] = $this->normalizeValue($value, $options);
            }
            return $out;
        }

        if ((bool) ($options['flatten_nested_keys'] ?? true)) {
            $flattened = [];
            $this->flattenAssoc($row, '', $flattened, (string) ($options['nested_separator'] ?? '.'));
            return $flattened;
        }

        $out = [];
        foreach ($row as $key => $value) {
            $out[(string) $key] = $this->normalizeValue($value, $options);
        }
        return $out;
    }

    /** @param array<int|string,mixed> $row @param array<string,mixed> $out */
    private function flattenAssoc(array $row, string $prefix, array &$out, string $separator): void
    {
        foreach ($row as $key => $value) {
            $path = $prefix === '' ? (string) $key : $prefix . $separator . (string) $key;
            if (is_array($value) && Arr::isAssoc($value)) {
                $this->flattenAssoc($value, $path, $out, $separator);
                continue;
            }
            $out[$path] = $this->normalizeValue($value, ['nested_separator' => $separator]);
        }
    }

    private function normalizeValue(mixed $value, array $options): mixed
    {
        if (is_array($value)) {
            if ($value === []) {
                return null;
            }
            if ((bool) ($options['json_encode_nested_lists'] ?? true)) {
                try {
                    return json_encode($value, JSON_UNESCAPED_UNICODE | JSON_UNESCAPED_SLASHES | JSON_THROW_ON_ERROR);
                } catch (\JsonException) {
                    return (string) count($value) . ' items';
                }
            }
            return $value;
        }

        return $value;
    }

    /**
     * @param array<string,mixed> $sheets
     * @return array<string, list<array<int|string,mixed>>>
     */
    private function normalizeSheetsValue(mixed $sheets, array $options): array
    {
        if (!is_array($sheets) || $sheets === []) {
            throw new MnbExcelException('Workbook JSON must contain at least one sheet.');
        }

        $workbook = [];

        if (!Arr::isAssoc($sheets)) {
            foreach ($sheets as $index => $sheet) {
                if (!is_array($sheet)) {
                    throw new MnbExcelException('Each JSON workbook sheet must be an object or array.');
                }
                $name = (string) ($sheet['name'] ?? 'Sheet' . ((int) $index + 1));
                $rows = $sheet['rows'] ?? [];
                $workbook[$this->uniqueSheetName($workbook, $name)] = $this->normalizeRows($rows, $options);
            }
            return $workbook;
        }

        foreach ($sheets as $name => $rows) {
            $workbook[$this->uniqueSheetName($workbook, (string) $name)] = $this->normalizeRows($rows, $options);
        }

        return $workbook;
    }

    /** @param array<string,mixed> $value */
    private function looksLikeSheetMap(array $value): bool
    {
        if ($value === []) {
            return false;
        }

        foreach ($value as $sheetRows) {
            if (!is_array($sheetRows)) {
                return false;
            }
            if ($sheetRows === []) {
                continue;
            }
            $first = reset($sheetRows);
            if (!is_array($first)) {
                return false;
            }
        }

        return true;
    }

    /** @param list<array<int|string,mixed>> $rows @return list<array<int|string,mixed>> */
    private function normalizeColumns(array $rows): array
    {
        $first = reset($rows);
        if (!is_array($first) || !Arr::isAssoc($first)) {
            return $rows;
        }

        $columns = $this->collectColumns($rows);
        $normalized = [];
        foreach ($rows as $row) {
            if (!Arr::isAssoc($row)) {
                $normalized[] = $row;
                continue;
            }
            $item = [];
            foreach ($columns as $column) {
                $item[$column] = $row[$column] ?? null;
            }
            $normalized[] = $item;
        }

        return $normalized;
    }

    /** @param list<array<int|string,mixed>> $rows @return list<string> */
    private function collectColumns(array $rows): array
    {
        $columns = [];
        foreach ($rows as $row) {
            if (!is_array($row) || !Arr::isAssoc($row)) {
                continue;
            }
            foreach ($row as $key => $_value) {
                $key = (string) $key;
                if (!in_array($key, $columns, true)) {
                    $columns[] = $key;
                }
            }
        }
        return $columns;
    }

    /** @param array<string,mixed> $existing */
    private function uniqueSheetName(array $existing, string $name): string
    {
        $base = $this->safeSheetName($name);
        $name = $base;
        $i = 2;
        while (array_key_exists($name, $existing)) {
            $suffix = ' ' . $i;
            $name = substr($base, 0, max(1, 31 - strlen($suffix))) . $suffix;
            $i++;
        }
        return $name;
    }

    private function safeSheetName(string $name): string
    {
        $name = trim(str_replace(['\\', '/', '?', '*', '[', ']', ':'], ' ', $name));
        $name = $name === '' ? 'Sheet' : $name;
        return substr($name, 0, 31);
    }

    private function stripBom(string $contents): string
    {
        return preg_replace('/^\xEF\xBB\xBF/', '', $contents) ?? $contents;
    }
}

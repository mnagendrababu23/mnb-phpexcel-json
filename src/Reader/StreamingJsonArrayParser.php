<?php

declare(strict_types=1);

namespace Mnb\PHPExcel\Reader;

use Mnb\PHPExcel\Support\ErrorCode;
use Mnb\PHPExcel\Support\MnbExcelException;

/**
 * Forward-only parser for a standard JSON document whose root value is an array.
 * Each array item is decoded independently, so the complete document is never
 * held in memory.
 */
final class StreamingJsonArrayParser
{
    /** @param array<string,mixed> $options @return \Generator<int,mixed> */
    public function parse(string $path, array $options = []): \Generator
    {
        if (!is_file($path)) {
            throw MnbExcelException::withCode('JSON file not found: ' . $path, ErrorCode::FILE_NOT_FOUND, ['path' => $path]);
        }

        $stream = new JsonCharacterStream($path, max(4096, (int) ($options['json_chunk_bytes'] ?? 65536)));
        $depth = max(1, (int) ($options['depth'] ?? 512));
        $maxRecordBytes = max(1, (int) ($options['max_json_record_bytes'] ?? 16 * 1024 * 1024));
        $maxRows = isset($options['max_source_rows']) ? max(0, (int) $options['max_source_rows']) : null;
        $index = 0;

        try {
            $stream->skipWhitespaceAndBom();
            if ($stream->get() !== '[') {
                throw MnbExcelException::withCode(
                    'Streaming JSON mode requires a top-level array.',
                    ErrorCode::JSON_INVALID,
                    ['path' => $path]
                );
            }

            while (true) {
                $stream->skipWhitespace();
                $next = $stream->peek();
                if ($next === ']') {
                    $stream->get();
                    break;
                }
                if ($next === null) {
                    throw $this->invalid($path, $stream, 'Unexpected end of JSON array.');
                }

                $raw = $this->readValue($stream, $maxRecordBytes, $path);
                try {
                    $value = json_decode($raw, true, $depth, JSON_THROW_ON_ERROR | JSON_BIGINT_AS_STRING);
                } catch (\JsonException $e) {
                    throw MnbExcelException::withCode(
                        'Invalid JSON array item ' . ($index + 1) . ': ' . $e->getMessage(),
                        ErrorCode::JSON_INVALID,
                        ['path' => $path, 'item' => $index + 1, 'line' => $stream->line(), 'column' => $stream->column()],
                        $e
                    );
                }

                $index++;
                if ($maxRows !== null && $index > $maxRows) {
                    throw MnbExcelException::withCode(
                        'JSON row limit exceeded. Rows: ' . $index . ', max_source_rows: ' . $maxRows,
                        ErrorCode::FILE_READ_FAILED,
                        ['path' => $path, 'rows' => $index, 'max_source_rows' => $maxRows]
                    );
                }
                yield $index - 1 => $value;

                $stream->skipWhitespace();
                $delimiter = $stream->get();
                if ($delimiter === ']') {
                    break;
                }
                if ($delimiter !== ',') {
                    throw $this->invalid($path, $stream, 'Expected a comma or closing bracket after JSON array item.');
                }
                $stream->skipWhitespace();
                if ($stream->peek() === ']') {
                    throw $this->invalid($path, $stream, 'Trailing commas are not valid JSON.');
                }
            }

            $stream->skipWhitespace();
            if ($stream->peek() !== null) {
                throw $this->invalid($path, $stream, 'Unexpected content after the JSON array.');
            }
        } finally {
            $stream->close();
        }
    }

    /** @param array<string,mixed> $options */
    public function isTopLevelArray(string $path, array $options = []): bool
    {
        if (!is_file($path)) {
            return false;
        }
        $stream = new JsonCharacterStream($path, max(4096, (int) ($options['json_chunk_bytes'] ?? 65536)));
        try {
            $stream->skipWhitespaceAndBom();
            return $stream->peek() === '[';
        } finally {
            $stream->close();
        }
    }

    private function readValue(JsonCharacterStream $stream, int $maxBytes, string $path): string
    {
        $first = $stream->peek();
        if ($first === null) {
            throw $this->invalid($path, $stream, 'Missing JSON array value.');
        }

        $raw = '';
        $inString = false;
        $escaped = false;
        $nested = 0;
        $compound = $first === '{' || $first === '[';
        $quoted = $first === '"';

        while (($char = $stream->peek()) !== null) {
            if (!$inString && !$compound && !$quoted && ($char === ',' || $char === ']')) {
                break;
            }
            if (!$inString && $compound && $nested === 0 && $raw !== '') {
                break;
            }
            if (!$inString && $quoted && $raw !== '' && substr($raw, -1) === '"') {
                break;
            }

            $char = (string) $stream->get();
            $raw .= $char;
            if (strlen($raw) > $maxBytes) {
                throw MnbExcelException::withCode(
                    'One JSON array item exceeds max_json_record_bytes.',
                    ErrorCode::FILE_READ_FAILED,
                    ['path' => $path, 'max_json_record_bytes' => $maxBytes, 'line' => $stream->line()]
                );
            }

            if ($inString) {
                if ($escaped) {
                    $escaped = false;
                } elseif ($char === '\\') {
                    $escaped = true;
                } elseif ($char === '"') {
                    $inString = false;
                    if ($quoted && $nested === 0) {
                        break;
                    }
                }
                continue;
            }

            if ($char === '"') {
                $inString = true;
                continue;
            }
            if ($char === '{' || $char === '[') {
                $nested++;
                continue;
            }
            if ($char === '}' || $char === ']') {
                $nested--;
                if ($compound && $nested === 0) {
                    break;
                }
                if ($nested < 0) {
                    throw $this->invalid($path, $stream, 'Unexpected closing token inside JSON value.');
                }
            }
        }

        $raw = trim($raw);
        if ($raw === '') {
            throw $this->invalid($path, $stream, 'Missing JSON array value.');
        }
        if ($inString || ($compound && $nested !== 0)) {
            throw $this->invalid($path, $stream, 'Unterminated JSON array value.');
        }

        return $raw;
    }

    private function invalid(string $path, JsonCharacterStream $stream, string $message): MnbExcelException
    {
        return MnbExcelException::withCode(
            $message . ' Near line ' . $stream->line() . ', column ' . $stream->column() . '.',
            ErrorCode::JSON_INVALID,
            ['path' => $path, 'line' => $stream->line(), 'column' => $stream->column()]
        );
    }
}

/** @internal */
final class JsonCharacterStream
{
    /** @var resource */
    private $handle;
    private string $buffer = '';
    private int $offset = 0;
    private bool $eof = false;
    private int $line = 1;
    private int $column = 0;

    public function __construct(string $path, private readonly int $chunkBytes)
    {
        $handle = fopen($path, 'rb');
        if ($handle === false) {
            throw MnbExcelException::withCode('Unable to open JSON file: ' . $path, ErrorCode::FILE_OPEN_FAILED, ['path' => $path]);
        }
        $this->handle = $handle;
    }

    public function peek(): ?string
    {
        if (!$this->ensureBuffer()) {
            return null;
        }
        return $this->buffer[$this->offset];
    }

    public function get(): ?string
    {
        $char = $this->peek();
        if ($char === null) {
            return null;
        }
        $this->offset++;
        if ($char === "\n") {
            $this->line++;
            $this->column = 0;
        } else {
            $this->column++;
        }
        return $char;
    }

    public function skipWhitespace(): void
    {
        while (($char = $this->peek()) !== null && ($char === ' ' || $char === "\t" || $char === "\r" || $char === "\n")) {
            $this->get();
        }
    }

    public function skipWhitespaceAndBom(): void
    {
        if ($this->peek() === "\xEF") {
            $bom = '';
            for ($i = 0; $i < 3; $i++) {
                $bom .= (string) $this->get();
            }
            if ($bom !== "\xEF\xBB\xBF") {
                throw MnbExcelException::withCode('Unsupported byte-order mark in JSON input.', ErrorCode::JSON_INVALID);
            }
        }
        $this->skipWhitespace();
    }

    public function line(): int
    {
        return $this->line;
    }

    public function column(): int
    {
        return max(1, $this->column);
    }

    public function close(): void
    {
        if (is_resource($this->handle)) {
            fclose($this->handle);
        }
    }

    private function ensureBuffer(): bool
    {
        if ($this->offset < strlen($this->buffer)) {
            return true;
        }
        if ($this->eof) {
            return false;
        }
        $chunk = fread($this->handle, $this->chunkBytes);
        if ($chunk === false) {
            throw MnbExcelException::withCode('Unable to read JSON stream.', ErrorCode::FILE_READ_FAILED);
        }
        if ($chunk === '') {
            $this->eof = true;
            return false;
        }
        $this->buffer = $chunk;
        $this->offset = 0;
        return true;
    }
}

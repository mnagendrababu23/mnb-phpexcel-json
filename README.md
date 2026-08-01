# MNB PHPExcel JSON

<<<<<<< HEAD
Streaming JSON and NDJSON reader/writer module for MNB PHPExcel.
Documentation URL: https://mnbphpexcel.space/getting-started/installation
This package is generated from the MNB PHPExcel monorepo. Do not copy source files between modules manually.
## MNB PHPExcel Assistant

Generate MNB PHPExcel code using our dedicated ChatGPT assistant:

[Open MNB PHPExcel AI Assistant](https://chatgpt.com/g/g-6a6e31d80350819194b68853d41c1561-mnb-phpexcel-assistant)
## Install
=======
Independent JSON and NDJSON reader/writer module. Requires only core and `ext-json`.
>>>>>>> 0a6708a (Release v2.0.0)

```bash
composer require mnb/mnb-phpexcel-json:^2.0
```

```php
use Mnb\PHPExcel\Format\Json;

$rows = Json::read('customers.json')->withHeaderRow()->toArray();
Json::write($rows, 'customers-export.json');
```

Supports top-level arrays, workbook-shaped JSON, NDJSON, streaming array parsing, sheet names, projection, and structured export.

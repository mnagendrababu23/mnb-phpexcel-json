# MNB PHPExcel JSON

<<<<<<< HEAD
Streaming JSON and NDJSON reader/writer module for MNB PHPExcel.
Documentation URL: https://mnbphpexcel.space/getting-started/installation
This package is generated from the MNB PHPExcel monorepo. Do not copy source files between modules manually.

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

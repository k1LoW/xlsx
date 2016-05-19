# xlsx: PHPExcel wrapper [![Travis](https://img.shields.io/travis/k1LoW/xlsx.svg)](https://travis-ci.org/k1LoW/xlsx)

## Usage

```php
<?php
use Xlsx\Xlsx;

$xlsx = new Xlsx();
$xlsx->read('/path/to/template.xlsx')
    ->setValue('testvalue', [
        'col' => 'B',
        'row' => '10',
    ])
    ->setValue('testvalue_with_sheet', [
        'col' => 'B',
        'row' => '10',
        'sheet' => 2,
    ])
    ->setValue('testvalue_with_border', [
        'col' => 'C',
        'row' => '10',
        'border' => [
            'top' => PHPExcel_Style_Border::BORDER_THICK,
            'right' => PHPExcel_Style_Border::BORDER_MEDIUM,
            'left' => PHPExcel_Style_Border::BORDER_THIN,
            'bottom' => PHPExcel_Style_Border::BORDER_DOUBLE,
        ],
    ])
    ->setValue('testvalue_with_border', [
        'col' => 'E',
        'row' => '10',
        'border' => PHPExcel_Style_Border::BORDER_THICK,
    ])
    ->setValue('testvalue_with_color', [
        'col' => 'F',
        'row' => '10',
        'color' => PHPExcel_Style_Color::COLOR_BLUE,
    ])
    ->setValue('testvalue_with_backgroud_color', [
        'col' => 'G',
        'row' => '10',
        'backgroundColor' => PHPExcel_Style_Color::COLOR_YELLOW,
    ])
    ->write('/path/to/output.xlsx');
```

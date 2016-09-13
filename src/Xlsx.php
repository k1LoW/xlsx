<?php

namespace Xlsx;

if (file_exists(dirname(__FILE__).'/../vendor/autoload.php')) {
    require_once dirname(__FILE__).'/../vendor/autoload.php';
}

use PHPExcel;
use PHPExcel_IOFactory;
use PHPExcel_Style_Fill;

/**
 * Xlsx.
 */
class Xlsx
{
    public $xlsx;
    private $data;

    public function __construct($templateFilePath = null)
    {
        $this->xlsx = null;
        $this->data = [];
        if (file_exists($templateFilePath)) {
            $this->read($templateFilePath);
        }
    }

    /**
     * __call.
     *
     * @param $method, $args
     */
    public function __call($method, $args)
    {
        if (!$this->xlsx) {
            return;
        }

        return call_user_func_array([$this->xlsx, $method], $args);
    }

    /**
     * read.
     */
    public function read($filePath)
    {
        $xlsReader = PHPExcel_IOFactory::createReaderForFile($filePath);
        $this->type = preg_replace('/^.+_/', '', get_class($xlsReader));
        $this->xlsx = $xlsReader->load($filePath);

        return $this;
    }

    /**
     * set.
     */
    public function set($key, $value = null)
    {
        if (is_array($key)) {
            foreach ($key as $k => $v) {
                $this->data[$k] = (string) $v;
            }
        } else {
            $this->data[$key] = $value;
        }

        return $this;
    }

    /**
     * write.
     */
    public function write($outputFilePath, $data = [])
    {
        if (!empty($data)) {
            $this->set($data);
        }
        if (!$this->xlsx) {
            $this->read($outputFilePath);
        }

        $sheets = $this->xlsx->getAllSheets();

        if (!empty($this->data)) {
            foreach ($sheets as $key => $sheet) {
                // Replace sheet title
                $title = $sheet->getTitle();
                if ($this->replaceValue($title)) {
                    $sheet->setTitle($this->replaceValue($title));
                }

                $rMax = $sheet->getHighestRow();
                $cMax = $sheet->getHighestColumn();
                for ($r = 1; $r <= $rMax; ++$r) {
                    for ($c = 0; $c <= self::alphabetToNumber($cMax); ++$c) {
                        $cell = $sheet->getCellByColumnAndRow($c, $r);
                        $value = $cell->getValue();
                        if (is_object($value)) {
                            $value = $cell->getPlainText();
                        }
                        if ($this->replaceValue($value)) {
                            $cell->setValue($this->replaceValue($value));
                        }
                    }
                }
            }
        }

        if (empty($this->type)) {
            $this->type = $this->getType($outputFilePath);
        }

        $xlsWriter = PHPExcel_IOFactory::createWriter($this->xlsx, $this->type);
        $xlsWriter->save($outputFilePath);
        if (!file_exists($outputFilePath)) {
            throw new Exception();
        }

        return true;
    }

    /**
     * output
     * Output xls with header.
     */
    public function output($filename = 'output.xlsx', $data = [])
    {
        $outputFilePath = TMP.uniqid('xls_', true).$filename;
        $this->write($outputFilePath, $data);
        header('Pragma: public');
        header('Expires: 0');
        header('Cache-Control: must-revalidate, post-check=0, pre-check=0');
        header('Content-Type: application/force-download');
        header('Content-Type: application/octet-stream');
        header('Content-Type: application/download');
        header("Content-Disposition: attachment; filename=\"$filename\"");
        header('Content-Transfer-Encoding: binary ');
        ob_clean();
        flush();
        echo file_get_contents($outputFilePath);
        exit;
    }

    /**
     * setValue.
     */
    public function setValue($value, $option = ['col' => 'A',
                                                'row' => '1',
                                                'sheet' => 0, ])
    {
        if (!$this->checkKeysFromOptions(['col', 'row'], $option)) {
            return false;
        }
        if (!array_key_exists('sheet', $option)) {
            $option['sheet'] = 0;
        }
        if (empty($this->xlsx)) {
            $this->xlsx = new PHPExcel();
        }
        $this->xlsx->setActiveSheetIndex($option['sheet']);
        $sheet = $this->xlsx->getActiveSheet();
        $sheet->setCellValueByColumnAndRow(self::alphabetToNumber($option['col']), $option['row'], $value);

        // border
        if (array_key_exists('border', $option)) {
            $this->setBorder($option);
        }

        // align horizontal
        if (array_key_exists('align', $option)) {
            $this->setAlign($option);
        }
        if (array_key_exists('hAlign', $option)) {
            $this->setHorizontal($option);
        }

        // align vertical
        if (array_key_exists('vAlign', $option)) {
            $this->setVertical($option);
        }

        // font
        if (array_key_exists('font', $option)) {
            $this->setFont($option);
        }

        // font color
        if (array_key_exists('color', $option)) {
            $this->setFontColoer($option);
        }

        // font size
        if (array_key_exists('size', $option)) {
            $this->setFontSize($option);
        }

        // backgroundColor / backgroundType
        if (array_key_exists('backgroundColor', $option)) {
            $type = empty($option['backgroundType']) ? PHPExcel_Style_Fill::FILL_SOLID : $option['backgroundType'];
            if (strlen($option['backgroundColor']) === 8) {
                $sheet->getStyleByColumnAndRow(self::alphabetToNumber($option['col']), $option['row'])
                    ->getFill()
                    ->setFillType($type);
                $sheet->getStyleByColumnAndRow(self::alphabetToNumber($option['col']), $option['row'])->getFill()
                    ->getStartColor()
                    ->setARGB($option['backgroundColor']);
            } elseif (strlen($option['backgroundColor']) === 6) {
                $sheet->getStyleByColumnAndRow(self::alphabetToNumber($option['col']), $option['row'])
                    ->getFill()
                    ->setFillType($type);
                $sheet->getStyleByColumnAndRow(self::alphabetToNumber($option['col']), $option['row'])
                    ->getFill()
                    ->getStartColor()
                    ->setRGB($option['backgroundColor']);
            }
        }

        return $this;
    }

    /**
     * mergeCells.
     */
    public function mergeCells($option = ['col' => 'A',
                                          'row' => '1',
                                          'toCol' => 'A',
                                          'toRow' => '1',
                                          'sheet' => 0, ])
    {
        if (!$this->checkKeysFromOptions(['fromCol', 'fromRow', 'toCol', 'toRow'], $option)) {
            return false;
        }
        if (!array_key_exists('sheet', $option)) {
            $option['sheet'] = 0;
        }
        if (empty($this->xlsx)) {
            $this->xlsx = new PHPExcel();
        }
        $this->xlsx->setActiveSheetIndex($option['sheet']);
        $sheet = $this->xlsx->getActiveSheet();

        // mergeCell
        $cell = $option['fromCol'].$option['fromRow'];
        $cell .= ':'.$option['toCol'].$option['toRow'];
        $sheet->mergeCells($cell);

        // border
        if (array_key_exists('border', $option)) {
            $this->setBorder($option);
        }

        // align horizontal
        if (array_key_exists('align', $option)) {
            $this->setAlign($option);
        }
        if (array_key_exists('hAlign', $option)) {
            $this->setHorizontal($option);
        }

        // align vertical
        if (array_key_exists('vAlign', $option)) {
            $this->setVertical($option);
        }

        // font
        if (array_key_exists('font', $option)) {
            $this->setFont($option);
        }

        // font color
        if (array_key_exists('color', $option)) {
            $this->setFontColoer($option);
        }

        // font size
        if (array_key_exists('size', $option)) {
            $this->setFontSize($option);
        }

        return $this;
    }

    /**
     * checkKeysFromOption.
     */
    private function checkKeysFromOption($keys, $option)
    {
        foreach ($keys as $key) {
            if (!array_key_exists($key, $option)) {
                return false;
            }
        }

        return true;
    }

    /**
     * setBorder.
     */
    private function setBorder($option)
    {
        $sheet = $this->xlsx->getActiveSheet();

        if (is_array($option['border'])) {
            foreach (['top', 'right', 'left', 'bottom'] as $position) {
                if (array_key_exists($position, $option['border'])) {
                    $sheet->getStyleByColumnAndRow(self::alphabetToNumber($option['col']), $option['row'])
                        ->getBorders()
                        ->{'get'.ucfirst($position)}()
                        ->setBorderStyle($option['border'][$position]);
                }
            }
        } else {
            $sheet->getStyleByColumnAndRow(self::alphabetToNumber($option['col']), $option['row'])
                ->getBorders()
                ->getAllBorders()
                ->setBorderStyle($option['border']);
        }
    }

    /**
     * setAlign.
     */
    private function setAlign($option)
    {
        $sheet = $this->xlsx->getActiveSheet();

        $sheet->getStyleByColumnAndRow(self::alphabetToNumber($option['col']), $option['row'])
                ->getAlignment()
                ->setHorizontal($option['align']);
    }

    /**
     * setHorizontal.
     */
    private function setHorizontal($option)
    {
        $sheet = $this->xlsx->getActiveSheet();

        $sheet->getStyleByColumnAndRow(self::alphabetToNumber($option['col']), $option['row'])
                ->getAlignment()
                ->setHorizontal($option['hAlign']);
    }

    /**
     * setVertical.
     */
    private function setVertical($option)
    {
        $sheet = $this->xlsx->getActiveSheet();

        $sheet->getStyleByColumnAndRow(self::alphabetToNumber($option['col']), $option['row'])
                ->getAlignment()
                ->setVertical($option['vAlign']);
    }

    /**
     * setFont.
     */
    private function setFont($option)
    {
        $sheet = $this->xlsx->getActiveSheet();

        $sheet->getStyleByColumnAndRow(self::alphabetToNumber($option['col']), $option['row'])
                ->getFont()
                ->setName($option['font']);
    }

    /**
     * setFontColoer.
     */
    private function setFontColoer($option)
    {
        $sheet = $this->xlsx->getActiveSheet();

        if (strlen($option['color']) === 8) {
            $sheet->getStyleByColumnAndRow(self::alphabetToNumber($option['col']), $option['row'])
                ->getFont()
                ->getColor()
                ->setARGB($option['color']);
        } elseif (strlen($option['color']) === 6) {
            $sheet->getStyleByColumnAndRow(self::alphabetToNumber($option['col']), $option['row'])
                ->getFont()
                ->getColor()
                ->setRGB($option['color']);
        }
    }

    /**
     * setFontSize.
     */
    private function setFontSize($option)
    {
        $sheet = $this->xlsx->getActiveSheet();

        $sheet->getStyleByColumnAndRow(self::alphabetToNumber($option['col']), $option['row'])
                ->getFont()
                ->setSize($option['size']);
    }

    /**
     * alphabetToNumber.
     */
    public static function alphabetToNumber($value)
    {
        if (is_numeric($value)) {
            return $value;
        }
        $alphabet = array_flip(str_split('abcdefghijklmnopqrstuvwxyz'));
        $strArray = array_reverse(str_split(strtolower($value)));
        $number = 0;
        foreach ($strArray as $n => $str) {
            if ($n == 0) {
                $number += $alphabet[$str];
            } else {
                $number += ($alphabet[$str] + 1) * pow(26, $n);
            }
        }

        return $number;
    }

    /**
     * getType.
     *
     * @see IOFactory::createReaderForFile
     */
    private function getType($filePath)
    {
        $pathinfo = pathinfo($filePath);

        if (!isset($pathinfo['extension'])) {
            return false;
        }
        switch (strtolower($pathinfo['extension'])) {
        case 'xlsx':            //  Excel (OfficeOpenXML) Spreadsheet
        case 'xlsm':            //  Excel (OfficeOpenXML) Macro Spreadsheet (macros will be discarded)
        case 'xltx':            //  Excel (OfficeOpenXML) Template
        case 'xltm':            //  Excel (OfficeOpenXML) Macro Template (macros will be discarded)
            $extensionType = 'Excel2007';
            break;
        case 'xls':             //  Excel (BIFF) Spreadsheet
        case 'xlt':             //  Excel (BIFF) Template
            $extensionType = 'Excel5';
            break;
        case 'ods':             //  Open/Libre Offic Calc
        case 'ots':             //  Open/Libre Offic Calc Template
            $extensionType = 'OOCalc';
            break;
        case 'slk':
            $extensionType = 'SYLK';
            break;
        case 'xml':             //  Excel 2003 SpreadSheetML
            $extensionType = 'Excel2003XML';
            break;
        case 'gnumeric':
            $extensionType = 'Gnumeric';
            break;
        case 'htm':
        case 'html':
            $extensionType = 'HTML';
            break;
        case 'csv':
            // Do nothing
            // We must not try to use CSV reader since it loads
            // all files including Excel files etc.
            return false;
            break;
        default:
            break;
        }

        return $extensionType;
    }

    /**
     * replaceValue.
     */
    private function replaceValue($value)
    {
        if (empty($value)) {
            return false;
        }
        if (array_key_exists((string) $value, $this->data)) {
            return $this->data[$value];
        } else {
            return false;
        }
    }
}

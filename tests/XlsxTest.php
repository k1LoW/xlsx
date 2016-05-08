<?php

namespace Xlsx\tests;

use Xlsx\Xlsx;
use PHPExcel_Style_Border;
use PHPExcel_Style_Color;

class XlsxTest extends \PHPUnit_Framework_TestCase
{
    public $xlsx;

    /**
     * setUp.
     */
    public function setUp()
    {
    }

    /**
     * tearDown.
     */
    public function tearDown()
    {
        unset($this->xlsx);
    }

    /**
     * testWriteExcel2007.
     */
    public function testWriteExcel2007()
    {
        $fileName = 'testbook.xlsx';
        $this->inputFilePath = '/tmp/'.$fileName;
        $this->outputFilePath = '/tmp/'.'output.xlsx';
        $this->_setTestFile($fileName, $this->inputFilePath);

        $this->xlsx = new Xlsx();
        $result = $this->xlsx->read($this->inputFilePath)
                ->setValue('testset', [
                    'col' => 'B', // jpn: col / row / sheetを指定してセット可能
                    'row' => '10',
                ])
                ->setValue('testset_with_sheet', [
                    'col' => 'B', // jpn: col / row / sheetを指定してセット可能
                    'row' => '10',
                    'sheet' => 2,
                ])
                ->setValue('testset_with_border', [
                    'col' => 'C', // jpn: borderを指定してセット可能
                    'row' => '10',
                    'border' => [
                        'top' => PHPExcel_Style_Border::BORDER_THICK,
                        'right' => PHPExcel_Style_Border::BORDER_MEDIUM,
                        'left' => PHPExcel_Style_Border::BORDER_THIN,
                        'bottom' => PHPExcel_Style_Border::BORDER_DOUBLE,
                    ],
                ])
                ->setValue('testset_with_border', [
                    'col' => 'E',
                    'row' => '10',
                    'border' => PHPExcel_Style_Border::BORDER_THICK,  // jpn: borderを一括指定可能
                ])
                ->setValue('testset_with_color', [
                    'col' => 'F', // jpn: colorを指定可能
                    'row' => '10',
                    'color' => PHPExcel_Style_Color::COLOR_BLUE,
                ])
                ->setValue('testset_with_backgroud_color', [
                    'col' => 'G', // jpn: backgroundColorを指定可能
                    'row' => '10',
                    'backgroundColor' => PHPExcel_Style_Color::COLOR_YELLOW,
                ])
                ->set([
                    'Sheet1' => 'シートタイトル', // jpn: 文字列置換でセット可能
                    'test' => 'replaced',
                    '4' => 5,
                ])
                ->write($this->outputFilePath);
        $this->assertTrue($result);
        echo PHP_EOL.'Look '.$this->outputFilePath.PHP_EOL;
        echo 'Peak memory usage: '.(memory_get_peak_usage(true) / 1024 / 1024).' MB'.PHP_EOL;
    }

    /**
     * testWriteExcel5.
     */
    public function testWriteExcel5()
    {
        $fileName = 'testbook.xls';
        $this->inputFilePath = '/tmp/'.$fileName;
        $this->outputFilePath = '/tmp/'.'output.xls';
        $this->_setTestFile($fileName, $this->inputFilePath);

        $this->xlsx = new Xlsx($this->inputFilePath);
        // jpn: col / row / sheetを指定してセット可能
        $this->xlsx
            ->setValue('testset', [
                'col' => 'B', // jpn: col / row / sheetを指定してセット可能
                'row' => '10',
            ])
            ->setValue('testset_with_sheet', [
                'col' => 'B', // jpn: col / row / sheetを指定してセット可能
                'row' => '10',
                'sheet' => 2,
            ])
            ->setValue('testset_with_border', [
                'col' => 'C', // jpn: borderを指定可能
                'row' => '10',
                'border' => [
                    'top' => PHPExcel_Style_Border::BORDER_THICK,
                    'right' => PHPExcel_Style_Border::BORDER_MEDIUM,
                    'left' => PHPExcel_Style_Border::BORDER_THIN,
                    'bottom' => PHPExcel_Style_Border::BORDER_DOUBLE,
                ],
            ])
            ->setValue('testset_with_border', [
                'col' => 'E',
                'row' => '10',
                'border' => PHPExcel_Style_Border::BORDER_THICK,  // jpn: borderを一括指定可能
            ])
            ->setValue('testset_with_color', [
                'col' => 'F', // jpn: colorを指定可能
                'row' => '10',
                'color' => PHPExcel_Style_Color::COLOR_BLUE,
            ])
            ->setValue('testset_with_backgroud_color', [
                'col' => 'G', // jpn: backgroundColorを指定可能
                'row' => '10',
                'backgroundColor' => PHPExcel_Style_Color::COLOR_YELLOW,
            ])
            ->set([
                'Sheet1' => 'シートタイトル',         // jpn: 文字列置換でセット可能
                'test' => 'replaced',
                '4' => 5,
            ]);
        $result = $this->xlsx->write($this->outputFilePath);
        $this->assertTrue($result);
        echo PHP_EOL.'Look '.$this->outputFilePath.PHP_EOL;
        echo 'Peak memory usage: '.(memory_get_peak_usage(true) / 1024 / 1024).' MB'.PHP_EOL;
    }

    /**
     * testNewExcel2007.
     */
    public function testNewExcel2007()
    {
        $this->outputFilePath = '/tmp/outputnew.xlsx';

        $this->xlsx = new Xlsx();
        $result = $this->xlsx
                ->setValue('testset', [
                    'col' => 'B', // jpn: col / row / sheetを指定してセット可能
                    'row' => '10',
                ])
                ->write($this->outputFilePath);
        $this->assertTrue($result);
        echo PHP_EOL.'Look '.$this->outputFilePath.PHP_EOL;
        echo 'Peak memory usage: '.(memory_get_peak_usage(true) / 1024 / 1024).' MB'.PHP_EOL;
    }

    /**
     * testNewExcel5.
     */
    public function testNewExcel5()
    {
        $this->outputFilePath = '/tmp/outputnew.xls';

        $this->xlsx = new Xlsx();
        $result = $this->xlsx
                ->setValue('testset', [
                    'col' => 'B', // jpn: col / row / sheetを指定してセット可能
                    'row' => '10',
                ])
                ->write($this->outputFilePath);
        $this->assertTrue($result);
        echo PHP_EOL.'Look '.$this->outputFilePath.PHP_EOL;
        echo 'Peak memory usage: '.(memory_get_peak_usage(true) / 1024 / 1024).' MB'.PHP_EOL;
    }

    /**
     * testSetValueManyExcel2007.
     *
     * @param
     */
    public function testSetValueManyExcel2007()
    {
        $fileName = 'testbook.xlsx';
        $this->inputFilePath = '/tmp/'.$fileName;
        $this->outputFilePath = '/tmp/'.'outputmany.xlsx';
        $this->_setTestFile($fileName, $this->inputFilePath);

        $this->xlsx = new Xlsx($this->inputFilePath);

        for ($c = 0; $c < 100; ++$c) {
            for ($r = 1; $r < 100; ++$r) {
                $this->xlsx
                    ->setValue('testset', [
                        'col' => $c,
                        'row' => $r,
                    ]);
            }
        }
        $result = $this->xlsx->write($this->outputFilePath);
        $this->assertTrue($result);
        echo PHP_EOL.'Look '.$this->outputFilePath.PHP_EOL;
        echo 'Peak memory usage: '.(memory_get_peak_usage(true) / 1024 / 1024).' MB'.PHP_EOL;
    }

    /**
     * testSetValueManyExcel5.
     *
     * @param
     */
    public function testSetValueManyExcel5()
    {
        $fileName = 'testbook.xls';
        $this->inputFilePath = '/tmp/'.$fileName;
        $this->outputFilePath = '/tmp/'.'outputmany.xls';
        $this->_setTestFile($fileName, $this->inputFilePath);

        $this->xlsx = new Xlsx($this->inputFilePath);

        for ($c = 0; $c < 100; ++$c) {
            for ($r = 1; $r < 100; ++$r) {
                $this->xlsx
                    ->setValue('testset', [
                        'col' => $c,
                        'row' => $r,
                    ]);
            }
        }
        $result = $this->xlsx->write($this->outputFilePath);
        $this->assertTrue($result);
        echo PHP_EOL.'Look '.$this->outputFilePath.PHP_EOL;
        echo 'Peak memory usage: '.(memory_get_peak_usage(true) / 1024 / 1024).' MB'.PHP_EOL;
    }

    /**
     * testSetManyExcel2007.
     */
    public function testSetManyExcel2007()
    {
        $fileName = 'testbookmany.xlsx';
        $this->inputFilePath = '/tmp/'.$fileName;
        $this->outputFilePath = '/tmp/'.'outputsetmany.xlsx';
        $this->_setTestFile($fileName, $this->inputFilePath);

        $this->xlsx = new Xlsx($this->inputFilePath);
        for ($c = 'a'; Xlsx::alphabetToNumber($c) < 100; ++$c) {
            for ($r = 1; $r < 100; ++$r) {
                $this->xlsx->set('$'.strtoupper($c).'$'.$r, 'replaced');
            }
        }

        $result = $this->xlsx->write($this->outputFilePath);
        $this->assertTrue($result);
        echo PHP_EOL.'Look '.$this->outputFilePath.PHP_EOL;
        echo 'Peak memory usage: '.(memory_get_peak_usage(true) / 1024 / 1024).' MB'.PHP_EOL;
    }

    /**
     * testSetManyExcel5.
     */
    public function testSetManyExcel5()
    {
        $fileName = 'testbookmany.xls';
        $this->inputFilePath = '/tmp/'.$fileName;
        $this->outputFilePath = '/tmp/'.'outputsetmany.xls';
        $this->_setTestFile($fileName, $this->inputFilePath);
        $this->xlsx = new Xlsx($this->inputFilePath);
        for ($c = 'a'; Xlsx::alphabetToNumber($c) < 100; ++$c) {
            for ($r = 1; $r < 100; ++$r) {
                $this->xlsx->set('$'.strtoupper($c).'$'.$r, 'replaced');
            }
        }

        $result = $this->xlsx->write($this->outputFilePath);
        $this->assertTrue($result);
        echo PHP_EOL.'Look '.$this->outputFilePath.PHP_EOL;
        echo 'Peak memory usage: '.(memory_get_peak_usage(true) / 1024 / 1024).' MB'.PHP_EOL;
    }

    /**
     * testAlphabetToNumber.
     *
     * @param $arg
     */
    public function testAlphabetToNumber()
    {
        $result = Xlsx::alphabetToNumber('A');
        $this->assertEquals($result, 0);
        $result = Xlsx::alphabetToNumber('AA');
        $this->assertEquals($result, 26);
        $result = Xlsx::alphabetToNumber('AM');
        $this->assertEquals($result, 38);
        $result = Xlsx::alphabetToNumber('AZ');
        $this->assertEquals($result, 51);
        $result = Xlsx::alphabetToNumber('BA');
        $this->assertEquals($result, 52);
        $result = Xlsx::alphabetToNumber('DL');
        $this->assertEquals($result, 115);
        $result = Xlsx::alphabetToNumber('ZZ');
        $this->assertEquals($result, 701);
        $result = Xlsx::alphabetToNumber('AAA');
        $this->assertEquals($result, 702);
    }

    /**
     * _setTestFile.
     *
     * @return
     */
    private function _setTestFile($fileName, $to = null)
    {
        if (!$fileName || !$to) {
            return false;
        }
        $from = dirname(__FILE__).'/'.$fileName;

        return copy($from, $to);
    }
}

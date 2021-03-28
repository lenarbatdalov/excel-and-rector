<?php

/** PHPExcel root directory */
if (!defined('PHPEXCEL_ROOT')) {
    /**
     * @ignore
     */
    define('PHPEXCEL_ROOT', dirname(__FILE__) . '/../../');
    require(PHPEXCEL_ROOT . 'PHPExcel/Autoloader.php');
}

namespace PhpOffice\PhpSpreadsheet\Reader;

/**
 * PHPExcel_Reader_SYLK
 *
 * Copyright (c) 2006 - 2015 PHPExcel
 *
 * This library is free software; you can redistribute it and/or
 * modify it under the terms of the GNU Lesser General Public
 * License as published by the Free Software Foundation; either
 * version 2.1 of the License, or (at your option) any later version.
 *
 * This library is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
 * Lesser General Public License for more details.
 *
 * You should have received a copy of the GNU Lesser General Public
 * License along with this library; if not, write to the Free Software
 * Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston, MA  02110-1301  USA
 *
 * @category   PHPExcel
 * @package    PHPExcel_Reader
 * @copyright  Copyright (c) 2006 - 2015 PHPExcel (http://www.codeplex.com/PHPExcel)
 * @license    http://www.gnu.org/licenses/old-licenses/lgpl-2.1.txt    LGPL
 * @version    ##VERSION##, ##DATE##
 */
class Slk extends \PhpOffice\PhpSpreadsheet\Reader\BaseReader implements \PhpOffice\PhpSpreadsheet\Reader\IReader
{
    /**
     * Input encoding
     *
     * @var string
     */
    private $inputEncoding = 'ANSI';

    /**
     * Sheet index to read
     *
     * @var int
     */
    private $sheetIndex = 0;

    /**
     * Formats
     *
     * @var array
     */
    private $formats = array();

    /**
     * Create a new PHPExcel_Reader_SYLK
     */
    public function __construct()
    {
        $this->readFilter = new \PhpOffice\PhpSpreadsheet\Reader\DefaultReadFilter();
    }

    /**
     * Validate that the current file is a SYLK file
     *
     * @return boolean
     */
    protected function isValidFormat()
    {
        // Read sample data (first 2 KB will do)
        $data = \fread($this->fileHandle, 2048);

        // Count delimiters in file
        $delimiterCount = \substr_count($data, ';');
        if ($delimiterCount < 1) {
            return \false;
        }

        // Analyze first line looking for ID; signature
        $lines = \explode("\n", $data);
        if (\substr($lines[0], 0, 4) != 'ID;P') {
            return \false;
        }

        return \true;
    }

    /**
     * Set input encoding
     *
     * @param string $pValue Input encoding
     */
    public function setInputEncoding($pValue = 'ANSI')
    {
        $this->inputEncoding = $pValue;
        return $this;
    }

    /**
     * Get input encoding
     *
     * @return string
     */
    public function getInputEncoding()
    {
        return $this->inputEncoding;
    }

    /**
     * Return worksheet info (Name, Last Column Letter, Last Column Index, Total Rows, Total Columns)
     *
     * @param   string     $pFilename
     * @throws   \PhpOffice\PhpSpreadsheet\Reader\Exception
     */
    public function listWorksheetInfo($pFilename)
    {
        // Open file
        $this->openFile($pFilename);
        if (!$this->isValidFormat()) {
            \fclose($this->fileHandle);
            throw new \PhpOffice\PhpSpreadsheet\Reader\Exception($pFilename . " is an Invalid Spreadsheet file.");
        }
        $fileHandle = $this->fileHandle;
        \rewind($fileHandle);

        $worksheetInfo = array();
        $worksheetInfo[0]['worksheetName'] = 'Worksheet';
        $worksheetInfo[0]['lastColumnLetter'] = 'A';
        $worksheetInfo[0]['lastColumnIndex'] = 0;
        $worksheetInfo[0]['totalRows'] = 0;
        $worksheetInfo[0]['totalColumns'] = 0;

        // Loop through file
        $rowData = array();

        // loop through one row (line) at a time in the file
        $rowIndex = 0;
        while (($rowData = \fgets($fileHandle)) !== \false) {
            $columnIndex = 0;

            // convert SYLK encoded $rowData to UTF-8
            $rowData = \PhpOffice\PhpSpreadsheet\Shared\StringHelper::SYLKtoUTF8($rowData);

            // explode each row at semicolons while taking into account that literal semicolon (;)
            // is escaped like this (;;)
            $rowData = \explode("\t", \str_replace('¤', ';', \str_replace(';', "\t", \str_replace(';;', '¤', \rtrim($rowData)))));

            $dataType = \array_shift($rowData);
            if ($dataType == 'C') {
                //  Read cell value data
                foreach ($rowData as $singleRowData) {
                    switch ($singleRowData{0}) {
                        case 'C':
                        case 'X':
                            $columnIndex = \substr($singleRowData, 1) - 1;
                            break;
                        case 'R':
                        case 'Y':
                            $rowIndex = \substr($singleRowData, 1);
                            break;
                    }

                    $worksheetInfo[0]['totalRows'] = \max($worksheetInfo[0]['totalRows'], $rowIndex);
                    $worksheetInfo[0]['lastColumnIndex'] = \max($worksheetInfo[0]['lastColumnIndex'], $columnIndex);
                }
            }
        }

        $worksheetInfo[0]['lastColumnLetter'] = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::stringFromColumnIndex($worksheetInfo[0]['lastColumnIndex']);
        $worksheetInfo[0]['totalColumns'] = $worksheetInfo[0]['lastColumnIndex'] + 1;

        // Close file
        \fclose($fileHandle);

        return $worksheetInfo;
    }

    /**
     * Loads PHPExcel from file
     *
     * @param     string         $pFilename
     * @return     \PhpOffice\PhpSpreadsheet\Spreadsheet
     * @throws     \PhpOffice\PhpSpreadsheet\Reader\Exception
     */
    public function load($pFilename)
    {
        // Create new PHPExcel
        $phpExcel = new \PhpOffice\PhpSpreadsheet\Spreadsheet();

        // Load into this instance
        return $this->loadIntoExisting($pFilename, $phpExcel);
    }

    /**
     * Loads PHPExcel from file into PHPExcel instance
     *
     * @param     string         $pFilename
     * @return     \PhpOffice\PhpSpreadsheet\Spreadsheet
     * @throws     \PhpOffice\PhpSpreadsheet\Reader\Exception
     */
    public function loadIntoExisting($pFilename, \PhpOffice\PhpSpreadsheet\Spreadsheet $phpExcel)
    {
        // Open file
        $this->openFile($pFilename);
        if (!$this->isValidFormat()) {
            \fclose($this->fileHandle);
            throw new \PhpOffice\PhpSpreadsheet\Reader\Exception($pFilename . " is an Invalid Spreadsheet file.");
        }
        $fileHandle = $this->fileHandle;
        \rewind($fileHandle);

        // Create new PHPExcel
        while ($phpExcel->getSheetCount() <= $this->sheetIndex) {
            $phpExcel->createSheet();
        }
        $phpExcel->setActiveSheetIndex($this->sheetIndex);

        $fromFormats    = array('\-',    '\ ');
        $toFormats        = array('-',    ' ');

        // Loop through file
        $rowData = array();
        $column = $row = '';

        // loop through one row (line) at a time in the file
        while (($rowData = \fgets($fileHandle)) !== \false) {
            // convert SYLK encoded $rowData to UTF-8
            $rowData = \PhpOffice\PhpSpreadsheet\Shared\StringHelper::SYLKtoUTF8($rowData);

            // explode each row at semicolons while taking into account that literal semicolon (;)
            // is escaped like this (;;)
            $rowData = \explode("\t", \str_replace('¤', ';', \str_replace(';', "\t", \str_replace(';;', '¤', \rtrim($rowData)))));

            $dataType = \array_shift($rowData);
            //    Read shared styles
            if ($dataType == 'P') {
                $formatArray = array();
                foreach ($rowData as $singleRowData) {
                    switch ($singleRowData{0}) {
                        case 'P':
                            $formatArray['numberformat']['code'] = \str_replace($fromFormats, $toFormats, \substr($singleRowData, 1));
                            break;
                        case 'E':
                        case 'F':
                            $formatArray['font']['name'] = \substr($singleRowData, 1);
                            break;
                        case 'L':
                            $formatArray['font']['size'] = \substr($singleRowData, 1);
                            break;
                        case 'S':
                            $styleSettings = \substr($singleRowData, 1);
                            for ($i=0; $i<\strlen($styleSettings); ++$i) {
                                switch ($styleSettings{$i}) {
                                    case 'I':
                                        $formatArray['font']['italic'] = \true;
                                        break;
                                    case 'D':
                                        $formatArray['font']['bold'] = \true;
                                        break;
                                    case 'T':
                                        $formatArray['borders']['top']['style'] = \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN;
                                        break;
                                    case 'B':
                                        $formatArray['borders']['bottom']['style'] = \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN;
                                        break;
                                    case 'L':
                                        $formatArray['borders']['left']['style'] = \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN;
                                        break;
                                    case 'R':
                                        $formatArray['borders']['right']['style'] = \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN;
                                        break;
                                }
                            }
                            break;
                    }
                }
                $this->format++;
            //    Read cell value data
            } elseif ($dataType == 'C') {
                $hasCalculatedValue = \false;
                $cellData = $cellDataFormula = '';
                foreach ($rowData as $singleRowData) {
                    switch ($singleRowData{0}) {
                        case 'C':
                        case 'X':
                            $column = \substr($singleRowData, 1);
                            break;
                        case 'R':
                        case 'Y':
                            $row = \substr($singleRowData, 1);
                            break;
                        case 'K':
                            $cellData = \substr($singleRowData, 1);
                            break;
                        case 'E':
                            $cellDataFormula = '='.\substr($singleRowData, 1);
                            //    Convert R1C1 style references to A1 style references (but only when not quoted)
                            $temp = \explode('"', $cellDataFormula);
                            $key = \false;
                            foreach ($temp as &$singleTemp) {
                                //    Only count/replace in alternate array entries
                                if ($key = !$key) {
                                    \preg_match_all('/(R(\[?-?\d*\]?))(C(\[?-?\d*\]?))/', $singleTemp, $cellReferences, \PREG_SET_ORDER+\PREG_OFFSET_CAPTURE);
                                    //    Reverse the matches array, otherwise all our offsets will become incorrect if we modify our way
                                    //        through the formula from left to right. Reversing means that we work right to left.through
                                    //        the formula
                                    $cellReferences = \array_reverse($cellReferences);
                                    //    Loop through each R1C1 style reference in turn, converting it to its A1 style equivalent,
                                    //        then modify the formula to use that new reference
                                    foreach ($cellReferences as $cellReference) {
                                        $rowReference = $cellReference[2][0];
                                        //    Empty R reference is the current row
                                        if ($rowReference == '') {
                                            $rowReference = $row;
                                        }
                                        //    Bracketed R references are relative to the current row
                                        if ($rowReference{0} == '[') {
                                            $rowReference = $row + \trim($rowReference, '[]');
                                        }
                                        $columnReference = $cellReference[4][0];
                                        //    Empty C reference is the current column
                                        if ($columnReference == '') {
                                            $columnReference = $column;
                                        }
                                        //    Bracketed C references are relative to the current column
                                        if ($columnReference{0} == '[') {
                                            $columnReference = $column + \trim($columnReference, '[]');
                                        }
                                        $A1CellReference = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::stringFromColumnIndex($columnReference-1).$rowReference;

                                        $singleTemp = \substr_replace($singleTemp, $A1CellReference, $cellReference[0][1], \strlen($cellReference[0][0]));
                                    }
                                }
                            }
                            unset($value);
                            //    Then rebuild the formula string
                            $cellDataFormula = \implode('"', $temp);
                            $hasCalculatedValue = \true;
                            break;
                    }
                }
                $columnLetter = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::stringFromColumnIndex($column-1);
                $cellData = \PhpOffice\PhpSpreadsheet\Calculation\Calculation::unwrapResult($cellData);

                // Set cell value
                $phpExcel->getActiveSheet()->getCell($columnLetter.$row, true)->setValue(($hasCalculatedValue) ? $cellDataFormula : $cellData);
                if ($hasCalculatedValue) {
                    $cellData = \PhpOffice\PhpSpreadsheet\Calculation\Calculation::unwrapResult($cellData);
                    $phpExcel->getActiveSheet()->getCell($columnLetter.$row, true)->setCalculatedValue($cellData);
                }
            //    Read cell formatting
            } elseif ($dataType == 'F') {
                $formatStyle = $columnWidth = $styleSettings = '';
                $styleData = array();
                foreach ($rowData as $singleRowData) {
                    switch ($singleRowData{0}) {
                        case 'C':
                        case 'X':
                            $column = \substr($singleRowData, 1);
                            break;
                        case 'R':
                        case 'Y':
                            $row = \substr($singleRowData, 1);
                            break;
                        case 'P':
                            $formatStyle = $singleRowData;
                            break;
                        case 'W':
                            list($startCol, $endCol, $columnWidth) = \explode(' ', \substr($singleRowData, 1));
                            break;
                        case 'S':
                            $styleSettings = \substr($singleRowData, 1);
                            for ($i=0; $i<\strlen($styleSettings); ++$i) {
                                switch ($styleSettings{$i}) {
                                    case 'I':
                                        $styleData['font']['italic'] = \true;
                                        break;
                                    case 'D':
                                        $styleData['font']['bold'] = \true;
                                        break;
                                    case 'T':
                                        $styleData['borders']['top']['style'] = \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN;
                                        break;
                                    case 'B':
                                        $styleData['borders']['bottom']['style'] = \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN;
                                        break;
                                    case 'L':
                                        $styleData['borders']['left']['style'] = \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN;
                                        break;
                                    case 'R':
                                        $styleData['borders']['right']['style'] = \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN;
                                        break;
                                }
                            }
                            break;
                    }
                }
                if (($formatStyle > '') && ($column > '') && ($row > '')) {
                    $columnLetter = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::stringFromColumnIndex($column-1);
                    if (isset($this->formats[$formatStyle])) {
                        $phpExcel->getActiveSheet()->getStyle($columnLetter.$row)->applyFromArray($this->formats[$formatStyle], true);
                    }
                }
                if ((!empty($styleData)) && ($column > '') && ($row > '')) {
                    $columnLetter = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::stringFromColumnIndex($column-1);
                    $phpExcel->getActiveSheet()->getStyle($columnLetter.$row)->applyFromArray($styleData, true);
                }
                if ($columnWidth > '') {
                    if ($startCol === $endCol) {
                        $startCol = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::stringFromColumnIndex($startCol-1);
                        $phpExcel->getActiveSheet()->getColumnDimension($startCol, true)->setWidth($columnWidth);
                    } else {
                        $startCol = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::stringFromColumnIndex($startCol-1);
                        $endCol = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::stringFromColumnIndex($endCol-1);
                        $phpExcel->getActiveSheet()->getColumnDimension($startCol, true)->setWidth($columnWidth);
                        do {
                            $phpExcel->getActiveSheet()->getColumnDimension(++$startCol, true)->setWidth($columnWidth);
                        } while ($startCol !== $endCol);
                    }
                }
            } else {
                foreach ($rowData as $singleRowData) {
                    switch ($singleRowData{0}) {
                        case 'C':
                        case 'X':
                            $column = \substr($singleRowData, 1);
                            break;
                        case 'R':
                        case 'Y':
                            $row = \substr($singleRowData, 1);
                            break;
                    }
                }
            }
        }

        // Close file
        \fclose($fileHandle);

        // Return
        return $phpExcel;
    }

    /**
     * Get sheet index
     *
     * @return int
     */
    public function getSheetIndex()
    {
        return $this->sheetIndex;
    }

    /**
     * Set sheet index
     *
     * @param    int        $pValue        Sheet index
     * @return \PhpOffice\PhpSpreadsheet\Reader\Slk
     */
    public function setSheetIndex($pValue = 0)
    {
        $this->sheetIndex = $pValue;
        return $this;
    }
}

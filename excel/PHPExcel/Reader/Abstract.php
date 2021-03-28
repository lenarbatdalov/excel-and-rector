<?php

namespace PhpOffice\PhpSpreadsheet\Reader;

/**
 * PHPExcel_Reader_Abstract
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
abstract class BaseReader implements \PhpOffice\PhpSpreadsheet\Reader\IReader
{
    /**
     * Read data only?
     * Identifies whether the Reader should only read data values for cells, and ignore any formatting information;
     *        or whether it should read both data and formatting
     *
     * @var    boolean
     */
    protected $readDataOnly = \false;

    /**
     * Read empty cells?
     * Identifies whether the Reader should read data values for cells all cells, or should ignore cells containing
     *         null value or empty string
     *
     * @var    boolean
     */
    protected $readEmptyCells = \true;

    /**
     * Read charts that are defined in the workbook?
     * Identifies whether the Reader should read the definitions for any charts that exist in the workbook;
     *
     * @var    boolean
     */
    protected $includeCharts = \false;

    /**
     * Restrict which sheets should be loaded?
     * This property holds an array of worksheet names to be loaded. If null, then all worksheets will be loaded.
     *
     * @var array of string
     */
    protected $loadSheetsOnly;

    /**
     * PHPExcel_Reader_IReadFilter instance
     *
     * @var \PhpOffice\PhpSpreadsheet\Reader\IReadFilter
     */
    protected $readFilter;

    protected $fileHandle;


    /**
     * Read data only?
     *        If this is true, then the Reader will only read data values for cells, it will not read any formatting information.
     *        If false (the default) it will read data and formatting.
     *
     * @return    boolean
     */
    public function isReadDataOnly()
    {
        return $this->readDataOnly;
    }

    /**
     * Set read data only
     *        Set to true, to advise the Reader only to read data values for cells, and to ignore any formatting information.
     *        Set to false (the default) to advise the Reader to read both data and formatting for cells.
     *
     * @param    boolean    $pValue
     *
     * @return    \PhpOffice\PhpSpreadsheet\Reader\IReader
     */
    public function setReadDataOnly($pValue = \false)
    {
        $this->readDataOnly = $pValue;
        return $this;
    }

    /**
     * Read empty cells?
     *        If this is true (the default), then the Reader will read data values for all cells, irrespective of value.
     *        If false it will not read data for cells containing a null value or an empty string.
     *
     * @return    boolean
     */
    public function isReadEmptyCells()
    {
        return $this->readEmptyCells;
    }

    /**
     * Set read empty cells
     *        Set to true (the default) to advise the Reader read data values for all cells, irrespective of value.
     *        Set to false to advise the Reader to ignore cells containing a null value or an empty string.
     *
     * @param    boolean    $pValue
     *
     * @return    \PhpOffice\PhpSpreadsheet\Reader\IReader
     */
    public function setReadEmptyCells($pValue = \true)
    {
        $this->readEmptyCells = $pValue;
        return $this;
    }

    /**
     * Read charts in workbook?
     *        If this is true, then the Reader will include any charts that exist in the workbook.
     *      Note that a ReadDataOnly value of false overrides, and charts won't be read regardless of the IncludeCharts value.
     *        If false (the default) it will ignore any charts defined in the workbook file.
     *
     * @return    boolean
     */
    public function isIncludeCharts()
    {
        return $this->includeCharts;
    }

    /**
     * Set read charts in workbook
     *        Set to true, to advise the Reader to include any charts that exist in the workbook.
     *      Note that a ReadDataOnly value of false overrides, and charts won't be read regardless of the IncludeCharts value.
     *        Set to false (the default) to discard charts.
     *
     * @param    boolean    $pValue
     *
     * @return    \PhpOffice\PhpSpreadsheet\Reader\IReader
     */
    public function setIncludeCharts($pValue = \false)
    {
        $this->includeCharts = $pValue;
        return $this;
    }

    /**
     * Get which sheets to load
     * Returns either an array of worksheet names (the list of worksheets that should be loaded), or a null
     *        indicating that all worksheets in the workbook should be loaded.
     *
     * @return mixed
     */
    public function getLoadSheetsOnly()
    {
        return $this->loadSheetsOnly;
    }

    /**
     * Set which sheets to load
     *
     * @param mixed $value
     *        This should be either an array of worksheet names to be loaded, or a string containing a single worksheet name.
     *        If NULL, then it tells the Reader to read all worksheets in the workbook
     *
     * @return \PhpOffice\PhpSpreadsheet\Reader\IReader
     */
    public function setLoadSheetsOnly($value = \null)
    {
        if ($value === \null) {
            return $this->setLoadAllSheets();
        }

        $this->loadSheetsOnly = \is_array($value) ? $value : array($value);
        return $this;
    }

    /**
     * Set all sheets to load
     *        Tells the Reader to load all worksheets from the workbook.
     *
     * @return \PhpOffice\PhpSpreadsheet\Reader\IReader
     */
    public function setLoadAllSheets()
    {
        $this->loadSheetsOnly = \null;
        return $this;
    }

    /**
     * Read filter
     *
     * @return \PhpOffice\PhpSpreadsheet\Reader\IReadFilter
     */
    public function getReadFilter()
    {
        return $this->readFilter;
    }

    /**
     * Set read filter
     *
     * @return \PhpOffice\PhpSpreadsheet\Reader\IReader
     */
    public function setReadFilter(\PhpOffice\PhpSpreadsheet\Reader\IReadFilter $phpExcelReaderIReadFilter)
    {
        $this->readFilter = $phpExcelReaderIReadFilter;
        return $this;
    }

    /**
     * Open file for reading
     *
     * @param string $pFilename
     * @throws    \PhpOffice\PhpSpreadsheet\Reader\Exception
     * @return resource
     */
    protected function openFile($pFilename)
    {
        // Check if file exists
        if (!\file_exists($pFilename) || !\is_readable($pFilename)) {
            throw new \PhpOffice\PhpSpreadsheet\Reader\Exception("Could not open " . $pFilename . " for reading! File does not exist.");
        }

        // Open file
        $this->fileHandle = \fopen($pFilename, 'r');
        if ($this->fileHandle === \false) {
            throw new \PhpOffice\PhpSpreadsheet\Reader\Exception("Could not open file " . $pFilename . " for reading.");
        }
    }

    /**
     * Can the current PHPExcel_Reader_IReader read the file?
     *
     * @param     string         $pFilename
     * @return boolean
     * @throws \PhpOffice\PhpSpreadsheet\Reader\Exception
     */
    public function canRead($pFilename)
    {
        // Check if file exists
        try {
            $this->openFile($pFilename);
        } catch (\Exception $e) {
            return \false;
        }

        $readable = $this->isValidFormat();
        \fclose($this->fileHandle);
        return $readable;
    }

    /**
     * Scan theXML for use of <!ENTITY to prevent XXE/XEE attacks
     *
     * @param     string         $xml
     * @throws \PhpOffice\PhpSpreadsheet\Reader\Exception
     */
    public function securityScan($xml)
    {
        $pattern = '/encoding="(.*?)"/';
        $result = \preg_match($pattern, $xml, $matches);
        $charset = $result ? $matches[1] : 'UTF-8';

        if ($charset !== 'UTF-8') {
            $xml = \mb_convert_encoding($xml, 'UTF-8', $charset);
        }

        $pattern = '/\\0?' . \implode('\\0?', \str_split('<!DOCTYPE')) . '\\0?/';
        if (\preg_match($pattern, $xml)) {
            throw new \PhpOffice\PhpSpreadsheet\Reader\Exception('Detected use of ENTITY in XML, spreadsheet file load() aborted to prevent XXE/XEE attacks');
        }
        return $xml;
    }

    /**
     * Scan theXML for use of <!ENTITY to prevent XXE/XEE attacks
     *
     * @param     string         $filestream
     * @throws \PhpOffice\PhpSpreadsheet\Reader\Exception
     */
    public function securityScanFile($filestream)
    {
        return $this->securityScan(\file_get_contents($filestream));
    }
}

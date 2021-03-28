<?php

/** PHPExcel root directory */
if (!defined('PHPEXCEL_ROOT')) {
    define('PHPEXCEL_ROOT', dirname(__FILE__) . '/');
    require(PHPEXCEL_ROOT . 'PHPExcel/Autoloader.php');
}

namespace PhpOffice\PhpSpreadsheet;

/**
 * PHPExcel
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
 * @package    PHPExcel
 * @copyright  Copyright (c) 2006 - 2015 PHPExcel (http://www.codeplex.com/PHPExcel)
 * @license    http://www.gnu.org/licenses/old-licenses/lgpl-2.1.txt    LGPL
 * @version    ##VERSION##, ##DATE##
 */
class Spreadsheet
{
    /**
     * Unique ID
     *
     * @var string
     */
    private $uniqueID;

    /**
     * Document properties
     *
     * @var \PhpOffice\PhpSpreadsheet\Document\Properties
     */
    private $phpExcelDocumentProperties;

    /**
     * Document security
     *
     * @var \PhpOffice\PhpSpreadsheet\Document\Security
     */
    private $phpExcelDocumentSecurity;

    /**
     * Collection of Worksheet objects
     *
     * @var \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet[]
     */
    private $workSheetCollection = array();

    /**
     * Calculation Engine
     *
     * @var \PhpOffice\PhpSpreadsheet\Calculation\Calculation
     */
    private $phpExcelCalculation;

    /**
     * Active sheet index
     *
     * @var integer
     */
    private $activeSheetIndex = 0;

    /**
     * Named ranges
     *
     * @var \PhpOffice\PhpSpreadsheet\NamedRange[]
     */
    private $namedRanges = array();

    /**
     * CellXf supervisor
     *
     * @var \PhpOffice\PhpSpreadsheet\Style\Style
     */
    private $phpExcelStyle;

    /**
     * CellXf collection
     *
     * @var \PhpOffice\PhpSpreadsheet\Style\Style[]
     */
    private $cellXfCollection = array();

    /**
     * CellStyleXf collection
     *
     * @var \PhpOffice\PhpSpreadsheet\Style\Style[]
     */
    private $cellStyleXfCollection = array();

    /**
    * hasMacros : this workbook have macros ?
    *
    * @var bool
    */
    private $hasMacros = \false;

    /**
    * macrosCode : all macros code (the vbaProject.bin file, this include form, code,  etc.), null if no macro
    *
    * @var binary
    */
    private $macrosCode;
    /**
    * macrosCertificate : if macros are signed, contains vbaProjectSignature.bin file, null if not signed
    *
    * @var binary
    */
    private $macrosCertificate;

    /**
    * ribbonXMLData : null if workbook is'nt Excel 2007 or not contain a customized UI
    *
    * @var null|string
    */
    private $ribbonXMLData;

    /**
    * ribbonBinObjects : null if workbook is'nt Excel 2007 or not contain embedded objects (picture(s)) for Ribbon Elements
    * ignored if $ribbonXMLData is null
    *
    * @var null|array
    */
    private $ribbonBinObjects;

    /**
    * The workbook has macros ?
    *
    * @return boolean true if workbook has macros, false if not
    */
    public function hasMacros()
    {
        return $this->hasMacros;
    }

    /**
    * Define if a workbook has macros
    *
    * @param boolean $hasMacros true|false
    */
    public function setHasMacros($hasMacros = \false)
    {
        $this->hasMacros = $hasMacros;
    }

    /**
    * Set the macros code
    *
    * @param string $MacrosCode string|null
    */
    public function setMacrosCode($MacrosCode = \null)
    {
        $this->macrosCode=$MacrosCode;
        $this->setHasMacros(!\is_null($MacrosCode));
    }

    /**
    * Return the macros code
    *
    * @return string|null
    */
    public function getMacrosCode()
    {
        return $this->macrosCode;
    }

    /**
    * Set the macros certificate
    *
    * @param string|null $Certificate
    */
    public function setMacrosCertificate($Certificate = \null)
    {
        $this->macrosCertificate=$Certificate;
    }

    /**
    * Is the project signed ?
    *
    * @return boolean true|false
    */
    public function hasMacrosCertificate()
    {
        return !\is_null($this->macrosCertificate);
    }

    /**
    * Return the macros certificate
    *
    * @return string|null
    */
    public function getMacrosCertificate()
    {
        return $this->macrosCertificate;
    }

    /**
    * Remove all macros, certificate from spreadsheet
    *
    */
    public function discardMacros()
    {
        $this->hasMacros=\false;
        $this->macrosCode=\null;
        $this->macrosCertificate=\null;
    }

    /**
    * set ribbon XML data
    *
    */
    public function setRibbonXMLData($Target = \null, $XMLData = \null)
    {
        if (!\is_null($Target) && !\is_null($XMLData)) {
            $this->ribbonXMLData = array('target' => $Target, 'data' => $XMLData);
        } else {
            $this->ribbonXMLData = \null;
        }
    }

    /**
    * retrieve ribbon XML Data
    *
    * return string|null|array
    */
    public function getRibbonXMLData($What = 'all') //we need some constants here...
    {
        $ReturnData = \null;
        $What = \strtolower($What);
        switch ($What){
            case 'all':
                $ReturnData = $this->ribbonXMLData;
                break;
            case 'target':
            case 'data':
                if (\is_array($this->ribbonXMLData) && \array_key_exists($What, $this->ribbonXMLData)) {
                    $ReturnData = $this->ribbonXMLData[$What];
                }
                break;
        }

        return $ReturnData;
    }

    /**
    * store binaries ribbon objects (pictures)
    *
    */
    public function setRibbonBinObjects($BinObjectsNames = \null, $BinObjectsData = \null)
    {
        if (!\is_null($BinObjectsNames) && !\is_null($BinObjectsData)) {
            $this->ribbonBinObjects = array('names' => $BinObjectsNames, 'data' => $BinObjectsData);
        } else {
            $this->ribbonBinObjects = \null;
        }
    }
    /**
    * return the extension of a filename. Internal use for a array_map callback (php<5.3 don't like lambda function)
    *
    */
    private function getExtensionOnly($ThePath)
    {
        return \pathinfo($ThePath, \PATHINFO_EXTENSION);
    }

    /**
    * retrieve Binaries Ribbon Objects
    *
    */
    public function getRibbonBinObjects($What = 'all')
    {
        $ReturnData = \null;
        $What = \strtolower($What);
        switch($What) {
            case 'all':
                return $this->ribbonBinObjects;
            case 'names':
            case 'data':
                if (\is_array($this->ribbonBinObjects) && \array_key_exists($What, $this->ribbonBinObjects)) {
                    $ReturnData=$this->ribbonBinObjects[$What];
                }
                break;
            case 'types':
                if (\is_array($this->ribbonBinObjects) &&
                    \array_key_exists('data', $this->ribbonBinObjects) && \is_array($this->ribbonBinObjects['data'])) {
                    $tmpTypes=\array_keys($this->ribbonBinObjects['data']);
                    $ReturnData = \array_unique(\array_map(function ($ThePath) {
                        return $this->getExtensionOnly($ThePath);
                    }, $tmpTypes));
                } else {
                    $ReturnData=array(); // the caller want an array... not null if empty
                }
                break;
        }
        return $ReturnData;
    }

    /**
    * This workbook have a custom UI ?
    *
    * @return boolean true|false
    */
    public function hasRibbon()
    {
        return !\is_null($this->ribbonXMLData);
    }

    /**
    * This workbook have additionnal object for the ribbon ?
    *
    * @return boolean true|false
    */
    public function hasRibbonBinObjects()
    {
        return !\is_null($this->ribbonBinObjects);
    }

    /**
     * Check if a sheet with a specified code name already exists
     *
     * @param string $pSheetCodeName  Name of the worksheet to check
     * @return boolean
     */
    public function sheetCodeNameExists($pSheetCodeName)
    {
        return ($this->getSheetByCodeName($pSheetCodeName) !== \null);
    }

    /**
     * Get sheet by code name. Warning : sheet don't have always a code name !
     *
     * @param string $pName Sheet name
     * @return \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet
     */
    public function getSheetByCodeName($pName = '')
    {
        $worksheetCount = \count($this->workSheetCollection);
        for ($i = 0; $i < $worksheetCount; ++$i) {
            if ($this->workSheetCollection[$i]->getCodeName() == $pName) {
                return $this->workSheetCollection[$i];
            }
        }

        return \null;
    }

     /**
     * Create a new PHPExcel with one Worksheet
     */
    public function __construct()
    {
        $this->uniqueID = \uniqid();
        $this->phpExcelCalculation = new \PhpOffice\PhpSpreadsheet\Calculation\Calculation($this);

        // Initialise worksheet collection and add one worksheet
        $this->workSheetCollection = array();
        $this->workSheetCollection[] = new \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet($this);
        $this->activeSheetIndex = 0;

        // Create document properties
        $this->phpExcelDocumentProperties = new \PhpOffice\PhpSpreadsheet\Document\Properties();

        // Create document security
        $this->phpExcelDocumentSecurity = new \PhpOffice\PhpSpreadsheet\Document\Security();

        // Set named ranges
        $this->namedRanges = array();

        // Create the cellXf supervisor
        $this->phpExcelStyle = new \PhpOffice\PhpSpreadsheet\Style\Style(\true);
        $this->phpExcelStyle->bindParent($this);

        // Create the default style
        $this->addCellXf(new \PhpOffice\PhpSpreadsheet\Style\Style);
        $this->addCellStyleXf(new \PhpOffice\PhpSpreadsheet\Style\Style);
    }

    /**
     * Code to execute when this worksheet is unset()
     *
     */
    public function __destruct()
    {
        $this->phpExcelCalculation = \null;
        $this->disconnectWorksheets();
    }

    /**
     * Disconnect all worksheets from this PHPExcel workbook object,
     *    typically so that the PHPExcel object can be unset
     *
     */
    public function disconnectWorksheets()
    {
        $worksheet = \null;
        foreach ($this->workSheetCollection as $k => &$worksheet) {
            $worksheet->disconnectCells();
            $this->workSheetCollection[$k] = \null;
        }
        unset($worksheet);
        $this->workSheetCollection = array();
    }

    /**
     * Return the calculation engine for this worksheet
     *
     * @return \PhpOffice\PhpSpreadsheet\Calculation\Calculation
     */
    public function getCalculationEngine()
    {
        return $this->phpExcelCalculation;
    }    //    function getCellCacheController()
    /**
     * Get properties
     *
     * @return \PhpOffice\PhpSpreadsheet\Document\Properties
     */
    public function getProperties()
    {
        return $this->phpExcelDocumentProperties;
    }

    /**
     * Set properties
     */
    public function setProperties(\PhpOffice\PhpSpreadsheet\Document\Properties $phpExcelDocumentProperties)
    {
        $this->phpExcelDocumentProperties = $phpExcelDocumentProperties;
    }

    /**
     * Get security
     *
     * @return \PhpOffice\PhpSpreadsheet\Document\Security
     */
    public function getSecurity()
    {
        return $this->phpExcelDocumentSecurity;
    }

    /**
     * Set security
     */
    public function setSecurity(\PhpOffice\PhpSpreadsheet\Document\Security $phpExcelDocumentSecurity)
    {
        $this->phpExcelDocumentSecurity = $phpExcelDocumentSecurity;
    }

    /**
     * Get active sheet
     *
     * @return \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet
     *
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     */
    public function getActiveSheet()
    {
        return $this->getSheet($this->activeSheetIndex);
    }

    /**
     * Create sheet and add it to this workbook
     *
     * @param  int|null $iSheetIndex Index where sheet should go (0,1,..., or null for last)
     * @return \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     */
    public function createSheet($iSheetIndex = \null)
    {
        $phpExcelWorksheet = new \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet($this);
        $this->addSheet($phpExcelWorksheet, $iSheetIndex);
        return $phpExcelWorksheet;
    }

    /**
     * Check if a sheet with a specified name already exists
     *
     * @param  string $pSheetName  Name of the worksheet to check
     * @return boolean
     */
    public function sheetNameExists($pSheetName)
    {
        return ($this->getSheetByName($pSheetName) !== \null);
    }

    /**
     * Add sheet
     *
     * @param  int|null $iSheetIndex Index where sheet should go (0,1,..., or null for last)
     * @return \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     */
    public function addSheet(\PhpOffice\PhpSpreadsheet\Worksheet\Worksheet $phpExcelWorksheet, $iSheetIndex = \null)
    {
        if ($this->sheetNameExists($phpExcelWorksheet->getTitle())) {
            throw new \PhpOffice\PhpSpreadsheet\Exception(
                "Workbook already contains a worksheet named '{$phpExcelWorksheet->getTitle()}'. Rename this worksheet first."
            );
        }

        if ($iSheetIndex === \null) {
            if ($this->activeSheetIndex < 0) {
                $this->activeSheetIndex = 0;
            }
            $this->workSheetCollection[] = $phpExcelWorksheet;
        } else {
            // Insert the sheet at the requested index
            \array_splice(
                $this->workSheetCollection,
                $iSheetIndex,
                0,
                array($phpExcelWorksheet)
            );

            // Adjust active sheet index if necessary
            if ($this->activeSheetIndex >= $iSheetIndex) {
                ++$this->activeSheetIndex;
            }
        }

        if ($phpExcelWorksheet->getParent() === \null) {
            $phpExcelWorksheet->rebindParent($this);
        }

        return $phpExcelWorksheet;
    }

    /**
     * Remove sheet by index
     *
     * @param  int $pIndex Active sheet index
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     */
    public function removeSheetByIndex($pIndex = 0)
    {

        $numSheets = \count($this->workSheetCollection);
        if ($pIndex > $numSheets - 1) {
            throw new \PhpOffice\PhpSpreadsheet\Exception(
                "You tried to remove a sheet by the out of bounds index: {$pIndex}. The actual number of sheets is {$numSheets}."
            );
        } else {
            \array_splice($this->workSheetCollection, $pIndex, 1);
        }
        // Adjust active sheet index if necessary
        if (($this->activeSheetIndex >= $pIndex) &&
            ($pIndex > \count($this->workSheetCollection) - 1)) {
            --$this->activeSheetIndex;
        }

    }

    /**
     * Get sheet by index
     *
     * @param  int $pIndex Sheet index
     * @return \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     */
    public function getSheet($pIndex = 0)
    {
        if (!isset($this->workSheetCollection[$pIndex])) {
            $numSheets = $this->getSheetCount();
            throw new \PhpOffice\PhpSpreadsheet\Exception(
                "Your requested sheet index: {$pIndex} is out of bounds. The actual number of sheets is {$numSheets}."
            );
        }

        return $this->workSheetCollection[$pIndex];
    }

    /**
     * Get all sheets
     *
     * @return \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet[]
     */
    public function getAllSheets()
    {
        return $this->workSheetCollection;
    }

    /**
     * Get sheet by name
     *
     * @param  string $pName Sheet name
     * @return \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet
     */
    public function getSheetByName($pName = '')
    {
        $worksheetCount = \count($this->workSheetCollection);
        for ($i = 0; $i < $worksheetCount; ++$i) {
            if ($this->workSheetCollection[$i]->getTitle() === $pName) {
                return $this->workSheetCollection[$i];
            }
        }

        return \null;
    }

    /**
     * Get index for sheet
     *
     * @return int Sheet index
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     */
    public function getIndex(\PhpOffice\PhpSpreadsheet\Worksheet\Worksheet $phpExcelWorksheet)
    {
        foreach ($this->workSheetCollection as $key => $value) {
            if ($value->getHashCode() == $phpExcelWorksheet->getHashCode()) {
                return $key;
            }
        }

        throw new \PhpOffice\PhpSpreadsheet\Exception("Sheet does not exist.");
    }

    /**
     * Set index for sheet by sheet name.
     *
     * @param  string $sheetName Sheet name to modify index for
     * @param  int $newIndex New index for the sheet
     * @return int New sheet index
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     */
    public function setIndexByName($sheetName, $newIndex)
    {
        $oldIndex = $this->getIndex($this->getSheetByName($sheetName));
        $pSheet = \array_splice(
            $this->workSheetCollection,
            $oldIndex,
            1
        );
        \array_splice(
            $this->workSheetCollection,
            $newIndex,
            0,
            $pSheet
        );
        return $newIndex;
    }

    /**
     * Get sheet count
     *
     * @return int
     */
    public function getSheetCount()
    {
        return \count($this->workSheetCollection);
    }

    /**
     * Get active sheet index
     *
     * @return int Active sheet index
     */
    public function getActiveSheetIndex()
    {
        return $this->activeSheetIndex;
    }

    /**
     * Set active sheet index
     *
     * @param  int $pIndex Active sheet index
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     * @return \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet
     */
    public function setActiveSheetIndex($pIndex = 0)
    {
        $numSheets = \count($this->workSheetCollection);

        if ($pIndex > $numSheets - 1) {
            throw new \PhpOffice\PhpSpreadsheet\Exception(
                "You tried to set a sheet active by the out of bounds index: {$pIndex}. The actual number of sheets is {$numSheets}."
            );
        } else {
            $this->activeSheetIndex = $pIndex;
        }
        return $this->getActiveSheet();
    }

    /**
     * Set active sheet index by name
     *
     * @param  string $pValue Sheet title
     * @return \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     */
    public function setActiveSheetIndexByName($pValue = '')
    {
        $this->setActiveSheetIndex($this->getIndex($worksheet));
        return $worksheet;
        throw new \PhpOffice\PhpSpreadsheet\Exception('Workbook does not contain sheet:' . $pValue);
    }

    /**
     * Get sheet names
     *
     * @return string[]
     */
    public function getSheetNames()
    {
        $returnValue = array();
        $worksheetCount = $this->getSheetCount();
        for ($i = 0; $i < $worksheetCount; ++$i) {
            $returnValue[] = $this->getSheet($i)->getTitle();
        }

        return $returnValue;
    }

    /**
     * Add external sheet
     *
     * @param \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet $phpExcelWorksheet External sheet to add
     * @param  int|null $iSheetIndex Index where sheet should go (0,1,..., or null for last)
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     * @return \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet
     */
    public function addExternalSheet(\PhpOffice\PhpSpreadsheet\Worksheet\Worksheet $phpExcelWorksheet, $iSheetIndex = \null)
    {
        if ($this->sheetNameExists($phpExcelWorksheet->getTitle())) {
            throw new \PhpOffice\PhpSpreadsheet\Exception("Workbook already contains a worksheet named '{$phpExcelWorksheet->getTitle()}'. Rename the external sheet first.");
        }

        // count how many cellXfs there are in this workbook currently, we will need this below
        $countCellXfs = \count($this->cellXfCollection);

        // copy all the shared cellXfs from the external workbook and append them to the current
        foreach ($phpExcelWorksheet->getParent()->getCellXfCollection() as $cellXf) {
            $this->addCellXf(clone $cellXf);
        }

        // move sheet to this workbook
        $phpExcelWorksheet->rebindParent($this);

        // update the cellXfs
        foreach ($phpExcelWorksheet->getCellCollection(\false) as $cellID) {
            $cell = $phpExcelWorksheet->getCell($cellID);
            $cell->setXfIndex($cell->getXfIndex() + $countCellXfs);
        }

        return $this->addSheet($phpExcelWorksheet, $iSheetIndex);
    }

    /**
     * Get named ranges
     *
     * @return \PhpOffice\PhpSpreadsheet\NamedRange[]
     */
    public function getNamedRanges()
    {
        return $this->namedRanges;
    }

    /**
     * Add named range
     *
     * @return boolean
     */
    public function addNamedRange(\PhpOffice\PhpSpreadsheet\NamedRange $phpExcelNamedRange)
    {
        if ($phpExcelNamedRange->getScope() == \null) {
            // global scope
            $this->namedRanges[$phpExcelNamedRange->getName()] = $phpExcelNamedRange;
        } else {
            // local scope
            $this->namedRanges[$phpExcelNamedRange->getScope()->getTitle().'!'.$phpExcelNamedRange->getName()] = $phpExcelNamedRange;
        }
        return \true;
    }

    /**
     * Get named range
     *
     * @param  string $namedRange
     * @param \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet|null $phpExcelWorksheet Scope. Use null for global scope
     * @return \PhpOffice\PhpSpreadsheet\NamedRange|null
     */
    public function getNamedRange($namedRange, \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet $phpExcelWorksheet = \null)
    {
        $returnValue = \null;

        if ($namedRange != '' && ($namedRange !== \null)) {
            // first look for global defined name
            if (isset($this->namedRanges[$namedRange])) {
                $returnValue = $this->namedRanges[$namedRange];
            }

            // then look for local defined name (has priority over global defined name if both names exist)
            if (($phpExcelWorksheet !== \null) && isset($this->namedRanges[$phpExcelWorksheet->getTitle() . '!' . $namedRange])) {
                $returnValue = $this->namedRanges[$phpExcelWorksheet->getTitle() . '!' . $namedRange];
            }
        }

        return $returnValue;
    }

    /**
     * Remove named range
     *
     * @param  string  $namedRange
     * @param \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet|null  $phpExcelWorksheet  Scope: use null for global scope.
     * @return \PhpOffice\PhpSpreadsheet\Spreadsheet
     */
    public function removeNamedRange($namedRange, \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet $phpExcelWorksheet = \null)
    {
        if ($phpExcelWorksheet === \null) {
            if (isset($this->namedRanges[$namedRange])) {
                unset($this->namedRanges[$namedRange]);
            }
        } elseif (isset($this->namedRanges[$phpExcelWorksheet->getTitle() . '!' . $namedRange])) {
            unset($this->namedRanges[$phpExcelWorksheet->getTitle() . '!' . $namedRange]);
        }
        return $this;
    }

    /**
     * Get worksheet iterator
     *
     * @return \PhpOffice\PhpSpreadsheet\Worksheet\Iterator
     */
    public function getWorksheetIterator()
    {
        return new \PhpOffice\PhpSpreadsheet\Worksheet\Iterator($this);
    }

    /**
     * Copy workbook (!= clone!)
     *
     * @return \PhpOffice\PhpSpreadsheet\Spreadsheet
     */
    public function copy()
    {
        $copied = clone $this;

        $worksheetCount = \count($this->workSheetCollection);
        for ($i = 0; $i < $worksheetCount; ++$i) {
            $this->workSheetCollection[$i] = $this->workSheetCollection[$i]->copy();
            $this->workSheetCollection[$i]->rebindParent($this);
        }

        return $copied;
    }

    /**
     * Implement PHP __clone to create a deep clone, not just a shallow copy.
     */
    public function __clone()
    {
        foreach ($this as $key => $val) {
            if (\is_object($val) || (\is_array($val))) {
                $this->{$key} = \unserialize(\serialize($val));
            }
        }
    }

    /**
     * Get the workbook collection of cellXfs
     *
     * @return \PhpOffice\PhpSpreadsheet\Style\Style[]
     */
    public function getCellXfCollection()
    {
        return $this->cellXfCollection;
    }

    /**
     * Get cellXf by index
     *
     * @param  int $pIndex
     * @return \PhpOffice\PhpSpreadsheet\Style\Style
     */
    public function getCellXfByIndex($pIndex = 0)
    {
        return $this->cellXfCollection[$pIndex];
    }

    /**
     * Get cellXf by hash code
     *
     * @param  string $pValue
     * @return \PhpOffice\PhpSpreadsheet\Style\Style|boolean False if no match found
     */
    public function getCellXfByHashCode($pValue = '')
    {
        foreach ($this->cellXfCollection as $singleCellXfCollection) {
            if ($singleCellXfCollection->getHashCode() == $pValue) {
                return $singleCellXfCollection;
            }
        }
        return \false;
    }

    /**
     * Check if style exists in style collection
     *
     * @param  \PhpOffice\PhpSpreadsheet\Style\Style $pCellStyle
     * @return boolean
     */
    public function cellXfExists($pCellStyle = \null)
    {
        return \in_array($pCellStyle, $this->cellXfCollection, \true);
    }

    /**
     * Get default style
     *
     * @return \PhpOffice\PhpSpreadsheet\Style\Style
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     */
    public function getDefaultStyle()
    {
        if (isset($this->cellXfCollection[0])) {
            return $this->cellXfCollection[0];
        }
        throw new \PhpOffice\PhpSpreadsheet\Exception('No default style found for this workbook');
    }

    /**
     * Add a cellXf to the workbook
     */
    public function addCellXf(\PhpOffice\PhpSpreadsheet\Style\Style $phpExcelStyle)
    {
        $this->cellXfCollection[] = $phpExcelStyle;
        $phpExcelStyle->setIndex(\count($this->cellXfCollection) - 1);
    }

    /**
     * Remove cellXf by index. It is ensured that all cells get their xf index updated.
     *
     * @param integer $pIndex Index to cellXf
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     */
    public function removeCellXfByIndex($pIndex = 0)
    {
        if ($pIndex > \count($this->cellXfCollection) - 1) {
            throw new \PhpOffice\PhpSpreadsheet\Exception("CellXf index is out of bounds.");
        } else {
            // first remove the cellXf
            \array_splice($this->cellXfCollection, $pIndex, 1);

            // then update cellXf indexes for cells
            foreach ($this->workSheetCollection as $singleWorkSheetCollection) {
                foreach ($singleWorkSheetCollection->getCellCollection(\false) as $cellID) {
                    $cell = $singleWorkSheetCollection->getCell($cellID);
                    $xfIndex = $cell->getXfIndex();
                    if ($xfIndex > $pIndex) {
                        // decrease xf index by 1
                        $cell->setXfIndex($xfIndex - 1);
                    } elseif ($xfIndex === $pIndex) {
                        // set to default xf index 0
                        $cell->setXfIndex(0);
                    }
                }
            }
        }
    }

    /**
     * Get the cellXf supervisor
     *
     * @return \PhpOffice\PhpSpreadsheet\Style\Style
     */
    public function getCellXfSupervisor()
    {
        return $this->phpExcelStyle;
    }

    /**
     * Get the workbook collection of cellStyleXfs
     *
     * @return \PhpOffice\PhpSpreadsheet\Style\Style[]
     */
    public function getCellStyleXfCollection()
    {
        return $this->cellStyleXfCollection;
    }

    /**
     * Get cellStyleXf by index
     *
     * @param integer $pIndex Index to cellXf
     * @return \PhpOffice\PhpSpreadsheet\Style\Style
     */
    public function getCellStyleXfByIndex($pIndex = 0)
    {
        return $this->cellStyleXfCollection[$pIndex];
    }

    /**
     * Get cellStyleXf by hash code
     *
     * @param  string $pValue
     * @return \PhpOffice\PhpSpreadsheet\Style\Style|boolean False if no match found
     */
    public function getCellStyleXfByHashCode($pValue = '')
    {
        foreach ($this->cellStyleXfCollection as $singleCellStyleXfCollection) {
            if ($singleCellStyleXfCollection->getHashCode() == $pValue) {
                return $singleCellStyleXfCollection;
            }
        }
        return \false;
    }

    /**
     * Add a cellStyleXf to the workbook
     */
    public function addCellStyleXf(\PhpOffice\PhpSpreadsheet\Style\Style $phpExcelStyle)
    {
        $this->cellStyleXfCollection[] = $phpExcelStyle;
        $phpExcelStyle->setIndex(\count($this->cellStyleXfCollection) - 1);
    }

    /**
     * Remove cellStyleXf by index
     *
     * @param integer $pIndex Index to cellXf
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     */
    public function removeCellStyleXfByIndex($pIndex = 0)
    {
        if ($pIndex > \count($this->cellStyleXfCollection) - 1) {
            throw new \PhpOffice\PhpSpreadsheet\Exception("CellStyleXf index is out of bounds.");
        } else {
            \array_splice($this->cellStyleXfCollection, $pIndex, 1);
        }
    }

    /**
     * Eliminate all unneeded cellXf and afterwards update the xfIndex for all cells
     * and columns in the workbook
     */
    public function garbageCollect()
    {
        // how many references are there to each cellXf ?
        $countReferencesCellXf = array();
        foreach (array_keys($this->cellXfCollection) as $index) {
            $countReferencesCellXf[$index] = 0;
        }

        foreach ($this->getWorksheetIterator() as $sheet) {
            // from cells
            foreach ($sheet->getCoordinates(\false) as $phpExcelCell) {
                $cell = $sheet->getCell($phpExcelCell, true);
                ++$countReferencesCellXf[$cell->getXfIndex()];
            }

            // from row dimensions
            foreach ($sheet->getRowDimensions() as $phpExcelWorksheetRowDimension) {
                if ($phpExcelWorksheetRowDimension->getXfIndex() !== \null) {
                    ++$countReferencesCellXf[$phpExcelWorksheetRowDimension->getXfIndex()];
                }
            }

            // from column dimensions
            foreach ($sheet->getColumnDimensions() as $phpExcelWorksheetColumnDimension) {
                ++$countReferencesCellXf[$phpExcelWorksheetColumnDimension->getXfIndex()];
            }
        }

        // remove cellXfs without references and create mapping so we can update xfIndex
        // for all cells and columns
        $countNeededCellXfs = 0;
        $map = array();
        foreach (array_keys($this->cellXfCollection) as $index) {
            if ($countReferencesCellXf[$index] > 0 || $index == 0) { // we must never remove the first cellXf
                ++$countNeededCellXfs;
            } else {
                unset($this->cellXfCollection[$index]);
            }
            $map[$index] = $countNeededCellXfs - 1;
        }
        $this->cellXfCollection = \array_values($this->cellXfCollection);

        // update the index for all cellXfs
        foreach ($this->cellXfCollection as $i => $cellXf) {
            $cellXf->setIndex($i);
        }

        // make sure there is always at least one cellXf (there should be)
        if (empty($this->cellXfCollection)) {
            $this->cellXfCollection[] = new \PhpOffice\PhpSpreadsheet\Style\Style();
        }

        // update the xfIndex for all cells, row dimensions, column dimensions
        foreach ($this->getWorksheetIterator() as $sheet) {
            // for all cells
            foreach ($sheet->getCoordinates(\false) as $phpExcelCell) {
                $cell = $sheet->getCell($phpExcelCell, true);
                $cell->setXfIndex($map[$cell->getXfIndex()]);
            }

            // for all row dimensions
            foreach ($sheet->getRowDimensions() as $phpExcelWorksheetRowDimension) {
                if ($phpExcelWorksheetRowDimension->getXfIndex() !== \null) {
                    $phpExcelWorksheetRowDimension->setXfIndex($map[$phpExcelWorksheetRowDimension->getXfIndex()]);
                }
            }

            // for all column dimensions
            foreach ($sheet->getColumnDimensions() as $phpExcelWorksheetColumnDimension) {
                $phpExcelWorksheetColumnDimension->setXfIndex($map[$phpExcelWorksheetColumnDimension->getXfIndex()]);
            }

            // also do garbage collection for all the sheets
            $sheet->garbageCollect();
        }
    }

    /**
     * Return the unique ID value assigned to this spreadsheet workbook
     *
     * @return string
     */
    public function getID()
    {
        return $this->uniqueID;
    }
}

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
 * PHPExcel_Reader_Excel2003XML
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
class Xml extends \PhpOffice\PhpSpreadsheet\Reader\BaseReader implements \PhpOffice\PhpSpreadsheet\Reader\IReader
{
    /**
     * Formats
     *
     * @var array
     */
    protected $styles = array();

    /**
     * Character set used in the file
     *
     * @var string
     */
    protected $charSet = 'UTF-8';

    /**
     * Create a new PHPExcel_Reader_Excel2003XML
     */
    public function __construct()
    {
        $this->readFilter = new \PhpOffice\PhpSpreadsheet\Reader\DefaultReadFilter();
    }


    /**
     * Can the current PHPExcel_Reader_IReader read the file?
     *
     * @param     string         $pFilename
     * @return     boolean
     * @throws \PhpOffice\PhpSpreadsheet\Reader\Exception
     */
    public function canRead($pFilename)
    {

        //    Office                    xmlns:o="urn:schemas-microsoft-com:office:office"
        //    Excel                    xmlns:x="urn:schemas-microsoft-com:office:excel"
        //    XML Spreadsheet            xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet"
        //    Spreadsheet component    xmlns:c="urn:schemas-microsoft-com:office:component:spreadsheet"
        //    XML schema                 xmlns:s="uuid:BDC6E3F0-6DA3-11d1-A2A3-00AA00C14882"
        //    XML data type            xmlns:dt="uuid:C2F41010-65B3-11d1-A29F-00AA00C14882"
        //    MS-persist recordset    xmlns:rs="urn:schemas-microsoft-com:rowset"
        //    Rowset                    xmlns:z="#RowsetSchema"
        //

        $signature = array(
                '<?xml version="1.0"',
                '<?mso-application progid="Excel.Sheet"?>'
            );

        // Open file
        $this->openFile($pFilename);
        $fileHandle = $this->fileHandle;
        
        // Read sample data (first 2 KB will do)
        $data = \fread($fileHandle, 2048);
        \fclose($fileHandle);

        $valid = \true;
        foreach ($signature as $singleSignature) {
            // every part of the signature must be present
            if (\strpos($data, $singleSignature) === \false) {
                $valid = \false;
                break;
            }
        }

        //    Retrieve charset encoding
        if (\preg_match('/<?xml.*encoding=[\'"](.*?)[\'"].*?>/um', $data, $matches)) {
            $this->charSet = \strtoupper($matches[1]);
        }
//        echo 'Character Set is ', $this->charSet,'<br />';

        return $valid;
    }


    /**
     * Reads names of the worksheets from a file, without parsing the whole file to a PHPExcel object
     *
     * @param     string         $pFilename
     * @throws     \PhpOffice\PhpSpreadsheet\Reader\Exception
     */
    public function listWorksheetNames($pFilename)
    {
        // Check if file exists
        if (!\file_exists($pFilename)) {
            throw new \PhpOffice\PhpSpreadsheet\Reader\Exception("Could not open " . $pFilename . " for reading! File does not exist.");
        }
        if (!$this->canRead($pFilename)) {
            throw new \PhpOffice\PhpSpreadsheet\Reader\Exception($pFilename . " is an Invalid Spreadsheet file.");
        }

        $worksheetNames = array();

        $xml = \simplexml_load_string($this->securityScan(\file_get_contents($pFilename)), 'SimpleXMLElement', \PhpOffice\PhpSpreadsheet\Settings::getLibXmlLoaderOptions());
        $namespaces = $xml->getNamespaces(\true);

        $xml_ss = $xml->children($namespaces['ss']);
        foreach ($xml_ss->Worksheet as $worksheet) {
            $worksheet_ss = $worksheet->attributes($namespaces['ss']);
            $worksheetNames[] = self::convertStringEncoding((string) $worksheet_ss['Name'], $this->charSet);
        }

        return $worksheetNames;
    }


    /**
     * Return worksheet info (Name, Last Column Letter, Last Column Index, Total Rows, Total Columns)
     *
     * @param   string     $pFilename
     * @throws   \PhpOffice\PhpSpreadsheet\Reader\Exception
     */
    public function listWorksheetInfo($pFilename)
    {
        // Check if file exists
        if (!\file_exists($pFilename)) {
            throw new \PhpOffice\PhpSpreadsheet\Reader\Exception("Could not open " . $pFilename . " for reading! File does not exist.");
        }

        $worksheetInfo = array();

        $xml = \simplexml_load_string($this->securityScan(\file_get_contents($pFilename)), 'SimpleXMLElement', \PhpOffice\PhpSpreadsheet\Settings::getLibXmlLoaderOptions());
        $namespaces = $xml->getNamespaces(\true);

        $worksheetID = 1;
        $xml_ss = $xml->children($namespaces['ss']);
        foreach ($xml_ss->Worksheet as $worksheet) {
            $worksheet_ss = $worksheet->attributes($namespaces['ss']);

            $tmpInfo = array();
            $tmpInfo['worksheetName'] = '';
            $tmpInfo['lastColumnLetter'] = 'A';
            $tmpInfo['lastColumnIndex'] = 0;
            $tmpInfo['totalRows'] = 0;
            $tmpInfo['totalColumns'] = 0;

            $tmpInfo['worksheetName'] = isset($worksheet_ss['Name']) ? (string) $worksheet_ss['Name'] : "Worksheet_{$worksheetID}";

            if (property_exists($worksheet->Table, 'Row') && $worksheet->Table->Row !== null) {
                $rowIndex = 0;

                foreach ($worksheet->Table->Row as $rowData) {
                    $columnIndex = 0;
                    $rowHasData = \false;

                    foreach ($rowData->Cell as $cell) {
                        if (property_exists($cell, 'Data') && $cell->Data !== null) {
                            $tmpInfo['lastColumnIndex'] = \max($tmpInfo['lastColumnIndex'], $columnIndex);
                            $rowHasData = \true;
                        }

                        ++$columnIndex;
                    }

                    ++$rowIndex;

                    if ($rowHasData) {
                        $tmpInfo['totalRows'] = \max($tmpInfo['totalRows'], $rowIndex);
                    }
                }
            }

            $tmpInfo['lastColumnLetter'] = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::stringFromColumnIndex($tmpInfo['lastColumnIndex']);
            $tmpInfo['totalColumns'] = $tmpInfo['lastColumnIndex'] + 1;

            $worksheetInfo[] = $tmpInfo;
            ++$worksheetID;
        }

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
        $phpExcel->removeSheetByIndex(0);

        // Load into this instance
        return $this->loadIntoExisting($pFilename, $phpExcel);
    }

    protected static function identifyFixedStyleValue($styleList, &$styleAttributeValue)
    {
        $styleAttributeValue = \strtolower($styleAttributeValue);
        foreach ($styleList as $singleStyleList) {
            if ($styleAttributeValue === \strtolower($singleStyleList)) {
                $styleAttributeValue = $singleStyleList;
                return \true;
            }
        }
        return \false;
    }

    /**
     * pixel units to excel width units(units of 1/256th of a character width)
     * @param pxs
     * @return
     */
    protected static function pixel2WidthUnits($pxs)
    {
        $UNIT_OFFSET_MAP = array(0, 36, 73, 109, 146, 182, 219);

        $widthUnits = 256 * ($pxs / 7);
        return $widthUnits + $UNIT_OFFSET_MAP[($pxs % 7)];
    }

    /**
     * excel width units(units of 1/256th of a character width) to pixel units
     * @param widthUnits
     * @return
     */
    protected static function widthUnits2Pixel($widthUnits)
    {
        $pixels = ($widthUnits / 256) * 7;
        $offsetWidthUnits = $widthUnits % 256;
        return $pixels + \round($offsetWidthUnits / (256 / 7));
    }

    protected static function hex2str($hex)
    {
        return \chr(\hexdec($hex[1]));
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
        $fromFormats    = array('\-', '\ ');
        $toFormats      = array('-', ' ');

        $underlineStyles = array (
            \PhpOffice\PhpSpreadsheet\Style\Font::UNDERLINE_NONE,
            \PhpOffice\PhpSpreadsheet\Style\Font::UNDERLINE_DOUBLE,
            \PhpOffice\PhpSpreadsheet\Style\Font::UNDERLINE_DOUBLEACCOUNTING,
            \PhpOffice\PhpSpreadsheet\Style\Font::UNDERLINE_SINGLE,
            \PhpOffice\PhpSpreadsheet\Style\Font::UNDERLINE_SINGLEACCOUNTING
        );
        $verticalAlignmentStyles = array (
            \PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_BOTTOM,
            \PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_TOP,
            \PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER,
            \PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_JUSTIFY
        );
        $horizontalAlignmentStyles = array (
            \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_GENERAL,
            \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_LEFT,
            \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_RIGHT,
            \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER,
            \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER_CONTINUOUS,
            \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_JUSTIFY
        );

        $timezoneObj = new \DateTimeZone('Europe/London');
        $GMT = new \DateTimeZone('UTC');

        // Check if file exists
        if (!\file_exists($pFilename)) {
            throw new \PhpOffice\PhpSpreadsheet\Reader\Exception("Could not open " . $pFilename . " for reading! File does not exist.");
        }

        if (!$this->canRead($pFilename)) {
            throw new \PhpOffice\PhpSpreadsheet\Reader\Exception($pFilename . " is an Invalid Spreadsheet file.");
        }

        $xml = \simplexml_load_string($this->securityScan(\file_get_contents($pFilename)), 'SimpleXMLElement', \PhpOffice\PhpSpreadsheet\Settings::getLibXmlLoaderOptions());
        $namespaces = $xml->getNamespaces(\true);

        $docProps = $phpExcel->getProperties();
        if (isset($xml->DocumentProperties[0])) {
            foreach ($xml->DocumentProperties[0] as $propertyName => $propertyValue) {
                switch ($propertyName) {
                    case 'Title':
                        $docProps->setTitle(self::convertStringEncoding($propertyValue, $this->charSet));
                        break;
                    case 'Subject':
                        $docProps->setSubject(self::convertStringEncoding($propertyValue, $this->charSet));
                        break;
                    case 'Author':
                        $docProps->setCreator(self::convertStringEncoding($propertyValue, $this->charSet));
                        break;
                    case 'Created':
                        $creationDate = \strtotime($propertyValue);
                        $docProps->setCreated($creationDate);
                        break;
                    case 'LastAuthor':
                        $docProps->setLastModifiedBy(self::convertStringEncoding($propertyValue, $this->charSet));
                        break;
                    case 'LastSaved':
                        $lastSaveDate = \strtotime($propertyValue);
                        $docProps->setModified($lastSaveDate);
                        break;
                    case 'Company':
                        $docProps->setCompany(self::convertStringEncoding($propertyValue, $this->charSet));
                        break;
                    case 'Category':
                        $docProps->setCategory(self::convertStringEncoding($propertyValue, $this->charSet));
                        break;
                    case 'Manager':
                        $docProps->setManager(self::convertStringEncoding($propertyValue, $this->charSet));
                        break;
                    case 'Keywords':
                        $docProps->setKeywords(self::convertStringEncoding($propertyValue, $this->charSet));
                        break;
                    case 'Description':
                        $docProps->setDescription(self::convertStringEncoding($propertyValue, $this->charSet));
                        break;
                }
            }
        }
        if (property_exists($xml, 'CustomDocumentProperties') && $xml->CustomDocumentProperties !== null) {
            foreach ($xml->CustomDocumentProperties[0] as $propertyName => $propertyValue) {
                $propertyAttributes = $propertyValue->attributes($namespaces['dt']);
                $propertyName = \preg_replace_callback('/_x([0-9a-z]{4})_/', 'PHPExcel_Reader_Excel2003XML::hex2str', $propertyName);
                $propertyType = \PhpOffice\PhpSpreadsheet\Document\Properties::PROPERTY_TYPE_UNKNOWN;
                switch ((string) $propertyAttributes) {
                    case 'string':
                        $propertyType = \PhpOffice\PhpSpreadsheet\Document\Properties::PROPERTY_TYPE_STRING;
                        $propertyValue = \trim($propertyValue);
                        break;
                    case 'boolean':
                        $propertyType = \PhpOffice\PhpSpreadsheet\Document\Properties::PROPERTY_TYPE_BOOLEAN;
                        $propertyValue = (bool) $propertyValue;
                        break;
                    case 'integer':
                        $propertyType = \PhpOffice\PhpSpreadsheet\Document\Properties::PROPERTY_TYPE_INTEGER;
                        $propertyValue = (int) $propertyValue;
                        break;
                    case 'float':
                        $propertyType = \PhpOffice\PhpSpreadsheet\Document\Properties::PROPERTY_TYPE_FLOAT;
                        $propertyValue = \floatval($propertyValue);
                        break;
                    case 'dateTime.tz':
                        $propertyType = \PhpOffice\PhpSpreadsheet\Document\Properties::PROPERTY_TYPE_DATE;
                        $propertyValue = \strtotime(\trim($propertyValue));
                        break;
                }
                $docProps->setCustomProperty($propertyName, $propertyValue, $propertyType);
            }
        }

        foreach ($xml->Styles[0] as $style) {
            $style_ss = $style->attributes($namespaces['ss']);
            $styleID = (string) $style_ss['ID'];
//            echo 'Style ID = '.$styleID.'<br />';
            $this->styles[$styleID] = (isset($this->styles['Default'])) ? $this->styles['Default'] : array();
            foreach ($style as $styleType => $styleData) {
                $styleAttributes = $styleData->attributes($namespaces['ss']);
//                echo $styleType.'<br />';
                switch ($styleType) {
                    case 'Alignment':
                        foreach ($styleAttributes as $styleAttributeKey => $styleAttributeValue) {
//                                echo $styleAttributeKey.' = '.$styleAttributeValue.'<br />';
                            $styleAttributeValue = (string) $styleAttributeValue;
                            switch ($styleAttributeKey) {
                                case 'Vertical':
                                    if (self::identifyFixedStyleValue($verticalAlignmentStyles, $styleAttributeValue)) {
                                        $this->styles[$styleID]['alignment']['vertical'] = $styleAttributeValue;
                                    }
                                    break;
                                case 'Horizontal':
                                    if (self::identifyFixedStyleValue($horizontalAlignmentStyles, $styleAttributeValue)) {
                                        $this->styles[$styleID]['alignment']['horizontal'] = $styleAttributeValue;
                                    }
                                    break;
                                case 'WrapText':
                                    $this->styles[$styleID]['alignment']['wrap'] = \true;
                                    break;
                            }
                        }
                        break;
                    case 'Borders':
                        foreach ($styleData->Border as $borderStyle) {
                            $borderAttributes = $borderStyle->attributes($namespaces['ss']);
                            $thisBorder = array();
                            foreach ($borderAttributes as $borderStyleKey => $borderStyleValue) {
//                                    echo $borderStyleKey.' = '.$borderStyleValue.'<br />';
                                switch ($borderStyleKey) {
                                    case 'LineStyle':
                                        $thisBorder['style'] = \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_MEDIUM;
//                                                $thisBorder['style'] = $borderStyleValue;
                                        break;
                                    case 'Weight':
//                                                $thisBorder['style'] = $borderStyleValue;
                                        break;
                                    case 'Position':
                                        $borderPosition = \strtolower($borderStyleValue);
                                        break;
                                    case 'Color':
                                        $borderColour = \substr($borderStyleValue, 1);
                                        $thisBorder['color']['rgb'] = $borderColour;
                                        break;
                                }
                            }
                            if (!empty($thisBorder) && (($borderPosition == 'left') || ($borderPosition == 'right') || ($borderPosition == 'top') || ($borderPosition == 'bottom'))) {
                                $this->styles[$styleID]['borders'][$borderPosition] = $thisBorder;
                            }
                        }
                        break;
                    case 'Font':
                        foreach ($styleAttributes as $styleAttributeKey => $styleAttributeValue) {
//                                echo $styleAttributeKey.' = '.$styleAttributeValue.'<br />';
                            $styleAttributeValue = (string) $styleAttributeValue;
                            switch ($styleAttributeKey) {
                                case 'FontName':
                                    $this->styles[$styleID]['font']['name'] = $styleAttributeValue;
                                    break;
                                case 'Size':
                                    $this->styles[$styleID]['font']['size'] = $styleAttributeValue;
                                    break;
                                case 'Color':
                                    $this->styles[$styleID]['font']['color']['rgb'] = \substr($styleAttributeValue, 1);
                                    break;
                                case 'Bold':
                                    $this->styles[$styleID]['font']['bold'] = \true;
                                    break;
                                case 'Italic':
                                    $this->styles[$styleID]['font']['italic'] = \true;
                                    break;
                                case 'Underline':
                                    if (self::identifyFixedStyleValue($underlineStyles, $styleAttributeValue)) {
                                        $this->styles[$styleID]['font']['underline'] = $styleAttributeValue;
                                    }
                                    break;
                            }
                        }
                        break;
                    case 'Interior':
                        foreach ($styleAttributes as $styleAttributeKey => $styleAttributeValue) {
if ($styleAttributeKey === 'Color') {
                                $this->styles[$styleID]['fill']['color']['rgb'] = \substr($styleAttributeValue, 1);
                            }
                        }
                        break;
                    case 'NumberFormat':
                        foreach ($styleAttributes as $styleAttribute) {
//                                echo $styleAttributeKey.' = '.$styleAttributeValue.'<br />';
                            $styleAttribute = \str_replace($fromFormats, $toFormats, $styleAttribute);
                            if ($styleAttribute === 'Short Date') {
                                $styleAttribute = 'dd/mm/yyyy';
                            }
                            if ($styleAttribute > '') {
                                $this->styles[$styleID]['numberformat']['code'] = $styleAttribute;
                            }
                        }
                        break;
                    case 'Protection':
                        foreach ($styleAttributes as $styleAttribute) {
//                                echo $styleAttributeKey.' = '.$styleAttributeValue.'<br />';
                        }
                        break;
                }
            }
//            print_r($this->styles[$styleID]);
//            echo '<hr />';
        }
//        echo '<hr />';

        $worksheetID = 0;
        $xml_ss = $xml->children($namespaces['ss']);

        foreach ($xml_ss->Worksheet as $worksheet) {
            $worksheet_ss = $worksheet->attributes($namespaces['ss']);

            if (($this->loadSheetsOnly !== null) && (isset($worksheet_ss['Name'])) &&
                (!\in_array($worksheet_ss['Name'], $this->loadSheetsOnly))) {
                continue;
            }

//            echo '<h3>Worksheet: ', $worksheet_ss['Name'],'<h3>';
//
            // Create new Worksheet
            $phpExcel->createSheet();
            $phpExcel->setActiveSheetIndex($worksheetID);
            if (isset($worksheet_ss['Name'])) {
                $worksheetName = self::convertStringEncoding((string) $worksheet_ss['Name'], $this->charSet);
                //    Use false for $updateFormulaCellReferences to prevent adjustment of worksheet references in
                //        formula cells... during the load, all formulae should be correct, and we're simply bringing
                //        the worksheet name in line with the formula, not the reverse
                $phpExcel->getActiveSheet()->setTitle($worksheetName, \false);
            }

            $columnID = 'A';
            if (property_exists($worksheet->Table, 'Column') && $worksheet->Table->Column !== null) {
                foreach ($worksheet->Table->Column as $columnData) {
                    $columnData_ss = $columnData->attributes($namespaces['ss']);
                    if (isset($columnData_ss['Index'])) {
                        $columnID = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::stringFromColumnIndex($columnData_ss['Index']-1);
                    }
                    if (isset($columnData_ss['Width'])) {
                        $columnWidth = $columnData_ss['Width'];
//                        echo '<b>Setting column width for '.$columnID.' to '.$columnWidth.'</b><br />';
                        $phpExcel->getActiveSheet()->getColumnDimension($columnID, true)->setWidth($columnWidth / 5.4);
                    }
                    ++$columnID;
                }
            }

            $rowID = 1;
            if (property_exists($worksheet->Table, 'Row') && $worksheet->Table->Row !== null) {
                $additionalMergedCells = 0;
                foreach ($worksheet->Table->Row as $rowData) {
                    $rowHasData = \false;
                    $row_ss = $rowData->attributes($namespaces['ss']);
                    if (isset($row_ss['Index'])) {
                        $rowID = (integer) $row_ss['Index'];
                    }
//                    echo '<b>Row '.$rowID.'</b><br />';

                    $columnID = 'A';
                    foreach ($rowData->Cell as $cell) {
                        $cell_ss = $cell->attributes($namespaces['ss']);
                        if (isset($cell_ss['Index'])) {
                            $columnID = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::stringFromColumnIndex($cell_ss['Index']-1);
                        }
                        $cellRange = $columnID.$rowID;

                        if ($this->getReadFilter() !== \null && !$this->getReadFilter()->readCell($columnID, $rowID, $worksheetName)) {
                            continue;
                        }

                        if ((isset($cell_ss['MergeAcross'])) || (isset($cell_ss['MergeDown']))) {
                            $columnTo = $columnID;
                            if (isset($cell_ss['MergeAcross'])) {
                                $additionalMergedCells += (int)$cell_ss['MergeAcross'];
                                $columnTo = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::stringFromColumnIndex(\PhpOffice\PhpSpreadsheet\Cell\Coordinate::columnIndexFromString($columnID) + $cell_ss['MergeAcross'] -1);
                            }
                            $rowTo = $rowID;
                            if (isset($cell_ss['MergeDown'])) {
                                $rowTo += $cell_ss['MergeDown'];
                            }
                            $cellRange .= ':'.$columnTo.$rowTo;
                            $phpExcel->getActiveSheet()->mergeCells($cellRange);
                        }

                        $cellIsSet = $hasCalculatedValue = \false;
                        $cellDataFormula = '';
                        if (isset($cell_ss['Formula'])) {
                            $cellDataFormula = $cell_ss['Formula'];
                            // added this as a check for array formulas
                            if (isset($cell_ss['ArrayRange'])) {
                                $cellDataCSEFormula = $cell_ss['ArrayRange'];
//                                echo "found an array formula at ".$columnID.$rowID."<br />";
                            }
                            $hasCalculatedValue = \true;
                        }
                        if (property_exists($cell, 'Data') && $cell->Data !== null) {
                            $cellValue = $cellData = $cell->Data;
                            $type = \PhpOffice\PhpSpreadsheet\Cell\DataType::TYPE_NULL;
                            $cellData_ss = $cellData->attributes($namespaces['ss']);
                            if (isset($cellData_ss['Type'])) {
                                $cellDataType = $cellData_ss['Type'];
                                switch ($cellDataType) {
                                    /*
                                    const TYPE_STRING        = 's';
                                    const TYPE_FORMULA        = 'f';
                                    const TYPE_NUMERIC        = 'n';
                                    const TYPE_BOOL            = 'b';
                                    const TYPE_NULL            = 'null';
                                    const TYPE_INLINE        = 'inlineStr';
                                    const TYPE_ERROR        = 'e';
                                    */
                                    case 'String':
                                        $cellValue = self::convertStringEncoding($cellValue, $this->charSet);
                                        $type = \PhpOffice\PhpSpreadsheet\Cell\DataType::TYPE_STRING;
                                        break;
                                    case 'Number':
                                        $type = \PhpOffice\PhpSpreadsheet\Cell\DataType::TYPE_NUMERIC;
                                        $cellValue = (float) $cellValue;
                                        if (\floor($cellValue) === $cellValue) {
                                            $cellValue = (integer) $cellValue;
                                        }
                                        break;
                                    case 'Boolean':
                                        $type = \PhpOffice\PhpSpreadsheet\Cell\DataType::TYPE_BOOL;
                                        $cellValue = ($cellValue != 0);
                                        break;
                                    case 'DateTime':
                                        $type = \PhpOffice\PhpSpreadsheet\Cell\DataType::TYPE_NUMERIC;
                                        $cellValue = \PhpOffice\PhpSpreadsheet\Shared\Date::PHPToExcel(\strtotime($cellValue));
                                        break;
                                    case 'Error':
                                        $type = \PhpOffice\PhpSpreadsheet\Cell\DataType::TYPE_ERROR;
                                        break;
                                }
                            }

                            if ($hasCalculatedValue) {
//                                echo 'FORMULA<br />';
                                $type = \PhpOffice\PhpSpreadsheet\Cell\DataType::TYPE_FORMULA;
                                $columnNumber = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::columnIndexFromString($columnID);
                                if (\substr($cellDataFormula, 0, 3) == 'of:') {
                                    $cellDataFormula = \substr($cellDataFormula, 3);
//                                    echo 'Before: ', $cellDataFormula,'<br />';
                                    $temp = \explode('"', $cellDataFormula);
                                    $key = \false;
                                    foreach ($temp as &$singleTemp) {
                                        //    Only replace in alternate array entries (i.e. non-quoted blocks)
                                        if ($key = !$key) {
                                            $singleTemp = \str_replace(array('[.', '.', ']'), '', $singleTemp);
                                        }
                                    }
                                } else {
                                    //    Convert R1C1 style references to A1 style references (but only when not quoted)
//                                    echo 'Before: ', $cellDataFormula,'<br />';
                                    $temp = \explode('"', $cellDataFormula);
                                    $key = \false;
                                    foreach ($temp as &$singleTemp) {
                                        //    Only replace in alternate array entries (i.e. non-quoted blocks)
                                        if ($key = !$key) {
                                            \preg_match_all('/(R(\[?-?\d*\]?))(C(\[?-?\d*\]?))/', $singleTemp, $cellReferences, \PREG_SET_ORDER + \PREG_OFFSET_CAPTURE);
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
                                                    $rowReference = $rowID;
                                                }
                                                //    Bracketed R references are relative to the current row
                                                if ($rowReference{0} == '[') {
                                                    $rowReference = $rowID + \trim($rowReference, '[]');
                                                }
                                                $columnReference = $cellReference[4][0];
                                                //    Empty C reference is the current column
                                                if ($columnReference == '') {
                                                    $columnReference = $columnNumber;
                                                }
                                                //    Bracketed C references are relative to the current column
                                                if ($columnReference{0} == '[') {
                                                    $columnReference = $columnNumber + \trim($columnReference, '[]');
                                                }
                                                $A1CellReference = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::stringFromColumnIndex($columnReference-1).$rowReference;
                                                $singleTemp = \substr_replace($singleTemp, $A1CellReference, $cellReference[0][1], \strlen($cellReference[0][0]));
                                            }
                                        }
                                    }
                                }
                                unset($value);
                                //    Then rebuild the formula string
                                $cellDataFormula = \implode('"', $temp);
//                                echo 'After: ', $cellDataFormula,'<br />';
                            }

//                            echo 'Cell '.$columnID.$rowID.' is a '.$type.' with a value of '.(($hasCalculatedValue) ? $cellDataFormula : $cellValue).'<br />';
//
                            $phpExcel->getActiveSheet()->getCell($columnID.$rowID, true)->setValueExplicit((($hasCalculatedValue) ? $cellDataFormula : $cellValue), $type);
                            if ($hasCalculatedValue) {
//                                echo 'Formula result is '.$cellValue.'<br />';
                                $phpExcel->getActiveSheet()->getCell($columnID.$rowID, true)->setCalculatedValue($cellValue);
                            }
                            $cellIsSet = $rowHasData = \true;
                        }

                        if (property_exists($cell, 'Comment') && $cell->Comment !== null) {
//                            echo '<b>comment found</b><br />';
                            $commentAttributes = $cell->Comment->attributes($namespaces['ss']);
                            $author = 'unknown';
                            if (property_exists($commentAttributes, 'Author') && $commentAttributes->Author !== null) {
                                $author = (string)$commentAttributes->Author;
//                                echo 'Author: ', $author,'<br />';
                            }
                            $node = $cell->Comment->Data->asXML();
//                            $annotation = str_replace('html:','',substr($node,49,-10));
//                            echo $annotation,'<br />';
                            $annotation = \strip_tags($node);
//                            echo 'Annotation: ', $annotation,'<br />';
                            $phpExcel->getActiveSheet()->getComment($columnID.$rowID)->setAuthor(self::convertStringEncoding($author, $this->charSet))->setText($this->parseRichText($annotation));
                        }

                        if (($cellIsSet) && (isset($cell_ss['StyleID']))) {
                            $style = (string) $cell_ss['StyleID'];
//                            echo 'Cell style for '.$columnID.$rowID.' is '.$style.'<br />';
                            if ((isset($this->styles[$style])) && (!empty($this->styles[$style]))) {
//                                echo 'Cell '.$columnID.$rowID.'<br />';
//                                print_r($this->styles[$style]);
//                                echo '<br />';
                                if (!$phpExcel->getActiveSheet()->cellExists($columnID.$rowID)) {
                                    $phpExcel->getActiveSheet()->getCell($columnID.$rowID, true)->setValue(\null);
                                }
                                $phpExcel->getActiveSheet()->getStyle($cellRange)->applyFromArray($this->styles[$style], true);
                            }
                        }
                        ++$columnID;
                        while ($additionalMergedCells > 0) {
                            ++$columnID;
                            $additionalMergedCells--;
                        }
                    }

                    if ($rowHasData) {
                        if (isset($row_ss['StyleID'])) {
                            $rowStyle = $row_ss['StyleID'];
                        }
                        if (isset($row_ss['Height'])) {
                            $rowHeight = $row_ss['Height'];
//                            echo '<b>Setting row height to '.$rowHeight.'</b><br />';
                            $phpExcel->getActiveSheet()->getRowDimension($rowID, true)->setRowHeight($rowHeight);
                        }
                    }

                    ++$rowID;
                }
            }
            ++$worksheetID;
        }

        // Return
        return $phpExcel;
    }


    protected static function convertStringEncoding($string, $charset)
    {
        if ($charset != 'UTF-8') {
            return \PhpOffice\PhpSpreadsheet\Shared\StringHelper::ConvertEncoding($string, 'UTF-8', $charset);
        }
        return $string;
    }


    protected function parseRichText($is = '')
    {
        $phpExcelRichText = new \PhpOffice\PhpSpreadsheet\RichText\RichText();

        $phpExcelRichText->createText(self::convertStringEncoding($is, $this->charSet));

        return $phpExcelRichText;
    }
}

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
 * PHPExcel_Reader_OOCalc
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
class Ods extends \PhpOffice\PhpSpreadsheet\Reader\BaseReader implements \PhpOffice\PhpSpreadsheet\Reader\IReader
{
    /**
     * Formats
     *
     * @var array
     */
    private $styles = array();

    /**
     * Create a new PHPExcel_Reader_OOCalc
     */
    public function __construct()
    {
        $this->readFilter     = new \PhpOffice\PhpSpreadsheet\Reader\DefaultReadFilter();
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
        // Check if file exists
        if (!\file_exists($pFilename)) {
            throw new \PhpOffice\PhpSpreadsheet\Reader\Exception("Could not open " . $pFilename . " for reading! File does not exist.");
        }

        $zipClass = \PhpOffice\PhpSpreadsheet\Settings::getZipClass();

        // Check if zip class exists
//        if (!class_exists($zipClass, false)) {
//            throw new PHPExcel_Reader_Exception($zipClass . " library is not enabled");
//        }

        $mimeType = 'UNKNOWN';
        // Load file
        $zip = new $zipClass;
        if ($zip->open($pFilename) === \true) {
            // check if it is an OOXML archive
            $stat = $zip->statName('mimetype');
            if ($stat && ($stat['size'] <= 255)) {
                $mimeType = $zip->getFromName($stat['name']);
            } elseif ($stat = $zip->statName('META-INF/manifest.xml')) {
                $xml = \simplexml_load_string($this->securityScan($zip->getFromName('META-INF/manifest.xml')), 'SimpleXMLElement', \PhpOffice\PhpSpreadsheet\Settings::getLibXmlLoaderOptions());
                $namespacesContent = $xml->getNamespaces(\true);
                if (isset($namespacesContent['manifest'])) {
                    $manifest = $xml->children($namespacesContent['manifest']);
                    foreach ($manifest as $singleManifest) {
                        $manifestAttributes = $singleManifest->attributes($namespacesContent['manifest']);
                        if ($manifestAttributes->{'full-path'} == '/') {
                            $mimeType = (string) $manifestAttributes->{'media-type'};
                            break;
                        }
                    }
                }
            }

            $zip->close();

            return ($mimeType === 'application/vnd.oasis.opendocument.spreadsheet');
        }

        return \false;
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

        $zipClass = \PhpOffice\PhpSpreadsheet\Settings::getZipClass();

        $zip = new $zipClass;
        if (!$zip->open($pFilename)) {
            throw new \PhpOffice\PhpSpreadsheet\Reader\Exception("Could not open " . $pFilename . " for reading! Error opening file.");
        }

        $worksheetNames = array();

        $xmlReader = new \XMLReader();
        $res = $xmlReader->xml($this->securityScanFile('zip://'.\realpath($pFilename).'#content.xml'), \null, \PhpOffice\PhpSpreadsheet\Settings::getLibXmlLoaderOptions());
        $xmlReader->setParserProperty(2, \true);

        //    Step into the first level of content of the XML
        $xmlReader->read();
        while ($xmlReader->read()) {
            //    Quickly jump through to the office:body node
            while ($xmlReader->name !== 'office:body') {
                if ($xmlReader->isEmptyElement) {
                    $xmlReader->read();
                } else {
                    $xmlReader->next();
                }
            }
            //    Now read each node until we find our first table:table node
            while ($xmlReader->read()) {
                if ($xmlReader->name == 'table:table' && $xmlReader->nodeType == \XMLReader::ELEMENT) {
                    //    Loop through each table:table node reading the table:name attribute for each worksheet name
                    do {
                        $worksheetNames[] = $xmlReader->getAttribute('table:name');
                        $xmlReader->next();
                    } while ($xmlReader->name == 'table:table' && $xmlReader->nodeType == \XMLReader::ELEMENT);
                }
            }
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

        $zipClass = \PhpOffice\PhpSpreadsheet\Settings::getZipClass();

        $zip = new $zipClass;
        if (!$zip->open($pFilename)) {
            throw new \PhpOffice\PhpSpreadsheet\Reader\Exception("Could not open " . $pFilename . " for reading! Error opening file.");
        }

        $xmlReader = new \XMLReader();
        $res = $xmlReader->xml($this->securityScanFile('zip://'.\realpath($pFilename).'#content.xml'), \null, \PhpOffice\PhpSpreadsheet\Settings::getLibXmlLoaderOptions());
        $xmlReader->setParserProperty(2, \true);

        //    Step into the first level of content of the XML
        $xmlReader->read();
        while ($xmlReader->read()) {
            //    Quickly jump through to the office:body node
            while ($xmlReader->name !== 'office:body') {
                if ($xmlReader->isEmptyElement) {
                    $xmlReader->read();
                } else {
                    $xmlReader->next();
                }
            }
                //    Now read each node until we find our first table:table node
            while ($xmlReader->read()) {
                if ($xmlReader->name == 'table:table' && $xmlReader->nodeType == \XMLReader::ELEMENT) {
                    $worksheetNames[] = $xmlReader->getAttribute('table:name');

                    $tmpInfo = array(
                        'worksheetName' => $xmlReader->getAttribute('table:name'),
                        'lastColumnLetter' => 'A',
                        'lastColumnIndex' => 0,
                        'totalRows' => 0,
                        'totalColumns' => 0,
                    );

                    //    Loop through each child node of the table:table element reading
                    $currCells = 0;
                    do {
                        $xmlReader->read();
                        if ($xmlReader->name == 'table:table-row' && $xmlReader->nodeType == \XMLReader::ELEMENT) {
                            $rowspan = $xmlReader->getAttribute('table:number-rows-repeated');
                            $rowspan = empty($rowspan) ? 1 : $rowspan;
                            $tmpInfo['totalRows'] += $rowspan;
                            $tmpInfo['totalColumns'] = \max($tmpInfo['totalColumns'], $currCells);
                            $currCells = 0;
                            //    Step into the row
                            $xmlReader->read();
                            do {
                                if ($xmlReader->name == 'table:table-cell' && $xmlReader->nodeType == \XMLReader::ELEMENT) {
                                    if (!$xmlReader->isEmptyElement) {
                                        $currCells++;
                                        $xmlReader->next();
                                    } else {
                                        $xmlReader->read();
                                    }
                                } elseif ($xmlReader->name == 'table:covered-table-cell' && $xmlReader->nodeType == \XMLReader::ELEMENT) {
                                    $mergeSize = $xmlReader->getAttribute('table:number-columns-repeated');
                                    $currCells += $mergeSize;
                                    $xmlReader->read();
                                }
                            } while ($xmlReader->name != 'table:table-row');
                        }
                    } while ($xmlReader->name != 'table:table');

                    $tmpInfo['totalColumns'] = \max($tmpInfo['totalColumns'], $currCells);
                    $tmpInfo['lastColumnIndex'] = $tmpInfo['totalColumns'] - 1;
                    $tmpInfo['lastColumnLetter'] = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::stringFromColumnIndex($tmpInfo['lastColumnIndex']);
                    $worksheetInfo[] = $tmpInfo;
                }
            }

//                foreach ($workbookData->table as $worksheetDataSet) {
//                    $worksheetData = $worksheetDataSet->children($namespacesContent['table']);
//                    $worksheetDataAttributes = $worksheetDataSet->attributes($namespacesContent['table']);
//
//                    $rowIndex = 0;
//                    foreach ($worksheetData as $key => $rowData) {
//                        switch ($key) {
//                            case 'table-row' :
//                                $rowDataTableAttributes = $rowData->attributes($namespacesContent['table']);
//                                $rowRepeats = (isset($rowDataTableAttributes['number-rows-repeated'])) ?
//                                        $rowDataTableAttributes['number-rows-repeated'] : 1;
//                                $columnIndex = 0;
//
//                                foreach ($rowData as $key => $cellData) {
//                                    $cellDataTableAttributes = $cellData->attributes($namespacesContent['table']);
//                                    $colRepeats = (isset($cellDataTableAttributes['number-columns-repeated'])) ?
//                                        $cellDataTableAttributes['number-columns-repeated'] : 1;
//                                    $cellDataOfficeAttributes = $cellData->attributes($namespacesContent['office']);
//                                    if (isset($cellDataOfficeAttributes['value-type'])) {
//                                        $tmpInfo['lastColumnIndex'] = max($tmpInfo['lastColumnIndex'], $columnIndex + $colRepeats - 1);
//                                        $tmpInfo['totalRows'] = max($tmpInfo['totalRows'], $rowIndex + $rowRepeats);
//                                    }
//                                    $columnIndex += $colRepeats;
//                                }
//                                $rowIndex += $rowRepeats;
//                                break;
//                        }
//                    }
//
//                    $tmpInfo['lastColumnLetter'] = PHPExcel_Cell::stringFromColumnIndex($tmpInfo['lastColumnIndex']);
//                    $tmpInfo['totalColumns'] = $tmpInfo['lastColumnIndex'] + 1;
//
//                }
//            }
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
        // Check if file exists
        if (!\file_exists($pFilename)) {
            throw new \PhpOffice\PhpSpreadsheet\Reader\Exception("Could not open " . $pFilename . " for reading! File does not exist.");
        }

        $timezoneObj = new \DateTimeZone('Europe/London');
        $GMT = new \DateTimeZone('UTC');

        $zipClass = \PhpOffice\PhpSpreadsheet\Settings::getZipClass();

        $zip = new $zipClass;
        if (!$zip->open($pFilename)) {
            throw new \PhpOffice\PhpSpreadsheet\Reader\Exception("Could not open " . $pFilename . " for reading! Error opening file.");
        }

//        echo '<h1>Meta Information</h1>';
        $xml = \simplexml_load_string($this->securityScan($zip->getFromName("meta.xml")), 'SimpleXMLElement', \PhpOffice\PhpSpreadsheet\Settings::getLibXmlLoaderOptions());
        $namespacesMeta = $xml->getNamespaces(\true);
//        echo '<pre>';
//        print_r($namespacesMeta);
//        echo '</pre><hr />';

        $docProps = $phpExcel->getProperties();
        $officeProperty = $xml->children($namespacesMeta['office']);
        foreach ($officeProperty as $singleOfficeProperty) {
            $officePropertyDC = array();
            if (isset($namespacesMeta['dc'])) {
                $officePropertyDC = $singleOfficeProperty->children($namespacesMeta['dc']);
            }
            foreach ($officePropertyDC as $propertyName => $propertyValue) {
                $propertyValue = (string) $propertyValue;
                switch ($propertyName) {
                    case 'title':
                        $docProps->setTitle($propertyValue);
                        break;
                    case 'subject':
                        $docProps->setSubject($propertyValue);
                        break;
                    case 'creator':
                        $docProps->setCreator($propertyValue);
                        $docProps->setLastModifiedBy($propertyValue);
                        break;
                    case 'date':
                        $creationDate = \strtotime($propertyValue);
                        $docProps->setCreated($creationDate);
                        $docProps->setModified($creationDate);
                        break;
                    case 'description':
                        $docProps->setDescription($propertyValue);
                        break;
                }
            }
            $officePropertyMeta = array();
            if (isset($namespacesMeta['dc'])) {
                $officePropertyMeta = $singleOfficeProperty->children($namespacesMeta['meta']);
            }
            foreach ($officePropertyMeta as $propertyName => $propertyValue) {
                $propertyValueAttributes = $propertyValue->attributes($namespacesMeta['meta']);
                $propertyValue = (string) $propertyValue;
                switch ($propertyName) {
                    case 'initial-creator':
                        $docProps->setCreator($propertyValue);
                        break;
                    case 'keyword':
                        $docProps->setKeywords($propertyValue);
                        break;
                    case 'creation-date':
                        $creationDate = \strtotime($propertyValue);
                        $docProps->setCreated($creationDate);
                        break;
                    case 'user-defined':
                        $propertyValueType = \PhpOffice\PhpSpreadsheet\Document\Properties::PROPERTY_TYPE_STRING;
                        foreach ($propertyValueAttributes as $key => $value) {
                            if ($key == 'name') {
                                $propertyValueName = (string) $value;
                            } elseif ($key == 'value-type') {
                                switch ($value) {
                                    case 'date':
                                        $propertyValue = \PhpOffice\PhpSpreadsheet\Document\Properties::convertProperty($propertyValue, 'date');
                                        $propertyValueType = \PhpOffice\PhpSpreadsheet\Document\Properties::PROPERTY_TYPE_DATE;
                                        break;
                                    case 'boolean':
                                        $propertyValue = \PhpOffice\PhpSpreadsheet\Document\Properties::convertProperty($propertyValue, 'bool');
                                        $propertyValueType = \PhpOffice\PhpSpreadsheet\Document\Properties::PROPERTY_TYPE_BOOLEAN;
                                        break;
                                    case 'float':
                                        $propertyValue = \PhpOffice\PhpSpreadsheet\Document\Properties::convertProperty($propertyValue, 'r4');
                                        $propertyValueType = \PhpOffice\PhpSpreadsheet\Document\Properties::PROPERTY_TYPE_FLOAT;
                                        break;
                                    default:
                                        $propertyValueType = \PhpOffice\PhpSpreadsheet\Document\Properties::PROPERTY_TYPE_STRING;
                                }
                            }
                        }
                        $docProps->setCustomProperty($propertyValueName, $propertyValue, $propertyValueType);
                        break;
                }
            }
        }


//        echo '<h1>Workbook Content</h1>';
        $xml = \simplexml_load_string($this->securityScan($zip->getFromName("content.xml")), 'SimpleXMLElement', \PhpOffice\PhpSpreadsheet\Settings::getLibXmlLoaderOptions());
        $namespacesContent = $xml->getNamespaces(\true);
//        echo '<pre>';
//        print_r($namespacesContent);
//        echo '</pre><hr />';

        $workbook = $xml->children($namespacesContent['office']);
        foreach ($workbook->body->spreadsheet as $workbookData) {
            $workbookData = $workbookData->children($namespacesContent['table']);
            $worksheetID = 0;
            foreach ($workbookData->table as $worksheetDataSet) {
                $worksheetData = $worksheetDataSet->children($namespacesContent['table']);
//                print_r($worksheetData);
//                echo '<br />';
                $worksheetDataAttributes = $worksheetDataSet->attributes($namespacesContent['table']);
//                print_r($worksheetDataAttributes);
//                echo '<br />';
                if (($this->loadSheetsOnly !== null) && (isset($worksheetDataAttributes['name'])) &&
                    (!\in_array($worksheetDataAttributes['name'], $this->loadSheetsOnly))) {
                    continue;
                }

//                echo '<h2>Worksheet '.$worksheetDataAttributes['name'].'</h2>';
                // Create new Worksheet
                $phpExcel->createSheet();
                $phpExcel->setActiveSheetIndex($worksheetID);
                if (isset($worksheetDataAttributes['name'])) {
                    $worksheetName = (string) $worksheetDataAttributes['name'];
                    //    Use false for $updateFormulaCellReferences to prevent adjustment of worksheet references in
                    //        formula cells... during the load, all formulae should be correct, and we're simply
                    //        bringing the worksheet name in line with the formula, not the reverse
                    $phpExcel->getActiveSheet()->setTitle($worksheetName, \false);
                }

                $rowID = 1;
                foreach ($worksheetData as $key => $rowData) {
//                    echo '<b>'.$key.'</b><br />';
                    switch ($key) {
                        case 'table-header-rows':
                            foreach ($rowData as $singleRowData) {
                                $rowData = $singleRowData;
                                break;
                            }
                        case 'table-row':
                            $rowDataTableAttributes = $rowData->attributes($namespacesContent['table']);
                            $rowRepeats = (isset($rowDataTableAttributes['number-rows-repeated'])) ? $rowDataTableAttributes['number-rows-repeated'] : 1;
                            $columnID = 'A';
                            foreach ($rowData as $singleRowData) {
                                if ($this->getReadFilter() !== \null && !$this->getReadFilter()->readCell($columnID, $rowID, $worksheetName)) {
                                    continue;
                                }

//                                echo '<b>'.$columnID.$rowID.'</b><br />';
                                $cellDataText = (isset($namespacesContent['text'])) ? $singleRowData->children($namespacesContent['text']) : '';
                                $cellDataOffice = $singleRowData->children($namespacesContent['office']);
                                $cellDataOfficeAttributes = $singleRowData->attributes($namespacesContent['office']);
                                $cellDataTableAttributes = $singleRowData->attributes($namespacesContent['table']);

//                                echo 'Office Attributes: ';
//                                print_r($cellDataOfficeAttributes);
//                                echo '<br />Table Attributes: ';
//                                print_r($cellDataTableAttributes);
//                                echo '<br />Cell Data Text';
//                                print_r($cellDataText);
//                                echo '<br />';
//
                                $type = $formatting = $hyperlink = \null;
                                $hasCalculatedValue = \false;
                                $cellDataFormula = '';
                                if (isset($cellDataTableAttributes['formula'])) {
                                    $cellDataFormula = $cellDataTableAttributes['formula'];
                                    $hasCalculatedValue = \true;
                                }

                                if (property_exists($cellDataOffice, 'annotation') && $cellDataOffice->annotation !== null) {
//                                    echo 'Cell has comment<br />';
                                    $annotationText = $cellDataOffice->annotation->children($namespacesContent['text']);
                                    $textArray = array();
                                    foreach ($annotationText as $singleAnnotationText) {
                                        if (property_exists($singleAnnotationText, 'span') && $singleAnnotationText->span !== null) {
                                            foreach ($singleAnnotationText->span as $text) {
                                                $textArray[] = (string)$text;
                                            }
                                        } else {
                                            $textArray[] = (string) $singleAnnotationText;
                                        }
                                    }
                                    $text = \implode("\n", $textArray);
//                                    echo $text, '<br />';
                                    $phpExcel->getActiveSheet()->getComment($columnID.$rowID)->setText($this->parseRichText($text));
//                                                                    ->setAuthor( $author )
                                }

                                if (property_exists($cellDataText, 'p') && $cellDataText->p !== null) {
                                    // Consolidate if there are multiple p records (maybe with spans as well)
                                    $dataArray = array();
                                    // Text can have multiple text:p and within those, multiple text:span.
                                    // text:p newlines, but text:span does not.
                                    // Also, here we assume there is no text data is span fields are specified, since
                                    // we have no way of knowing proper positioning anyway.
                                    foreach ($cellDataText->p as $pData) {
                                        if (property_exists($pData, 'span') && $pData->span !== null) {
                                            // span sections do not newline, so we just create one large string here
                                            $spanSection = "";
                                            foreach ($pData->span as $spanData) {
                                                $spanSection .= $spanData;
                                            }
                                            $dataArray[] = $spanSection;
                                        } else {
                                            $dataArray[] = $pData;
                                        }
                                    }
                                    $allCellDataText = \implode($dataArray, "\n");

//                                    echo 'Value Type is '.$cellDataOfficeAttributes['value-type'].'<br />';
                                    switch ($cellDataOfficeAttributes['value-type']) {
                                        case 'string':
                                            $type = \PhpOffice\PhpSpreadsheet\Cell\DataType::TYPE_STRING;
                                            $dataValue = $allCellDataText;
                                            if (property_exists($dataValue, 'a') && $dataValue->a !== null) {
                                                $dataValue = $dataValue->a;
                                                $cellXLinkAttributes = $dataValue->attributes($namespacesContent['xlink']);
                                                $hyperlink = $cellXLinkAttributes['href'];
                                            }
                                            break;
                                        case 'boolean':
                                            $type = \PhpOffice\PhpSpreadsheet\Cell\DataType::TYPE_BOOL;
                                            $dataValue = $allCellDataText == 'TRUE';
                                            break;
                                        case 'percentage':
                                            $type = \PhpOffice\PhpSpreadsheet\Cell\DataType::TYPE_NUMERIC;
                                            $dataValue = (float) $cellDataOfficeAttributes['value'];
                                            if (\floor($dataValue) === $dataValue) {
                                                $dataValue = (integer) $dataValue;
                                            }
                                            $formatting = \PhpOffice\PhpSpreadsheet\Style\NumberFormat::FORMAT_PERCENTAGE_00;
                                            break;
                                        case 'currency':
                                            $type = \PhpOffice\PhpSpreadsheet\Cell\DataType::TYPE_NUMERIC;
                                            $dataValue = (float) $cellDataOfficeAttributes['value'];
                                            if (\floor($dataValue) === $dataValue) {
                                                $dataValue = (integer) $dataValue;
                                            }
                                            $formatting = \PhpOffice\PhpSpreadsheet\Style\NumberFormat::FORMAT_CURRENCY_USD_SIMPLE;
                                            break;
                                        case 'float':
                                            $type = \PhpOffice\PhpSpreadsheet\Cell\DataType::TYPE_NUMERIC;
                                            $dataValue = (float) $cellDataOfficeAttributes['value'];
                                            if (\floor($dataValue) === $dataValue) {
                                                $dataValue = $dataValue == (integer) $dataValue ? (integer) $dataValue : $dataValue;
                                            }
                                            break;
                                        case 'date':
                                            $type = \PhpOffice\PhpSpreadsheet\Cell\DataType::TYPE_NUMERIC;
                                            $dateObj = new \DateTime($cellDataOfficeAttributes['date-value'], $GMT);
                                            $dateObj->setTimeZone($timezoneObj);
                                            list($year, $month, $day, $hour, $minute, $second) = \explode(' ', $dateObj->format('Y m d H i s'));
                                            $dataValue = \PhpOffice\PhpSpreadsheet\Shared\Date::formattedPHPToExcel($year, $month, $day, $hour, $minute, $second);
                                            if ($dataValue != \floor($dataValue)) {
                                                $formatting = \PhpOffice\PhpSpreadsheet\Style\NumberFormat::FORMAT_DATE_XLSX15.' '.\PhpOffice\PhpSpreadsheet\Style\NumberFormat::FORMAT_DATE_TIME4;
                                            } else {
                                                $formatting = \PhpOffice\PhpSpreadsheet\Style\NumberFormat::FORMAT_DATE_XLSX15;
                                            }
                                            break;
                                        case 'time':
                                            $type = \PhpOffice\PhpSpreadsheet\Cell\DataType::TYPE_NUMERIC;
                                            $dataValue = \PhpOffice\PhpSpreadsheet\Shared\Date::PHPToExcel(\strtotime('01-01-1970 '.\implode(':', \sscanf($cellDataOfficeAttributes['time-value'], 'PT%dH%dM%dS'))));
                                            $formatting = \PhpOffice\PhpSpreadsheet\Style\NumberFormat::FORMAT_DATE_TIME4;
                                            break;
                                    }
//                                    echo 'Data value is '.$dataValue.'<br />';
//                                    if ($hyperlink !== null) {
//                                        echo 'Hyperlink is '.$hyperlink.'<br />';
//                                    }
                                } else {
                                    $type = \PhpOffice\PhpSpreadsheet\Cell\DataType::TYPE_NULL;
                                    $dataValue = \null;
                                }

                                if ($hasCalculatedValue) {
                                    $type = \PhpOffice\PhpSpreadsheet\Cell\DataType::TYPE_FORMULA;
//                                    echo 'Formula: ', $cellDataFormula, PHP_EOL;
                                    $cellDataFormula = \substr($cellDataFormula, \strpos($cellDataFormula, ':=')+1);
                                    $temp = \explode('"', $cellDataFormula);
                                    $tKey = \false;
                                    foreach ($temp as &$singleTemp) {
                                        //    Only replace in alternate array entries (i.e. non-quoted blocks)
                                        if ($tKey = !$tKey) {
                                            $singleTemp = \preg_replace('/\[([^\.]+)\.([^\.]+):\.([^\.]+)\]/Ui', '$1!$2:$3', $singleTemp);    //  Cell range reference in another sheet
                                            $singleTemp = \preg_replace('/\[([^\.]+)\.([^\.]+)\]/Ui', '$1!$2', $singleTemp);       //  Cell reference in another sheet
                                            $singleTemp = \preg_replace('/\[\.([^\.]+):\.([^\.]+)\]/Ui', '$1:$2', $singleTemp);    //  Cell range reference
                                            $singleTemp = \preg_replace('/\[\.([^\.]+)\]/Ui', '$1', $singleTemp);                  //  Simple cell reference
                                            $singleTemp = \PhpOffice\PhpSpreadsheet\Calculation\Calculation::translateSeparator(';', ',', $singleTemp, $inBraces);
                                        }
                                    }
                                    unset($value);
                                    //    Then rebuild the formula string
                                    $cellDataFormula = \implode('"', $temp);
//                                    echo 'Adjusted Formula: ', $cellDataFormula, PHP_EOL;
                                }

                                $colRepeats = (isset($cellDataTableAttributes['number-columns-repeated'])) ? $cellDataTableAttributes['number-columns-repeated'] : 1;
                                if ($type !== \null) {
                                    for ($i = 0; $i < $colRepeats; ++$i) {
                                        if ($i > 0) {
                                            ++$columnID;
                                        }
                                        if ($type !== \PhpOffice\PhpSpreadsheet\Cell\DataType::TYPE_NULL) {
                                            for ($rowAdjust = 0; $rowAdjust < $rowRepeats; ++$rowAdjust) {
                                                $rID = $rowID + $rowAdjust;
                                                $phpExcel->getActiveSheet()->getCell($columnID.$rID, true)->setValueExplicit((($hasCalculatedValue) ? $cellDataFormula : $dataValue), $type);
                                                if ($hasCalculatedValue) {
//                                                    echo 'Forumla result is '.$dataValue.'<br />';
                                                    $phpExcel->getActiveSheet()->getCell($columnID.$rID, true)->setCalculatedValue($dataValue);
                                                }
                                                if ($formatting !== \null) {
                                                    $phpExcel->getActiveSheet()->getStyle($columnID.$rID)->getNumberFormat()->setFormatCode($formatting);
                                                } else {
                                                    $phpExcel->getActiveSheet()->getStyle($columnID.$rID)->getNumberFormat()->setFormatCode(\PhpOffice\PhpSpreadsheet\Style\NumberFormat::FORMAT_GENERAL);
                                                }
                                                if ($hyperlink !== \null) {
                                                    $phpExcel->getActiveSheet()->getCell($columnID.$rID, true)->getHyperlink()->setUrl($hyperlink);
                                                }
                                            }
                                        }
                                    }
                                }

                                //    Merged cells
                                if (((isset($cellDataTableAttributes['number-columns-spanned'])) || (isset($cellDataTableAttributes['number-rows-spanned']))) && (($type !== \PhpOffice\PhpSpreadsheet\Cell\DataType::TYPE_NULL) || (!$this->readDataOnly))) {
                                    $columnTo = $columnID;
                                    if (isset($cellDataTableAttributes['number-columns-spanned'])) {
                                        $columnTo = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::stringFromColumnIndex(\PhpOffice\PhpSpreadsheet\Cell\Coordinate::columnIndexFromString($columnID) + $cellDataTableAttributes['number-columns-spanned'] -2);
                                    }
                                    $rowTo = $rowID;
                                    if (isset($cellDataTableAttributes['number-rows-spanned'])) {
                                        $rowTo = $rowTo + $cellDataTableAttributes['number-rows-spanned'] - 1;
                                    }
                                    $cellRange = $columnID.$rowID.':'.$columnTo.$rowTo;
                                    $phpExcel->getActiveSheet()->mergeCells($cellRange);
                                }

                                ++$columnID;
                            }
                            $rowID += $rowRepeats;
                            break;
                    }
                }
                ++$worksheetID;
            }
        }

        // Return
        return $phpExcel;
    }

    private function parseRichText($is = '')
    {
        $phpExcelRichText = new \PhpOffice\PhpSpreadsheet\RichText\RichText();

        $phpExcelRichText->createText($is);

        return $phpExcelRichText;
    }
}

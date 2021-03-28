<?php

namespace PhpOffice\PhpSpreadsheet\Writer\Xlsx;

/**
 * PHPExcel_Writer_Excel2007_Workbook
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
 * @package    PHPExcel_Writer_Excel2007
 * @copyright  Copyright (c) 2006 - 2015 PHPExcel (http://www.codeplex.com/PHPExcel)
 * @license    http://www.gnu.org/licenses/old-licenses/lgpl-2.1.txt    LGPL
 * @version    ##VERSION##, ##DATE##
 */
class Workbook extends \PhpOffice\PhpSpreadsheet\Writer\Xlsx\WriterPart
{
    /**
     * Write workbook to XML format
     *
     * @param    boolean        $recalcRequired    Indicate whether formulas should be recalculated before writing
     * @return     string         XML Output
     * @throws     \PhpOffice\PhpSpreadsheet\Writer\Exception
     */
    public function writeWorkbook(\PhpOffice\PhpSpreadsheet\Spreadsheet $phpExcel = \null, $recalcRequired = \false)
    {
        // Create XML writer
        $objWriter = \null;
        if ($this->getParentWriter()->getUseDiskCaching()) {
            $objWriter = new \PhpOffice\PhpSpreadsheet\Shared\XMLWriter(\PhpOffice\PhpSpreadsheet\Shared\XMLWriter::STORAGE_DISK, $this->getParentWriter()->getDiskCachingDirectory());
        } else {
            $objWriter = new \PhpOffice\PhpSpreadsheet\Shared\XMLWriter(\PhpOffice\PhpSpreadsheet\Shared\XMLWriter::STORAGE_MEMORY);
        }

        // XML header
        $objWriter->startDocument('1.0', 'UTF-8', 'yes');

        // workbook
        $objWriter->startElement('workbook');
        $objWriter->writeAttribute('xml:space', 'preserve');
        $objWriter->writeAttribute('xmlns', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main');
        $objWriter->writeAttribute('xmlns:r', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships');

        // fileVersion
        $this->writeFileVersion($objWriter);

        // workbookPr
        $this->writeWorkbookPr($objWriter);

        // workbookProtection
        $this->writeWorkbookProtection($objWriter, $phpExcel);

        // bookViews
        if ($this->getParentWriter()->getOffice2003Compatibility() === \false) {
            $this->writeBookViews($objWriter, $phpExcel);
        }

        // sheets
        $this->writeSheets($objWriter, $phpExcel);

        // definedNames
        $this->writeDefinedNames($objWriter, $phpExcel);

        // calcPr
        $this->writeCalcPr($objWriter, $recalcRequired);

        $objWriter->endElement();

        // Return
        return $objWriter->getData();
    }

    /**
     * Write file version
     *
     * @param \PhpOffice\PhpSpreadsheet\Shared\XMLWriter $phpExcelSharedXMLWriter         XML Writer
     * @throws     \PhpOffice\PhpSpreadsheet\Writer\Exception
     */
    private function writeFileVersion(\PhpOffice\PhpSpreadsheet\Shared\XMLWriter $phpExcelSharedXMLWriter = \null)
    {
        $phpExcelSharedXMLWriter->startElement('fileVersion');
        $phpExcelSharedXMLWriter->writeAttribute('appName', 'xl');
        $phpExcelSharedXMLWriter->writeAttribute('lastEdited', '4');
        $phpExcelSharedXMLWriter->writeAttribute('lowestEdited', '4');
        $phpExcelSharedXMLWriter->writeAttribute('rupBuild', '4505');
        $phpExcelSharedXMLWriter->endElement();
    }

    /**
     * Write WorkbookPr
     *
     * @param \PhpOffice\PhpSpreadsheet\Shared\XMLWriter $phpExcelSharedXMLWriter         XML Writer
     * @throws     \PhpOffice\PhpSpreadsheet\Writer\Exception
     */
    private function writeWorkbookPr(\PhpOffice\PhpSpreadsheet\Shared\XMLWriter $phpExcelSharedXMLWriter = \null)
    {
        $phpExcelSharedXMLWriter->startElement('workbookPr');

        if (\PhpOffice\PhpSpreadsheet\Shared\Date::getExcelCalendar() == \PhpOffice\PhpSpreadsheet\Shared\Date::CALENDAR_MAC_1904) {
            $phpExcelSharedXMLWriter->writeAttribute('date1904', '1');
        }

        $phpExcelSharedXMLWriter->writeAttribute('codeName', 'ThisWorkbook');

        $phpExcelSharedXMLWriter->endElement();
    }

    /**
     * Write BookViews
     *
     * @param \PhpOffice\PhpSpreadsheet\Shared\XMLWriter     $phpExcelSharedXMLWriter         XML Writer
     * @throws     \PhpOffice\PhpSpreadsheet\Writer\Exception
     */
    private function writeBookViews(\PhpOffice\PhpSpreadsheet\Shared\XMLWriter $phpExcelSharedXMLWriter = \null, \PhpOffice\PhpSpreadsheet\Spreadsheet $phpExcel = \null)
    {
        // bookViews
        $phpExcelSharedXMLWriter->startElement('bookViews');

        // workbookView
        $phpExcelSharedXMLWriter->startElement('workbookView');

        $phpExcelSharedXMLWriter->writeAttribute('activeTab', $phpExcel->getActiveSheetIndex());
        $phpExcelSharedXMLWriter->writeAttribute('autoFilterDateGrouping', '1');
        $phpExcelSharedXMLWriter->writeAttribute('firstSheet', '0');
        $phpExcelSharedXMLWriter->writeAttribute('minimized', '0');
        $phpExcelSharedXMLWriter->writeAttribute('showHorizontalScroll', '1');
        $phpExcelSharedXMLWriter->writeAttribute('showSheetTabs', '1');
        $phpExcelSharedXMLWriter->writeAttribute('showVerticalScroll', '1');
        $phpExcelSharedXMLWriter->writeAttribute('tabRatio', '600');
        $phpExcelSharedXMLWriter->writeAttribute('visibility', 'visible');

        $phpExcelSharedXMLWriter->endElement();

        $phpExcelSharedXMLWriter->endElement();
    }

    /**
     * Write WorkbookProtection
     *
     * @param \PhpOffice\PhpSpreadsheet\Shared\XMLWriter     $phpExcelSharedXMLWriter         XML Writer
     * @throws     \PhpOffice\PhpSpreadsheet\Writer\Exception
     */
    private function writeWorkbookProtection(\PhpOffice\PhpSpreadsheet\Shared\XMLWriter $phpExcelSharedXMLWriter = \null, \PhpOffice\PhpSpreadsheet\Spreadsheet $phpExcel = \null)
    {
        if ($phpExcel->getSecurity()->isSecurityEnabled()) {
            $phpExcelSharedXMLWriter->startElement('workbookProtection');
            $phpExcelSharedXMLWriter->writeAttribute('lockRevision', ($phpExcel->getSecurity()->getLockRevision() ? 'true' : 'false'));
            $phpExcelSharedXMLWriter->writeAttribute('lockStructure', ($phpExcel->getSecurity()->getLockStructure() ? 'true' : 'false'));
            $phpExcelSharedXMLWriter->writeAttribute('lockWindows', ($phpExcel->getSecurity()->getLockWindows() ? 'true' : 'false'));

            if ($phpExcel->getSecurity()->getRevisionsPassword() != '') {
                $phpExcelSharedXMLWriter->writeAttribute('revisionsPassword', $phpExcel->getSecurity()->getRevisionsPassword());
            }

            if ($phpExcel->getSecurity()->getWorkbookPassword() != '') {
                $phpExcelSharedXMLWriter->writeAttribute('workbookPassword', $phpExcel->getSecurity()->getWorkbookPassword());
            }

            $phpExcelSharedXMLWriter->endElement();
        }
    }

    /**
     * Write calcPr
     *
     * @param \PhpOffice\PhpSpreadsheet\Shared\XMLWriter    $phpExcelSharedXMLWriter        XML Writer
     * @param    boolean                        $recalcRequired    Indicate whether formulas should be recalculated before writing
     * @throws     \PhpOffice\PhpSpreadsheet\Writer\Exception
     */
    private function writeCalcPr(\PhpOffice\PhpSpreadsheet\Shared\XMLWriter $phpExcelSharedXMLWriter = \null, $recalcRequired = \true)
    {
        $phpExcelSharedXMLWriter->startElement('calcPr');

        //    Set the calcid to a higher value than Excel itself will use, otherwise Excel will always recalc
        //  If MS Excel does do a recalc, then users opening a file in MS Excel will be prompted to save on exit
        //     because the file has changed
        $phpExcelSharedXMLWriter->writeAttribute('calcId', '999999');
        $phpExcelSharedXMLWriter->writeAttribute('calcMode', 'auto');
        //    fullCalcOnLoad isn't needed if we've recalculating for the save
        $phpExcelSharedXMLWriter->writeAttribute('calcCompleted', ($recalcRequired) ? 1 : 0);
        $phpExcelSharedXMLWriter->writeAttribute('fullCalcOnLoad', ($recalcRequired) ? 0 : 1);

        $phpExcelSharedXMLWriter->endElement();
    }

    /**
     * Write sheets
     *
     * @param \PhpOffice\PhpSpreadsheet\Shared\XMLWriter     $phpExcelSharedXMLWriter         XML Writer
     * @throws     \PhpOffice\PhpSpreadsheet\Writer\Exception
     */
    private function writeSheets(\PhpOffice\PhpSpreadsheet\Shared\XMLWriter $phpExcelSharedXMLWriter = \null, \PhpOffice\PhpSpreadsheet\Spreadsheet $phpExcel = \null)
    {
        // Write sheets
        $phpExcelSharedXMLWriter->startElement('sheets');
        $sheetCount = $phpExcel->getSheetCount();
        for ($i = 0; $i < $sheetCount; ++$i) {
            // sheet
            $this->writeSheet(
                $phpExcelSharedXMLWriter,
                $phpExcel->getSheet($i)->getTitle(),
                ($i + 1),
                ($i + 1 + 3),
                $phpExcel->getSheet($i)->getSheetState()
            );
        }

        $phpExcelSharedXMLWriter->endElement();
    }

    /**
     * Write sheet
     *
     * @param \PhpOffice\PhpSpreadsheet\Shared\XMLWriter     $phpExcelSharedXMLWriter         XML Writer
     * @param     string                         $pSheetname         Sheet name
     * @param     int                            $pSheetId             Sheet id
     * @param     int                            $pRelId                Relationship ID
     * @param   string                      $sheetState         Sheet state (visible, hidden, veryHidden)
     * @throws     \PhpOffice\PhpSpreadsheet\Writer\Exception
     */
    private function writeSheet(\PhpOffice\PhpSpreadsheet\Shared\XMLWriter $phpExcelSharedXMLWriter = \null, $pSheetname = '', $pSheetId = 1, $pRelId = 1, $sheetState = 'visible')
    {
        if ($pSheetname != '') {
            // Write sheet
            $phpExcelSharedXMLWriter->startElement('sheet');
            $phpExcelSharedXMLWriter->writeAttribute('name', $pSheetname);
            $phpExcelSharedXMLWriter->writeAttribute('sheetId', $pSheetId);
            if ($sheetState != 'visible' && $sheetState != '') {
                $phpExcelSharedXMLWriter->writeAttribute('state', $sheetState);
            }
            $phpExcelSharedXMLWriter->writeAttribute('r:id', 'rId' . $pRelId);
            $phpExcelSharedXMLWriter->endElement();
        } else {
            throw new \PhpOffice\PhpSpreadsheet\Writer\Exception("Invalid parameters passed.");
        }
    }

    /**
     * Write Defined Names
     *
     * @param \PhpOffice\PhpSpreadsheet\Shared\XMLWriter    $phpExcelSharedXMLWriter         XML Writer
     * @throws     \PhpOffice\PhpSpreadsheet\Writer\Exception
     */
    private function writeDefinedNames(\PhpOffice\PhpSpreadsheet\Shared\XMLWriter $phpExcelSharedXMLWriter = \null, \PhpOffice\PhpSpreadsheet\Spreadsheet $phpExcel = \null)
    {
        // Write defined names
        $phpExcelSharedXMLWriter->startElement('definedNames');

        // Named ranges
        if (\count($phpExcel->getNamedRanges()) > 0) {
            // Named ranges
            $this->writeNamedRanges($phpExcelSharedXMLWriter, $phpExcel);
        }

        // Other defined names
        $sheetCount = $phpExcel->getSheetCount();
        for ($i = 0; $i < $sheetCount; ++$i) {
            // definedName for autoFilter
            $this->writeDefinedNameForAutofilter($phpExcelSharedXMLWriter, $phpExcel->getSheet($i), $i);

            // definedName for Print_Titles
            $this->writeDefinedNameForPrintTitles($phpExcelSharedXMLWriter, $phpExcel->getSheet($i), $i);

            // definedName for Print_Area
            $this->writeDefinedNameForPrintArea($phpExcelSharedXMLWriter, $phpExcel->getSheet($i), $i);
        }

        $phpExcelSharedXMLWriter->endElement();
    }

    /**
     * Write named ranges
     *
     * @param \PhpOffice\PhpSpreadsheet\Shared\XMLWriter    $phpExcelSharedXMLWriter         XML Writer
     * @throws     \PhpOffice\PhpSpreadsheet\Writer\Exception
     */
    private function writeNamedRanges(\PhpOffice\PhpSpreadsheet\Shared\XMLWriter $phpExcelSharedXMLWriter = \null, \PhpOffice\PhpSpreadsheet\Spreadsheet $phpExcel)
    {
        // Loop named ranges
        $namedRanges = $phpExcel->getNamedRanges();
        foreach ($namedRanges as $namedRange) {
            $this->writeDefinedNameForNamedRange($phpExcelSharedXMLWriter, $namedRange);
        }
    }

    /**
     * Write Defined Name for named range
     *
     * @param \PhpOffice\PhpSpreadsheet\Shared\XMLWriter    $phpExcelSharedXMLWriter         XML Writer
     * @throws     \PhpOffice\PhpSpreadsheet\Writer\Exception
     */
    private function writeDefinedNameForNamedRange(\PhpOffice\PhpSpreadsheet\Shared\XMLWriter $phpExcelSharedXMLWriter = \null, \PhpOffice\PhpSpreadsheet\NamedRange $phpExcelNamedRange)
    {
        // definedName for named range
        $phpExcelSharedXMLWriter->startElement('definedName');
        $phpExcelSharedXMLWriter->writeAttribute('name', $phpExcelNamedRange->getName());
        if ($phpExcelNamedRange->isLocalOnly()) {
            $phpExcelSharedXMLWriter->writeAttribute('localSheetId', $phpExcelNamedRange->getScope()->getParent()->getIndex($phpExcelNamedRange->getScope()));
        }

        // Create absolute coordinate and write as raw text
        $range = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::splitRange($phpExcelNamedRange->getRange());
        foreach ($range as $i => $singleRange) {
            $range[$i][0] = '\'' . \str_replace("'", "''", $phpExcelNamedRange->getWorksheet()->getTitle()) . '\'!' . \PhpOffice\PhpSpreadsheet\Cell\Coordinate::absoluteReference($singleRange[0]);
            if (isset($singleRange[1])) {
                $range[$i][1] = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::absoluteReference($singleRange[1]);
            }
        }
        $range = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::buildRange($range);

        $phpExcelSharedXMLWriter->writeRawData($range);

        $phpExcelSharedXMLWriter->endElement();
    }

    /**
     * Write Defined Name for autoFilter
     *
     * @param \PhpOffice\PhpSpreadsheet\Shared\XMLWriter    $phpExcelSharedXMLWriter         XML Writer
     * @param     int                            $pSheetId
     * @throws     \PhpOffice\PhpSpreadsheet\Writer\Exception
     */
    private function writeDefinedNameForAutofilter(\PhpOffice\PhpSpreadsheet\Shared\XMLWriter $phpExcelSharedXMLWriter = \null, \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet $phpExcelWorksheet = \null, $pSheetId = 0)
    {
        // definedName for autoFilter
        $autoFilterRange = $phpExcelWorksheet->getAutoFilter()->getRange();
        if (!empty($autoFilterRange)) {
            $phpExcelSharedXMLWriter->startElement('definedName');
            $phpExcelSharedXMLWriter->writeAttribute('name', '_xlnm._FilterDatabase');
            $phpExcelSharedXMLWriter->writeAttribute('localSheetId', $pSheetId);
            $phpExcelSharedXMLWriter->writeAttribute('hidden', '1');

            // Create absolute coordinate and write as raw text
            $range = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::splitRange($autoFilterRange);
            $range = $range[0];
            //    Strip any worksheet ref so we can make the cell ref absolute
            if (\strpos($range[0], '!') !== \false) {
                list($ws, $range[0]) = \explode('!', $range[0]);
            }

            $range[0] = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::absoluteCoordinate($range[0]);
            $range[1] = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::absoluteCoordinate($range[1]);
            $range = \implode(':', $range);

            $phpExcelSharedXMLWriter->writeRawData('\'' . \str_replace("'", "''", $phpExcelWorksheet->getTitle()) . '\'!' . $range);

            $phpExcelSharedXMLWriter->endElement();
        }
    }

    /**
     * Write Defined Name for PrintTitles
     *
     * @param \PhpOffice\PhpSpreadsheet\Shared\XMLWriter    $phpExcelSharedXMLWriter         XML Writer
     * @param     int                            $pSheetId
     * @throws     \PhpOffice\PhpSpreadsheet\Writer\Exception
     */
    private function writeDefinedNameForPrintTitles(\PhpOffice\PhpSpreadsheet\Shared\XMLWriter $phpExcelSharedXMLWriter = \null, \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet $phpExcelWorksheet = \null, $pSheetId = 0)
    {
        // definedName for PrintTitles
        if ($phpExcelWorksheet->getPageSetup()->isColumnsToRepeatAtLeftSet() || $phpExcelWorksheet->getPageSetup()->isRowsToRepeatAtTopSet()) {
            $phpExcelSharedXMLWriter->startElement('definedName');
            $phpExcelSharedXMLWriter->writeAttribute('name', '_xlnm.Print_Titles');
            $phpExcelSharedXMLWriter->writeAttribute('localSheetId', $pSheetId);

            // Setting string
            $settingString = '';

            // Columns to repeat
            if ($phpExcelWorksheet->getPageSetup()->isColumnsToRepeatAtLeftSet()) {
                $repeat = $phpExcelWorksheet->getPageSetup()->getColumnsToRepeatAtLeft();

                $settingString .= '\'' . \str_replace("'", "''", $phpExcelWorksheet->getTitle()) . '\'!$' . $repeat[0] . ':$' . $repeat[1];
            }

            // Rows to repeat
            if ($phpExcelWorksheet->getPageSetup()->isRowsToRepeatAtTopSet()) {
                if ($phpExcelWorksheet->getPageSetup()->isColumnsToRepeatAtLeftSet()) {
                    $settingString .= ',';
                }

                $repeat = $phpExcelWorksheet->getPageSetup()->getRowsToRepeatAtTop();

                $settingString .= '\'' . \str_replace("'", "''", $phpExcelWorksheet->getTitle()) . '\'!$' . $repeat[0] . ':$' . $repeat[1];
            }

            $phpExcelSharedXMLWriter->writeRawData($settingString);

            $phpExcelSharedXMLWriter->endElement();
        }
    }

    /**
     * Write Defined Name for PrintTitles
     *
     * @param \PhpOffice\PhpSpreadsheet\Shared\XMLWriter    $phpExcelSharedXMLWriter         XML Writer
     * @param     int                            $pSheetId
     * @throws     \PhpOffice\PhpSpreadsheet\Writer\Exception
     */
    private function writeDefinedNameForPrintArea(\PhpOffice\PhpSpreadsheet\Shared\XMLWriter $phpExcelSharedXMLWriter = \null, \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet $phpExcelWorksheet = \null, $pSheetId = 0)
    {
        // definedName for PrintArea
        if ($phpExcelWorksheet->getPageSetup()->isPrintAreaSet()) {
            $phpExcelSharedXMLWriter->startElement('definedName');
            $phpExcelSharedXMLWriter->writeAttribute('name', '_xlnm.Print_Area');
            $phpExcelSharedXMLWriter->writeAttribute('localSheetId', $pSheetId);

            // Setting string
            $settingString = '';

            // Print area
            $printArea = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::splitRange($phpExcelWorksheet->getPageSetup()->getPrintArea());

            $chunks = array();
            foreach ($printArea as $singlePrintArea) {
                $singlePrintArea[0] = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::absoluteReference($singlePrintArea[0]);
                $singlePrintArea[1] = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::absoluteReference($singlePrintArea[1]);
                $chunks[] = '\'' . \str_replace("'", "''", $phpExcelWorksheet->getTitle()) . '\'!' . \implode(':', $singlePrintArea);
            }

            $phpExcelSharedXMLWriter->writeRawData(\implode(',', $chunks));

            $phpExcelSharedXMLWriter->endElement();
        }
    }
}

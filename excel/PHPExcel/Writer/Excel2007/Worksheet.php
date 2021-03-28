<?php
namespace PhpOffice\PhpSpreadsheet\Writer\Xlsx;

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
 * @package    PHPExcel_Writer_Excel2007
 * @copyright  Copyright (c) 2006 - 2015 PHPExcel (http://www.codeplex.com/PHPExcel)
 * @license    http://www.gnu.org/licenses/old-licenses/lgpl-2.1.txt    LGPL
 * @version    ##VERSION##, ##DATE##
 */
/**
 * PHPExcel_Writer_Excel2007_Worksheet
 *
 * @category   PHPExcel
 * @package    PHPExcel_Writer_Excel2007
 * @copyright  Copyright (c) 2006 - 2015 PHPExcel (http://www.codeplex.com/PHPExcel)
 */
class Worksheet extends \PhpOffice\PhpSpreadsheet\Writer\Xlsx\WriterPart
{
    /**
     * Write worksheet to XML format
     *
     * @param    \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet        $pSheet
     * @param    string[]                $pStringTable
     * @param    boolean                    $includeCharts    Flag indicating if we should write charts
     * @return    string                    XML Output
     * @throws    \PhpOffice\PhpSpreadsheet\Writer\Exception
     */
    public function writeWorksheet($pSheet = \null, $pStringTable = \null, $includeCharts = \false)
    {
        if (!\is_null($pSheet)) {
            // Create XML writer
            $objWriter = \null;
            if ($this->getParentWriter()->getUseDiskCaching()) {
                $objWriter = new \PhpOffice\PhpSpreadsheet\Shared\XMLWriter(\PhpOffice\PhpSpreadsheet\Shared\XMLWriter::STORAGE_DISK, $this->getParentWriter()->getDiskCachingDirectory());
            } else {
                $objWriter = new \PhpOffice\PhpSpreadsheet\Shared\XMLWriter(\PhpOffice\PhpSpreadsheet\Shared\XMLWriter::STORAGE_MEMORY);
            }

            // XML header
            $objWriter->startDocument('1.0', 'UTF-8', 'yes');

            // Worksheet
            $objWriter->startElement('worksheet');
            $objWriter->writeAttribute('xml:space', 'preserve');
            $objWriter->writeAttribute('xmlns', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main');
            $objWriter->writeAttribute('xmlns:r', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships');

                // sheetPr
                $this->writeSheetPr($objWriter, $pSheet);

                // Dimension
                $this->writeDimension($objWriter, $pSheet);

                // sheetViews
                $this->writeSheetViews($objWriter, $pSheet);

                // sheetFormatPr
                $this->writeSheetFormatPr($objWriter, $pSheet);

                // cols
                $this->writeCols($objWriter, $pSheet);

                // sheetData
                $this->writeSheetData($objWriter, $pSheet, $pStringTable);

                // sheetProtection
                $this->writeSheetProtection($objWriter, $pSheet);

                // protectedRanges
                $this->writeProtectedRanges($objWriter, $pSheet);

                // autoFilter
                $this->writeAutoFilter($objWriter, $pSheet);

                // mergeCells
                $this->writeMergeCells($objWriter, $pSheet);

                // conditionalFormatting
                $this->writeConditionalFormatting($objWriter, $pSheet);

                // dataValidations
                $this->writeDataValidations($objWriter, $pSheet);

                // hyperlinks
                $this->writeHyperlinks($objWriter, $pSheet);

                // Print options
                $this->writePrintOptions($objWriter, $pSheet);

                // Page margins
                $this->writePageMargins($objWriter, $pSheet);

                // Page setup
                $this->writePageSetup($objWriter, $pSheet);

                // Header / footer
                $this->writeHeaderFooter($objWriter, $pSheet);

                // Breaks
                $this->writeBreaks($objWriter, $pSheet);

                // Drawings and/or Charts
                $this->writeDrawings($objWriter, $pSheet, $includeCharts);

                // LegacyDrawing
                $this->writeLegacyDrawing($objWriter, $pSheet);

                // LegacyDrawingHF
                $this->writeLegacyDrawingHF($objWriter, $pSheet);

            $objWriter->endElement();

            // Return
            return $objWriter->getData();
        } else {
            throw new \PhpOffice\PhpSpreadsheet\Writer\Exception("Invalid PHPExcel_Worksheet object passed.");
        }
    }

    /**
     * Write SheetPr
     *
     * @param \PhpOffice\PhpSpreadsheet\Shared\XMLWriter        $phpExcelSharedXMLWriter        XML Writer
     * @param \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet                $phpExcelWorksheet            Worksheet
     * @throws    \PhpOffice\PhpSpreadsheet\Writer\Exception
     */
    private function writeSheetPr(\PhpOffice\PhpSpreadsheet\Shared\XMLWriter $phpExcelSharedXMLWriter = \null, \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet $phpExcelWorksheet = \null)
    {
        // sheetPr
        $phpExcelSharedXMLWriter->startElement('sheetPr');
        //$objWriter->writeAttribute('codeName',        $pSheet->getTitle());
        if ($phpExcelWorksheet->getParent()->hasMacros()) {//if the workbook have macros, we need to have codeName for the sheet
            if ($phpExcelWorksheet->hasCodeName()==\false) {
                $phpExcelWorksheet->setCodeName($phpExcelWorksheet->getTitle());
            }
            $phpExcelSharedXMLWriter->writeAttribute('codeName', $phpExcelWorksheet->getCodeName());
        }
        $autoFilterRange = $phpExcelWorksheet->getAutoFilter()->getRange();
        if (!empty($autoFilterRange)) {
            $phpExcelSharedXMLWriter->writeAttribute('filterMode', 1);
            $phpExcelWorksheet->getAutoFilter()->showHideRows();
        }

        // tabColor
        if ($phpExcelWorksheet->isTabColorSet()) {
            $phpExcelSharedXMLWriter->startElement('tabColor');
            $phpExcelSharedXMLWriter->writeAttribute('rgb', $phpExcelWorksheet->getTabColor()->getARGB());
            $phpExcelSharedXMLWriter->endElement();
        }

        // outlinePr
        $phpExcelSharedXMLWriter->startElement('outlinePr');
        $phpExcelSharedXMLWriter->writeAttribute('summaryBelow', ($phpExcelWorksheet->getShowSummaryBelow() ? '1' : '0'));
        $phpExcelSharedXMLWriter->writeAttribute('summaryRight', ($phpExcelWorksheet->getShowSummaryRight() ? '1' : '0'));
        $phpExcelSharedXMLWriter->endElement();

        // pageSetUpPr
        if ($phpExcelWorksheet->getPageSetup()->getFitToPage()) {
            $phpExcelSharedXMLWriter->startElement('pageSetUpPr');
            $phpExcelSharedXMLWriter->writeAttribute('fitToPage', '1');
            $phpExcelSharedXMLWriter->endElement();
        }

        $phpExcelSharedXMLWriter->endElement();
    }

    /**
     * Write Dimension
     *
     * @param \PhpOffice\PhpSpreadsheet\Shared\XMLWriter    $phpExcelSharedXMLWriter        XML Writer
     * @param \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet            $phpExcelWorksheet            Worksheet
     * @throws    \PhpOffice\PhpSpreadsheet\Writer\Exception
     */
    private function writeDimension(\PhpOffice\PhpSpreadsheet\Shared\XMLWriter $phpExcelSharedXMLWriter = \null, \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet $phpExcelWorksheet = \null)
    {
        // dimension
        $phpExcelSharedXMLWriter->startElement('dimension');
        $phpExcelSharedXMLWriter->writeAttribute('ref', $phpExcelWorksheet->calculateWorksheetDimension());
        $phpExcelSharedXMLWriter->endElement();
    }

    /**
     * Write SheetViews
     *
     * @param \PhpOffice\PhpSpreadsheet\Shared\XMLWriter            $phpExcelSharedXMLWriter        XML Writer
     * @param \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet                    $phpExcelWorksheet            Worksheet
     * @throws    \PhpOffice\PhpSpreadsheet\Writer\Exception
     */
    private function writeSheetViews(\PhpOffice\PhpSpreadsheet\Shared\XMLWriter $phpExcelSharedXMLWriter = \null, \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet $phpExcelWorksheet = \null)
    {
        // sheetViews
        $phpExcelSharedXMLWriter->startElement('sheetViews');

        // Sheet selected?
        $sheetSelected = \false;
        if ($this->getParentWriter()->getPHPExcel()->getIndex($phpExcelWorksheet) == $this->getParentWriter()->getPHPExcel()->getActiveSheetIndex()) {
            $sheetSelected = \true;
        }

        // sheetView
        $phpExcelSharedXMLWriter->startElement('sheetView');
        $phpExcelSharedXMLWriter->writeAttribute('tabSelected', $sheetSelected ? '1' : '0');
        $phpExcelSharedXMLWriter->writeAttribute('workbookViewId', '0');

        // Zoom scales
        if ($phpExcelWorksheet->getSheetView()->getZoomScale() != 100) {
            $phpExcelSharedXMLWriter->writeAttribute('zoomScale', $phpExcelWorksheet->getSheetView()->getZoomScale());
        }
        if ($phpExcelWorksheet->getSheetView()->getZoomScaleNormal() != 100) {
            $phpExcelSharedXMLWriter->writeAttribute('zoomScaleNormal', $phpExcelWorksheet->getSheetView()->getZoomScaleNormal());
        }

        // View Layout Type
        if ($phpExcelWorksheet->getSheetView()->getView() !== \PhpOffice\PhpSpreadsheet\Worksheet\SheetView::SHEETVIEW_NORMAL) {
            $phpExcelSharedXMLWriter->writeAttribute('view', $phpExcelWorksheet->getSheetView()->getView());
        }

        // Gridlines
        if ($phpExcelWorksheet->getShowGridlines()) {
            $phpExcelSharedXMLWriter->writeAttribute('showGridLines', 'true');
        } else {
            $phpExcelSharedXMLWriter->writeAttribute('showGridLines', 'false');
        }

        // Row and column headers
        if ($phpExcelWorksheet->getShowRowColHeaders()) {
            $phpExcelSharedXMLWriter->writeAttribute('showRowColHeaders', '1');
        } else {
            $phpExcelSharedXMLWriter->writeAttribute('showRowColHeaders', '0');
        }

        // Right-to-left
        if ($phpExcelWorksheet->getRightToLeft()) {
            $phpExcelSharedXMLWriter->writeAttribute('rightToLeft', 'true');
        }

        $activeCell = $phpExcelWorksheet->getActiveCell();

        // Pane
        $pane = '';
        $topLeftCell = $phpExcelWorksheet->getFreezePane();
        if (($topLeftCell != '') && ($topLeftCell != 'A1')) {
            $activeCell = empty($activeCell) ? $topLeftCell : $activeCell;
            // Calculate freeze coordinates
            $xSplit = $ySplit = 0;

            list($xSplit, $ySplit) = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::coordinateFromString($topLeftCell);
            $xSplit = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::columnIndexFromString($xSplit);

            // pane
            $pane = 'topRight';
            $phpExcelSharedXMLWriter->startElement('pane');
            if ($xSplit > 1) {
                $phpExcelSharedXMLWriter->writeAttribute('xSplit', $xSplit - 1);
            }
            if ($ySplit > 1) {
                $phpExcelSharedXMLWriter->writeAttribute('ySplit', $ySplit - 1);
                $pane = ($xSplit > 1) ? 'bottomRight' : 'bottomLeft';
            }
            $phpExcelSharedXMLWriter->writeAttribute('topLeftCell', $topLeftCell);
            $phpExcelSharedXMLWriter->writeAttribute('activePane', $pane);
            $phpExcelSharedXMLWriter->writeAttribute('state', 'frozen');
            $phpExcelSharedXMLWriter->endElement();

            if (($xSplit > 1) && ($ySplit > 1)) {
                //    Write additional selections if more than two panes (ie both an X and a Y split)
                $phpExcelSharedXMLWriter->startElement('selection');
                $phpExcelSharedXMLWriter->writeAttribute('pane', 'topRight');
                $phpExcelSharedXMLWriter->endElement();
                $phpExcelSharedXMLWriter->startElement('selection');
                $phpExcelSharedXMLWriter->writeAttribute('pane', 'bottomLeft');
                $phpExcelSharedXMLWriter->endElement();
            }
        }

        // Selection
//      if ($pane != '') {
        // Only need to write selection element if we have a split pane
        // We cheat a little by over-riding the active cell selection, setting it to the split cell
        $phpExcelSharedXMLWriter->startElement('selection');
        if ($pane != '') {
            $phpExcelSharedXMLWriter->writeAttribute('pane', $pane);
        }
        $phpExcelSharedXMLWriter->writeAttribute('activeCell', $activeCell);
        $phpExcelSharedXMLWriter->writeAttribute('sqref', $activeCell);
        $phpExcelSharedXMLWriter->endElement();
//      }

        $phpExcelSharedXMLWriter->endElement();

        $phpExcelSharedXMLWriter->endElement();
    }

    /**
     * Write SheetFormatPr
     *
     * @param \PhpOffice\PhpSpreadsheet\Shared\XMLWriter $phpExcelSharedXMLWriter        XML Writer
     * @param \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet          $phpExcelWorksheet            Worksheet
     * @throws    \PhpOffice\PhpSpreadsheet\Writer\Exception
     */
    private function writeSheetFormatPr(\PhpOffice\PhpSpreadsheet\Shared\XMLWriter $phpExcelSharedXMLWriter = \null, \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet $phpExcelWorksheet = \null)
    {
        // sheetFormatPr
        $phpExcelSharedXMLWriter->startElement('sheetFormatPr');

        // Default row height
        if ($phpExcelWorksheet->getDefaultRowDimension()->getRowHeight() >= 0) {
            $phpExcelSharedXMLWriter->writeAttribute('customHeight', 'true');
            $phpExcelSharedXMLWriter->writeAttribute('defaultRowHeight', \PhpOffice\PhpSpreadsheet\Shared\StringHelper::FormatNumber($phpExcelWorksheet->getDefaultRowDimension()->getRowHeight()));
        } else {
            $phpExcelSharedXMLWriter->writeAttribute('defaultRowHeight', '14.4');
        }

        // Set Zero Height row
        if ((string)$phpExcelWorksheet->getDefaultRowDimension()->getZeroHeight()  == '1' ||
            \strtolower((string)$phpExcelWorksheet->getDefaultRowDimension()->getZeroHeight()) == 'true') {
            $phpExcelSharedXMLWriter->writeAttribute('zeroHeight', '1');
        }

        // Default column width
        if ($phpExcelWorksheet->getDefaultColumnDimension()->getWidth() >= 0) {
            $phpExcelSharedXMLWriter->writeAttribute('defaultColWidth', \PhpOffice\PhpSpreadsheet\Shared\StringHelper::FormatNumber($phpExcelWorksheet->getDefaultColumnDimension()->getWidth()));
        }

        // Outline level - row
        $outlineLevelRow = 0;
        foreach ($phpExcelWorksheet->getRowDimensions() as $columnDimension) {
            if ($columnDimension->getOutlineLevel() > $outlineLevelRow) {
                $outlineLevelRow = $columnDimension->getOutlineLevel();
            }
        }
        $phpExcelSharedXMLWriter->writeAttribute('outlineLevelRow', $outlineLevelRow);

        // Outline level - column
        $outlineLevelCol = 0;
        foreach ($phpExcelWorksheet->getColumnDimensions() as $columnDimension) {
            if ($columnDimension->getOutlineLevel() > $outlineLevelCol) {
                $outlineLevelCol = $columnDimension->getOutlineLevel();
            }
        }
        $phpExcelSharedXMLWriter->writeAttribute('outlineLevelCol', $outlineLevelCol);

        $phpExcelSharedXMLWriter->endElement();
    }

    /**
     * Write Cols
     *
     * @param \PhpOffice\PhpSpreadsheet\Shared\XMLWriter $phpExcelSharedXMLWriter XML Writer
     * @param \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet $phpExcelWorksheet Worksheet
     * @throws    \PhpOffice\PhpSpreadsheet\Writer\Exception
     */
    private function writeCols(\PhpOffice\PhpSpreadsheet\Shared\XMLWriter $phpExcelSharedXMLWriter = \null, \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet $phpExcelWorksheet = \null)
    {
        // cols
        if (\count($phpExcelWorksheet->getColumnDimensions()) > 0) {
            $phpExcelSharedXMLWriter->startElement('cols');

            $phpExcelWorksheet->calculateColumnWidths();

            // Loop through column dimensions
            foreach ($phpExcelWorksheet->getColumnDimensions() as $phpExcelWorksheetColumnDimension) {
                // col
                $phpExcelSharedXMLWriter->startElement('col');
                $phpExcelSharedXMLWriter->writeAttribute('min', \PhpOffice\PhpSpreadsheet\Cell\Coordinate::columnIndexFromString($phpExcelWorksheetColumnDimension->getColumnIndex()));
                $phpExcelSharedXMLWriter->writeAttribute('max', \PhpOffice\PhpSpreadsheet\Cell\Coordinate::columnIndexFromString($phpExcelWorksheetColumnDimension->getColumnIndex()));

                if ($phpExcelWorksheetColumnDimension->getWidth() < 0) {
                    // No width set, apply default of 10
                    $phpExcelSharedXMLWriter->writeAttribute('width', '9.10');
                } else {
                    // Width set
                    $phpExcelSharedXMLWriter->writeAttribute('width', \PhpOffice\PhpSpreadsheet\Shared\StringHelper::FormatNumber($phpExcelWorksheetColumnDimension->getWidth()));
                }

                // Column visibility
                if ($phpExcelWorksheetColumnDimension->isVisible() == \false) {
                    $phpExcelSharedXMLWriter->writeAttribute('hidden', 'true');
                }

                // Auto size?
                if ($phpExcelWorksheetColumnDimension->isAutoSize()) {
                    $phpExcelSharedXMLWriter->writeAttribute('bestFit', 'true');
                }

                // Custom width?
                if ($phpExcelWorksheetColumnDimension->getWidth() != $phpExcelWorksheet->getDefaultColumnDimension()->getWidth()) {
                    $phpExcelSharedXMLWriter->writeAttribute('customWidth', 'true');
                }

                // Collapsed
                if ($phpExcelWorksheetColumnDimension->isCollapsed() == \true) {
                    $phpExcelSharedXMLWriter->writeAttribute('collapsed', 'true');
                }

                // Outline level
                if ($phpExcelWorksheetColumnDimension->getOutlineLevel() > 0) {
                    $phpExcelSharedXMLWriter->writeAttribute('outlineLevel', $phpExcelWorksheetColumnDimension->getOutlineLevel());
                }

                // Style
                $phpExcelSharedXMLWriter->writeAttribute('style', $phpExcelWorksheetColumnDimension->getXfIndex());

                $phpExcelSharedXMLWriter->endElement();
            }

            $phpExcelSharedXMLWriter->endElement();
        }
    }

    /**
     * Write SheetProtection
     *
     * @param \PhpOffice\PhpSpreadsheet\Shared\XMLWriter            $phpExcelSharedXMLWriter        XML Writer
     * @param \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet                    $phpExcelWorksheet            Worksheet
     * @throws    \PhpOffice\PhpSpreadsheet\Writer\Exception
     */
    private function writeSheetProtection(\PhpOffice\PhpSpreadsheet\Shared\XMLWriter $phpExcelSharedXMLWriter = \null, \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet $phpExcelWorksheet = \null)
    {
        // sheetProtection
        $phpExcelSharedXMLWriter->startElement('sheetProtection');

        if ($phpExcelWorksheet->getProtection()->getPassword() != '') {
            $phpExcelSharedXMLWriter->writeAttribute('password', $phpExcelWorksheet->getProtection()->getPassword());
        }

        $phpExcelSharedXMLWriter->writeAttribute('sheet', ($phpExcelWorksheet->getProtection()->getSheet() ? 'true' : 'false'));
        $phpExcelSharedXMLWriter->writeAttribute('objects', ($phpExcelWorksheet->getProtection()->getObjects() ? 'true' : 'false'));
        $phpExcelSharedXMLWriter->writeAttribute('scenarios', ($phpExcelWorksheet->getProtection()->getScenarios() ? 'true' : 'false'));
        $phpExcelSharedXMLWriter->writeAttribute('formatCells', ($phpExcelWorksheet->getProtection()->getFormatCells() ? 'true' : 'false'));
        $phpExcelSharedXMLWriter->writeAttribute('formatColumns', ($phpExcelWorksheet->getProtection()->getFormatColumns() ? 'true' : 'false'));
        $phpExcelSharedXMLWriter->writeAttribute('formatRows', ($phpExcelWorksheet->getProtection()->getFormatRows() ? 'true' : 'false'));
        $phpExcelSharedXMLWriter->writeAttribute('insertColumns', ($phpExcelWorksheet->getProtection()->getInsertColumns() ? 'true' : 'false'));
        $phpExcelSharedXMLWriter->writeAttribute('insertRows', ($phpExcelWorksheet->getProtection()->getInsertRows() ? 'true' : 'false'));
        $phpExcelSharedXMLWriter->writeAttribute('insertHyperlinks', ($phpExcelWorksheet->getProtection()->getInsertHyperlinks() ? 'true' : 'false'));
        $phpExcelSharedXMLWriter->writeAttribute('deleteColumns', ($phpExcelWorksheet->getProtection()->getDeleteColumns() ? 'true' : 'false'));
        $phpExcelSharedXMLWriter->writeAttribute('deleteRows', ($phpExcelWorksheet->getProtection()->getDeleteRows() ? 'true' : 'false'));
        $phpExcelSharedXMLWriter->writeAttribute('selectLockedCells', ($phpExcelWorksheet->getProtection()->getSelectLockedCells() ? 'true' : 'false'));
        $phpExcelSharedXMLWriter->writeAttribute('sort', ($phpExcelWorksheet->getProtection()->getSort() ? 'true' : 'false'));
        $phpExcelSharedXMLWriter->writeAttribute('autoFilter', ($phpExcelWorksheet->getProtection()->getAutoFilter() ? 'true' : 'false'));
        $phpExcelSharedXMLWriter->writeAttribute('pivotTables', ($phpExcelWorksheet->getProtection()->getPivotTables() ? 'true' : 'false'));
        $phpExcelSharedXMLWriter->writeAttribute('selectUnlockedCells', ($phpExcelWorksheet->getProtection()->getSelectUnlockedCells() ? 'true' : 'false'));
        $phpExcelSharedXMLWriter->endElement();
    }

    /**
     * Write ConditionalFormatting
     *
     * @param \PhpOffice\PhpSpreadsheet\Shared\XMLWriter            $phpExcelSharedXMLWriter        XML Writer
     * @param \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet                    $phpExcelWorksheet            Worksheet
     * @throws    \PhpOffice\PhpSpreadsheet\Writer\Exception
     */
    private function writeConditionalFormatting(\PhpOffice\PhpSpreadsheet\Shared\XMLWriter $phpExcelSharedXMLWriter = \null, \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet $phpExcelWorksheet = \null)
    {
        // Conditional id
        $id = 1;

        // Loop through styles in the current worksheet
        foreach ($phpExcelWorksheet->getConditionalStylesCollection() as $cellCoordinate => $conditionalStyles) {
            foreach ($conditionalStyles as $conditionalStyle) {
                // WHY was this again?
                // if ($this->getParentWriter()->getStylesConditionalHashTable()->getIndexForHashCode($conditional->getHashCode()) == '') {
                //    continue;
                // }
                if ($conditionalStyle->getConditionType() != \PhpOffice\PhpSpreadsheet\Style\Conditional::CONDITION_NONE) {
                    // conditionalFormatting
                    $phpExcelSharedXMLWriter->startElement('conditionalFormatting');
                    $phpExcelSharedXMLWriter->writeAttribute('sqref', $cellCoordinate);

                    // cfRule
                    $phpExcelSharedXMLWriter->startElement('cfRule');
                    $phpExcelSharedXMLWriter->writeAttribute('type', $conditionalStyle->getConditionType());
                    $phpExcelSharedXMLWriter->writeAttribute('dxfId', $this->getParentWriter()->getStylesConditionalHashTable()->getIndexForHashCode($conditionalStyle->getHashCode()));
                    $phpExcelSharedXMLWriter->writeAttribute('priority', $id++);

                    if (($conditionalStyle->getConditionType() == \PhpOffice\PhpSpreadsheet\Style\Conditional::CONDITION_CELLIS || $conditionalStyle->getConditionType() == \PhpOffice\PhpSpreadsheet\Style\Conditional::CONDITION_CONTAINSTEXT)
                        && $conditionalStyle->getOperatorType() != \PhpOffice\PhpSpreadsheet\Style\Conditional::OPERATOR_NONE) {
                        $phpExcelSharedXMLWriter->writeAttribute('operator', $conditionalStyle->getOperatorType());
                    }

                    if ($conditionalStyle->getConditionType() == \PhpOffice\PhpSpreadsheet\Style\Conditional::CONDITION_CONTAINSTEXT
                        && !\is_null($conditionalStyle->getText())) {
                        $phpExcelSharedXMLWriter->writeAttribute('text', $conditionalStyle->getText());
                    }

                    if ($conditionalStyle->getConditionType() == \PhpOffice\PhpSpreadsheet\Style\Conditional::CONDITION_CONTAINSTEXT
                        && $conditionalStyle->getOperatorType() == \PhpOffice\PhpSpreadsheet\Style\Conditional::OPERATOR_CONTAINSTEXT
                        && !\is_null($conditionalStyle->getText())) {
                        $phpExcelSharedXMLWriter->writeElement('formula', 'NOT(ISERROR(SEARCH("' . $conditionalStyle->getText() . '",' . $cellCoordinate . ')))');
                    } elseif ($conditionalStyle->getConditionType() == \PhpOffice\PhpSpreadsheet\Style\Conditional::CONDITION_CONTAINSTEXT
                        && $conditionalStyle->getOperatorType() == \PhpOffice\PhpSpreadsheet\Style\Conditional::OPERATOR_BEGINSWITH
                        && !\is_null($conditionalStyle->getText())) {
                        $phpExcelSharedXMLWriter->writeElement('formula', 'LEFT(' . $cellCoordinate . ',' . \strlen($conditionalStyle->getText()) . ')="' . $conditionalStyle->getText() . '"');
                    } elseif ($conditionalStyle->getConditionType() == \PhpOffice\PhpSpreadsheet\Style\Conditional::CONDITION_CONTAINSTEXT
                        && $conditionalStyle->getOperatorType() == \PhpOffice\PhpSpreadsheet\Style\Conditional::OPERATOR_ENDSWITH
                        && !\is_null($conditionalStyle->getText())) {
                        $phpExcelSharedXMLWriter->writeElement('formula', 'RIGHT(' . $cellCoordinate . ',' . \strlen($conditionalStyle->getText()) . ')="' . $conditionalStyle->getText() . '"');
                    } elseif ($conditionalStyle->getConditionType() == \PhpOffice\PhpSpreadsheet\Style\Conditional::CONDITION_CONTAINSTEXT
                        && $conditionalStyle->getOperatorType() == \PhpOffice\PhpSpreadsheet\Style\Conditional::OPERATOR_NOTCONTAINS
                        && !\is_null($conditionalStyle->getText())) {
                        $phpExcelSharedXMLWriter->writeElement('formula', 'ISERROR(SEARCH("' . $conditionalStyle->getText() . '",' . $cellCoordinate . '))');
                    } elseif ($conditionalStyle->getConditionType() == \PhpOffice\PhpSpreadsheet\Style\Conditional::CONDITION_CELLIS
                        || $conditionalStyle->getConditionType() == \PhpOffice\PhpSpreadsheet\Style\Conditional::CONDITION_CONTAINSTEXT
                        || $conditionalStyle->getConditionType() == \PhpOffice\PhpSpreadsheet\Style\Conditional::CONDITION_EXPRESSION) {
                        foreach ($conditionalStyle->getConditions() as $condition) {
                            // Formula
                            $phpExcelSharedXMLWriter->writeElement('formula', $condition);
                        }
                    }

                    $phpExcelSharedXMLWriter->endElement();

                    $phpExcelSharedXMLWriter->endElement();
                }
            }
        }
    }

    /**
     * Write DataValidations
     *
     * @param \PhpOffice\PhpSpreadsheet\Shared\XMLWriter            $phpExcelSharedXMLWriter        XML Writer
     * @param \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet                    $phpExcelWorksheet            Worksheet
     * @throws    \PhpOffice\PhpSpreadsheet\Writer\Exception
     */
    private function writeDataValidations(\PhpOffice\PhpSpreadsheet\Shared\XMLWriter $phpExcelSharedXMLWriter = \null, \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet $phpExcelWorksheet = \null)
    {
        // Datavalidation collection
        $dataValidationCollection = $phpExcelWorksheet->getDataValidationCollection();

        // Write data validations?
        if (!empty($dataValidationCollection)) {
            $phpExcelSharedXMLWriter->startElement('dataValidations');
            $phpExcelSharedXMLWriter->writeAttribute('count', \count($dataValidationCollection));

            foreach ($dataValidationCollection as $coordinate => $dv) {
                $phpExcelSharedXMLWriter->startElement('dataValidation');

                if ($dv->getType() != '') {
                    $phpExcelSharedXMLWriter->writeAttribute('type', $dv->getType());
                }

                if ($dv->getErrorStyle() != '') {
                    $phpExcelSharedXMLWriter->writeAttribute('errorStyle', $dv->getErrorStyle());
                }

                if ($dv->getOperator() != '') {
                    $phpExcelSharedXMLWriter->writeAttribute('operator', $dv->getOperator());
                }

                $phpExcelSharedXMLWriter->writeAttribute('allowBlank', ($dv->isAllowBlank() ? '1'  : '0'));
                $phpExcelSharedXMLWriter->writeAttribute('showDropDown', ($dv->isShowDropDown() ? '0'  : '1'));
                $phpExcelSharedXMLWriter->writeAttribute('showInputMessage', ($dv->isShowInputMessage() ? '1'  : '0'));
                $phpExcelSharedXMLWriter->writeAttribute('showErrorMessage', ($dv->isShowErrorMessage() ? '1'  : '0'));

                if ($dv->getErrorTitle() !== '') {
                    $phpExcelSharedXMLWriter->writeAttribute('errorTitle', $dv->getErrorTitle());
                }
                if ($dv->getError() !== '') {
                    $phpExcelSharedXMLWriter->writeAttribute('error', $dv->getError());
                }
                if ($dv->getPromptTitle() !== '') {
                    $phpExcelSharedXMLWriter->writeAttribute('promptTitle', $dv->getPromptTitle());
                }
                if ($dv->getPrompt() !== '') {
                    $phpExcelSharedXMLWriter->writeAttribute('prompt', $dv->getPrompt());
                }

                $phpExcelSharedXMLWriter->writeAttribute('sqref', $coordinate);

                if ($dv->getFormula1() !== '') {
                    $phpExcelSharedXMLWriter->writeElement('formula1', $dv->getFormula1());
                }
                if ($dv->getFormula2() !== '') {
                    $phpExcelSharedXMLWriter->writeElement('formula2', $dv->getFormula2());
                }

                $phpExcelSharedXMLWriter->endElement();
            }

            $phpExcelSharedXMLWriter->endElement();
        }
    }

    /**
     * Write Hyperlinks
     *
     * @param \PhpOffice\PhpSpreadsheet\Shared\XMLWriter            $phpExcelSharedXMLWriter        XML Writer
     * @param \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet                    $phpExcelWorksheet            Worksheet
     * @throws    \PhpOffice\PhpSpreadsheet\Writer\Exception
     */
    private function writeHyperlinks(\PhpOffice\PhpSpreadsheet\Shared\XMLWriter $phpExcelSharedXMLWriter = \null, \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet $phpExcelWorksheet = \null)
    {
        // Hyperlink collection
        $hyperlinkCollection = $phpExcelWorksheet->getHyperlinkCollection();

        // Relation ID
        $relationId = 1;

        // Write hyperlinks?
        if (!empty($hyperlinkCollection)) {
            $phpExcelSharedXMLWriter->startElement('hyperlinks');

            foreach ($hyperlinkCollection as $coordinate => $hyperlink) {
                $phpExcelSharedXMLWriter->startElement('hyperlink');

                $phpExcelSharedXMLWriter->writeAttribute('ref', $coordinate);
                if (!$hyperlink->isInternal()) {
                    $phpExcelSharedXMLWriter->writeAttribute('r:id', 'rId_hyperlink_' . $relationId);
                    ++$relationId;
                } else {
                    $phpExcelSharedXMLWriter->writeAttribute('location', \str_replace('sheet://', '', $hyperlink->getUrl()));
                }

                if ($hyperlink->getTooltip() != '') {
                    $phpExcelSharedXMLWriter->writeAttribute('tooltip', $hyperlink->getTooltip());
                }

                $phpExcelSharedXMLWriter->endElement();
            }

            $phpExcelSharedXMLWriter->endElement();
        }
    }

    /**
     * Write ProtectedRanges
     *
     * @param \PhpOffice\PhpSpreadsheet\Shared\XMLWriter            $phpExcelSharedXMLWriter        XML Writer
     * @param \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet                    $phpExcelWorksheet            Worksheet
     * @throws    \PhpOffice\PhpSpreadsheet\Writer\Exception
     */
    private function writeProtectedRanges(\PhpOffice\PhpSpreadsheet\Shared\XMLWriter $phpExcelSharedXMLWriter = \null, \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet $phpExcelWorksheet = \null)
    {
        if (\count($phpExcelWorksheet->getProtectedCells()) > 0) {
            // protectedRanges
            $phpExcelSharedXMLWriter->startElement('protectedRanges');

            // Loop protectedRanges
            foreach ($phpExcelWorksheet->getProtectedCells() as $protectedCell => $passwordHash) {
                // protectedRange
                $phpExcelSharedXMLWriter->startElement('protectedRange');
                $phpExcelSharedXMLWriter->writeAttribute('name', 'p' . \md5($protectedCell));
                $phpExcelSharedXMLWriter->writeAttribute('sqref', $protectedCell);
                if (!empty($passwordHash)) {
                    $phpExcelSharedXMLWriter->writeAttribute('password', $passwordHash);
                }
                $phpExcelSharedXMLWriter->endElement();
            }

            $phpExcelSharedXMLWriter->endElement();
        }
    }

    /**
     * Write MergeCells
     *
     * @param \PhpOffice\PhpSpreadsheet\Shared\XMLWriter            $phpExcelSharedXMLWriter        XML Writer
     * @param \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet                    $phpExcelWorksheet            Worksheet
     * @throws    \PhpOffice\PhpSpreadsheet\Writer\Exception
     */
    private function writeMergeCells(\PhpOffice\PhpSpreadsheet\Shared\XMLWriter $phpExcelSharedXMLWriter = \null, \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet $phpExcelWorksheet = \null)
    {
        if (\count($phpExcelWorksheet->getMergeCells()) > 0) {
            // mergeCells
            $phpExcelSharedXMLWriter->startElement('mergeCells');

            // Loop mergeCells
            foreach ($phpExcelWorksheet->getMergeCells() as $mergeCell) {
                // mergeCell
                $phpExcelSharedXMLWriter->startElement('mergeCell');
                $phpExcelSharedXMLWriter->writeAttribute('ref', $mergeCell);
                $phpExcelSharedXMLWriter->endElement();
            }

            $phpExcelSharedXMLWriter->endElement();
        }
    }

    /**
     * Write PrintOptions
     *
     * @param \PhpOffice\PhpSpreadsheet\Shared\XMLWriter            $phpExcelSharedXMLWriter        XML Writer
     * @param \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet                    $phpExcelWorksheet            Worksheet
     * @throws    \PhpOffice\PhpSpreadsheet\Writer\Exception
     */
    private function writePrintOptions(\PhpOffice\PhpSpreadsheet\Shared\XMLWriter $phpExcelSharedXMLWriter = \null, \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet $phpExcelWorksheet = \null)
    {
        // printOptions
        $phpExcelSharedXMLWriter->startElement('printOptions');

        $phpExcelSharedXMLWriter->writeAttribute('gridLines', ($phpExcelWorksheet->getPrintGridlines() ? 'true': 'false'));
        $phpExcelSharedXMLWriter->writeAttribute('gridLinesSet', 'true');

        if ($phpExcelWorksheet->getPageSetup()->getHorizontalCentered()) {
            $phpExcelSharedXMLWriter->writeAttribute('horizontalCentered', 'true');
        }

        if ($phpExcelWorksheet->getPageSetup()->getVerticalCentered()) {
            $phpExcelSharedXMLWriter->writeAttribute('verticalCentered', 'true');
        }

        $phpExcelSharedXMLWriter->endElement();
    }

    /**
     * Write PageMargins
     *
     * @param \PhpOffice\PhpSpreadsheet\Shared\XMLWriter                $phpExcelSharedXMLWriter        XML Writer
     * @param \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet                        $phpExcelWorksheet            Worksheet
     * @throws    \PhpOffice\PhpSpreadsheet\Writer\Exception
     */
    private function writePageMargins(\PhpOffice\PhpSpreadsheet\Shared\XMLWriter $phpExcelSharedXMLWriter = \null, \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet $phpExcelWorksheet = \null)
    {
        // pageMargins
        $phpExcelSharedXMLWriter->startElement('pageMargins');
        $phpExcelSharedXMLWriter->writeAttribute('left', \PhpOffice\PhpSpreadsheet\Shared\StringHelper::FormatNumber($phpExcelWorksheet->getPageMargins()->getLeft()));
        $phpExcelSharedXMLWriter->writeAttribute('right', \PhpOffice\PhpSpreadsheet\Shared\StringHelper::FormatNumber($phpExcelWorksheet->getPageMargins()->getRight()));
        $phpExcelSharedXMLWriter->writeAttribute('top', \PhpOffice\PhpSpreadsheet\Shared\StringHelper::FormatNumber($phpExcelWorksheet->getPageMargins()->getTop()));
        $phpExcelSharedXMLWriter->writeAttribute('bottom', \PhpOffice\PhpSpreadsheet\Shared\StringHelper::FormatNumber($phpExcelWorksheet->getPageMargins()->getBottom()));
        $phpExcelSharedXMLWriter->writeAttribute('header', \PhpOffice\PhpSpreadsheet\Shared\StringHelper::FormatNumber($phpExcelWorksheet->getPageMargins()->getHeader()));
        $phpExcelSharedXMLWriter->writeAttribute('footer', \PhpOffice\PhpSpreadsheet\Shared\StringHelper::FormatNumber($phpExcelWorksheet->getPageMargins()->getFooter()));
        $phpExcelSharedXMLWriter->endElement();
    }

    /**
     * Write AutoFilter
     *
     * @param \PhpOffice\PhpSpreadsheet\Shared\XMLWriter                $phpExcelSharedXMLWriter        XML Writer
     * @param \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet                        $phpExcelWorksheet            Worksheet
     * @throws    \PhpOffice\PhpSpreadsheet\Writer\Exception
     */
    private function writeAutoFilter(\PhpOffice\PhpSpreadsheet\Shared\XMLWriter $phpExcelSharedXMLWriter = \null, \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet $phpExcelWorksheet = \null)
    {
        $autoFilterRange = $phpExcelWorksheet->getAutoFilter()->getRange();
        if (!empty($autoFilterRange)) {
            // autoFilter
            $phpExcelSharedXMLWriter->startElement('autoFilter');

            // Strip any worksheet reference from the filter coordinates
            $range = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::splitRange($autoFilterRange);
            $range = $range[0];
            //    Strip any worksheet ref
            if (\strpos($range[0], '!') !== \false) {
                list($ws, $range[0]) = \explode('!', $range[0]);
            }
            $range = \implode(':', $range);

            $phpExcelSharedXMLWriter->writeAttribute('ref', \str_replace('$', '', $range));

            $columns = $phpExcelWorksheet->getAutoFilter()->getColumns();
            if (\count($columns > 0) > 0) {
                foreach ($columns as $columnID => $column) {
                    $rules = $column->getRules();
                    if (\count($rules) > 0) {
                        $phpExcelSharedXMLWriter->startElement('filterColumn');
                        $phpExcelSharedXMLWriter->writeAttribute('colId', $phpExcelWorksheet->getAutoFilter()->getColumnOffset($columnID));

                        $phpExcelSharedXMLWriter->startElement($column->getFilterType());
                        if ($column->getJoin() == \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column::AUTOFILTER_COLUMN_JOIN_AND) {
                            $phpExcelSharedXMLWriter->writeAttribute('and', 1);
                        }

                        foreach ($rules as $rule) {
                            if (($column->getFilterType() === \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column::AUTOFILTER_FILTERTYPE_FILTER) &&
                                ($rule->getOperator() === \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column\Rule::AUTOFILTER_COLUMN_RULE_EQUAL) &&
                                ($rule->getValue() === '')) {
                                //    Filter rule for Blanks
                                $phpExcelSharedXMLWriter->writeAttribute('blank', 1);
                            } elseif ($rule->getRuleType() === \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column\Rule::AUTOFILTER_RULETYPE_DYNAMICFILTER) {
                                //    Dynamic Filter Rule
                                $phpExcelSharedXMLWriter->writeAttribute('type', $rule->getGrouping());
                                $val = $column->getAttribute('val');
                                if ($val !== \null) {
                                    $phpExcelSharedXMLWriter->writeAttribute('val', $val);
                                }
                                $maxVal = $column->getAttribute('maxVal');
                                if ($maxVal !== \null) {
                                    $phpExcelSharedXMLWriter->writeAttribute('maxVal', $maxVal);
                                }
                            } elseif ($rule->getRuleType() === \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column\Rule::AUTOFILTER_RULETYPE_TOPTENFILTER) {
                                //    Top 10 Filter Rule
                                $phpExcelSharedXMLWriter->writeAttribute('val', $rule->getValue());
                                $phpExcelSharedXMLWriter->writeAttribute('percent', (($rule->getOperator() === \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column\Rule::AUTOFILTER_COLUMN_RULE_TOPTEN_PERCENT) ? '1' : '0'));
                                $phpExcelSharedXMLWriter->writeAttribute('top', (($rule->getGrouping() === \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column\Rule::AUTOFILTER_COLUMN_RULE_TOPTEN_TOP) ? '1': '0'));
                            } else {
                                //    Filter, DateGroupItem or CustomFilter
                                $phpExcelSharedXMLWriter->startElement($rule->getRuleType());

                                if ($rule->getOperator() !== \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column\Rule::AUTOFILTER_COLUMN_RULE_EQUAL) {
                                    $phpExcelSharedXMLWriter->writeAttribute('operator', $rule->getOperator());
                                }
                                if ($rule->getRuleType() === \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column\Rule::AUTOFILTER_RULETYPE_DATEGROUP) {
                                    // Date Group filters
                                    foreach ($rule->getValue() as $key => $value) {
                                        if ($value > '') {
                                            $phpExcelSharedXMLWriter->writeAttribute($key, $value);
                                        }
                                    }
                                    $phpExcelSharedXMLWriter->writeAttribute('dateTimeGrouping', $rule->getGrouping());
                                } else {
                                    $phpExcelSharedXMLWriter->writeAttribute('val', $rule->getValue());
                                }

                                $phpExcelSharedXMLWriter->endElement();
                            }
                        }

                        $phpExcelSharedXMLWriter->endElement();

                        $phpExcelSharedXMLWriter->endElement();
                    }
                }
            }
            $phpExcelSharedXMLWriter->endElement();
        }
    }

    /**
     * Write PageSetup
     *
     * @param \PhpOffice\PhpSpreadsheet\Shared\XMLWriter            $phpExcelSharedXMLWriter        XML Writer
     * @param \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet                    $phpExcelWorksheet            Worksheet
     * @throws    \PhpOffice\PhpSpreadsheet\Writer\Exception
     */
    private function writePageSetup(\PhpOffice\PhpSpreadsheet\Shared\XMLWriter $phpExcelSharedXMLWriter = \null, \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet $phpExcelWorksheet = \null)
    {
        // pageSetup
        $phpExcelSharedXMLWriter->startElement('pageSetup');
        $phpExcelSharedXMLWriter->writeAttribute('paperSize', $phpExcelWorksheet->getPageSetup()->getPaperSize());
        $phpExcelSharedXMLWriter->writeAttribute('orientation', $phpExcelWorksheet->getPageSetup()->getOrientation());

        if (!\is_null($phpExcelWorksheet->getPageSetup()->getScale())) {
            $phpExcelSharedXMLWriter->writeAttribute('scale', $phpExcelWorksheet->getPageSetup()->getScale());
        }
        if (!\is_null($phpExcelWorksheet->getPageSetup()->getFitToHeight())) {
            $phpExcelSharedXMLWriter->writeAttribute('fitToHeight', $phpExcelWorksheet->getPageSetup()->getFitToHeight());
        } else {
            $phpExcelSharedXMLWriter->writeAttribute('fitToHeight', '0');
        }
        if (!\is_null($phpExcelWorksheet->getPageSetup()->getFitToWidth())) {
            $phpExcelSharedXMLWriter->writeAttribute('fitToWidth', $phpExcelWorksheet->getPageSetup()->getFitToWidth());
        } else {
            $phpExcelSharedXMLWriter->writeAttribute('fitToWidth', '0');
        }
        if (!\is_null($phpExcelWorksheet->getPageSetup()->getFirstPageNumber())) {
            $phpExcelSharedXMLWriter->writeAttribute('firstPageNumber', $phpExcelWorksheet->getPageSetup()->getFirstPageNumber());
            $phpExcelSharedXMLWriter->writeAttribute('useFirstPageNumber', '1');
        }

        $phpExcelSharedXMLWriter->endElement();
    }

    /**
     * Write Header / Footer
     *
     * @param \PhpOffice\PhpSpreadsheet\Shared\XMLWriter        $phpExcelSharedXMLWriter        XML Writer
     * @param \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet                $phpExcelWorksheet            Worksheet
     * @throws    \PhpOffice\PhpSpreadsheet\Writer\Exception
     */
    private function writeHeaderFooter(\PhpOffice\PhpSpreadsheet\Shared\XMLWriter $phpExcelSharedXMLWriter = \null, \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet $phpExcelWorksheet = \null)
    {
        // headerFooter
        $phpExcelSharedXMLWriter->startElement('headerFooter');
        $phpExcelSharedXMLWriter->writeAttribute('differentOddEven', ($phpExcelWorksheet->getHeaderFooter()->getDifferentOddEven() ? 'true' : 'false'));
        $phpExcelSharedXMLWriter->writeAttribute('differentFirst', ($phpExcelWorksheet->getHeaderFooter()->getDifferentFirst() ? 'true' : 'false'));
        $phpExcelSharedXMLWriter->writeAttribute('scaleWithDoc', ($phpExcelWorksheet->getHeaderFooter()->getScaleWithDocument() ? 'true' : 'false'));
        $phpExcelSharedXMLWriter->writeAttribute('alignWithMargins', ($phpExcelWorksheet->getHeaderFooter()->getAlignWithMargins() ? 'true' : 'false'));

        $phpExcelSharedXMLWriter->writeElement('oddHeader', $phpExcelWorksheet->getHeaderFooter()->getOddHeader());
        $phpExcelSharedXMLWriter->writeElement('oddFooter', $phpExcelWorksheet->getHeaderFooter()->getOddFooter());
        $phpExcelSharedXMLWriter->writeElement('evenHeader', $phpExcelWorksheet->getHeaderFooter()->getEvenHeader());
        $phpExcelSharedXMLWriter->writeElement('evenFooter', $phpExcelWorksheet->getHeaderFooter()->getEvenFooter());
        $phpExcelSharedXMLWriter->writeElement('firstHeader', $phpExcelWorksheet->getHeaderFooter()->getFirstHeader());
        $phpExcelSharedXMLWriter->writeElement('firstFooter', $phpExcelWorksheet->getHeaderFooter()->getFirstFooter());
        $phpExcelSharedXMLWriter->endElement();
    }

    /**
     * Write Breaks
     *
     * @param \PhpOffice\PhpSpreadsheet\Shared\XMLWriter        $phpExcelSharedXMLWriter        XML Writer
     * @param \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet                $phpExcelWorksheet            Worksheet
     * @throws    \PhpOffice\PhpSpreadsheet\Writer\Exception
     */
    private function writeBreaks(\PhpOffice\PhpSpreadsheet\Shared\XMLWriter $phpExcelSharedXMLWriter = \null, \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet $phpExcelWorksheet = \null)
    {
        // Get row and column breaks
        $aRowBreaks = array();
        $aColumnBreaks = array();
        foreach ($phpExcelWorksheet->getBreaks() as $cell => $breakType) {
            if ($breakType == \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet::BREAK_ROW) {
                $aRowBreaks[] = $cell;
            } elseif ($breakType == \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet::BREAK_COLUMN) {
                $aColumnBreaks[] = $cell;
            }
        }

        // rowBreaks
        if (!empty($aRowBreaks)) {
            $phpExcelSharedXMLWriter->startElement('rowBreaks');
            $phpExcelSharedXMLWriter->writeAttribute('count', \count($aRowBreaks));
            $phpExcelSharedXMLWriter->writeAttribute('manualBreakCount', \count($aRowBreaks));

            foreach ($aRowBreaks as $aRowBreak) {
                $coords = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::coordinateFromString($aRowBreak);

                $phpExcelSharedXMLWriter->startElement('brk');
                $phpExcelSharedXMLWriter->writeAttribute('id', $coords[1]);
                $phpExcelSharedXMLWriter->writeAttribute('man', '1');
                $phpExcelSharedXMLWriter->endElement();
            }

            $phpExcelSharedXMLWriter->endElement();
        }

        // Second, write column breaks
        if (!empty($aColumnBreaks)) {
            $phpExcelSharedXMLWriter->startElement('colBreaks');
            $phpExcelSharedXMLWriter->writeAttribute('count', \count($aColumnBreaks));
            $phpExcelSharedXMLWriter->writeAttribute('manualBreakCount', \count($aColumnBreaks));

            foreach ($aColumnBreaks as $aColumnBreak) {
                $coords = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::coordinateFromString($aColumnBreak);

                $phpExcelSharedXMLWriter->startElement('brk');
                $phpExcelSharedXMLWriter->writeAttribute('id', \PhpOffice\PhpSpreadsheet\Cell\Coordinate::columnIndexFromString($coords[0]) - 1);
                $phpExcelSharedXMLWriter->writeAttribute('man', '1');
                $phpExcelSharedXMLWriter->endElement();
            }

            $phpExcelSharedXMLWriter->endElement();
        }
    }

    /**
     * Write SheetData
     *
     * @param \PhpOffice\PhpSpreadsheet\Shared\XMLWriter        $phpExcelSharedXMLWriter        XML Writer
     * @param \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet                $phpExcelWorksheet            Worksheet
     * @param    string[]                        $pStringTable    String table
     * @throws    \PhpOffice\PhpSpreadsheet\Writer\Exception
     */
    private function writeSheetData(\PhpOffice\PhpSpreadsheet\Shared\XMLWriter $phpExcelSharedXMLWriter = \null, \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet $phpExcelWorksheet = \null, $pStringTable = \null)
    {
        if (\is_array($pStringTable)) {
            // Flipped stringtable, for faster index searching
            $aFlippedStringTable = $this->getParentWriter()->getWriterPart('stringtable')->flipStringTable($pStringTable);

            // sheetData
            $phpExcelSharedXMLWriter->startElement('sheetData');

            // Get column count
            $colCount = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::columnIndexFromString($phpExcelWorksheet->getHighestColumn());

            // Highest row number
            $highestRow = $phpExcelWorksheet->getHighestRow();

            // Loop through cells
            $cellsByRow = array();
            foreach ($phpExcelWorksheet->getCellCollection() as $cellID) {
                $cellAddress = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::coordinateFromString($cellID);
                $cellsByRow[$cellAddress[1]][] = $cellID;
            }

            $currentRow = 0;
            while ($currentRow++ < $highestRow) {
                // Get row dimension
                $rowDimension = $phpExcelWorksheet->getRowDimension($currentRow);

                // Write current row?
                $writeCurrentRow = isset($cellsByRow[$currentRow]) || $rowDimension->getRowHeight() >= 0 || $rowDimension->isVisible() == \false || $rowDimension->isCollapsed() == \true || $rowDimension->getOutlineLevel() > 0 || $rowDimension->getXfIndex() !== \null;

                if ($writeCurrentRow) {
                    // Start a new row
                    $phpExcelSharedXMLWriter->startElement('row');
                    $phpExcelSharedXMLWriter->writeAttribute('r', $currentRow);
                    $phpExcelSharedXMLWriter->writeAttribute('spans', '1:' . $colCount);

                    // Row dimensions
                    if ($rowDimension->getRowHeight() >= 0) {
                        $phpExcelSharedXMLWriter->writeAttribute('customHeight', '1');
                        $phpExcelSharedXMLWriter->writeAttribute('ht', \PhpOffice\PhpSpreadsheet\Shared\StringHelper::FormatNumber($rowDimension->getRowHeight()));
                    }

                    // Row visibility
                    if ($rowDimension->isVisible() == \false) {
                        $phpExcelSharedXMLWriter->writeAttribute('hidden', 'true');
                    }

                    // Collapsed
                    if ($rowDimension->isCollapsed() == \true) {
                        $phpExcelSharedXMLWriter->writeAttribute('collapsed', 'true');
                    }

                    // Outline level
                    if ($rowDimension->getOutlineLevel() > 0) {
                        $phpExcelSharedXMLWriter->writeAttribute('outlineLevel', $rowDimension->getOutlineLevel());
                    }

                    // Style
                    if ($rowDimension->getXfIndex() !== \null) {
                        $phpExcelSharedXMLWriter->writeAttribute('s', $rowDimension->getXfIndex());
                        $phpExcelSharedXMLWriter->writeAttribute('customFormat', '1');
                    }

                    // Write cells
                    if (isset($cellsByRow[$currentRow])) {
                        foreach ($cellsByRow[$currentRow] as $cellAddress) {
                            // Write cell
                            $this->writeCell($phpExcelSharedXMLWriter, $phpExcelWorksheet, $cellAddress, $pStringTable, $aFlippedStringTable);
                        }
                    }

                    // End row
                    $phpExcelSharedXMLWriter->endElement();
                }
            }

            $phpExcelSharedXMLWriter->endElement();
        } else {
            throw new \PhpOffice\PhpSpreadsheet\Writer\Exception("Invalid parameters passed.");
        }
    }

    /**
     * Write Cell
     *
     * @param \PhpOffice\PhpSpreadsheet\Shared\XMLWriter    $phpExcelSharedXMLWriter                XML Writer
     * @param \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet            $phpExcelWorksheet                    Worksheet
     * @param    \PhpOffice\PhpSpreadsheet\Cell\Cell                $pCellAddress            Cell Address
     * @param    string[]                    $pStringTable            String table
     * @param    string[]                    $pFlippedStringTable    String table (flipped), for faster index searching
     * @throws    \PhpOffice\PhpSpreadsheet\Writer\Exception
     */
    private function writeCell(\PhpOffice\PhpSpreadsheet\Shared\XMLWriter $phpExcelSharedXMLWriter = \null, \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet $phpExcelWorksheet = \null, $pCellAddress = \null, $pStringTable = \null, $pFlippedStringTable = \null)
    {
        if (\is_array($pStringTable) && \is_array($pFlippedStringTable)) {
            // Cell
            $pCell = $phpExcelWorksheet->getCell($pCellAddress);
            $phpExcelSharedXMLWriter->startElement('c');
            $phpExcelSharedXMLWriter->writeAttribute('r', $pCellAddress);

            // Sheet styles
            if ($pCell->getXfIndex() != '') {
                $phpExcelSharedXMLWriter->writeAttribute('s', $pCell->getXfIndex());
            }

            // If cell value is supplied, write cell value
            $cellValue = $pCell->getValue();
            if (\is_object($cellValue) || $cellValue !== '') {
                // Map type
                $dataType = $pCell->getDataType();

                // Write data type depending on its type
                switch (\strtolower($dataType)) {
                    case 'inlinestr':    // Inline string
                    case 's':            // String
                    case 'b':            // Boolean
                        $phpExcelSharedXMLWriter->writeAttribute('t', $dataType);
                        break;
                    case 'f':            // Formula
                        $calculatedValue = ($this->getParentWriter()->getPreCalculateFormulas()) ?
                            $pCell->getCalculatedValue() :
                            $cellValue;
                        if (\is_string($calculatedValue)) {
                            $phpExcelSharedXMLWriter->writeAttribute('t', 'str');
                        }
                        break;
                    case 'e':            // Error
                        $phpExcelSharedXMLWriter->writeAttribute('t', $dataType);
                }

                // Write data depending on its type
                switch (\strtolower($dataType)) {
                    case 'inlinestr':    // Inline string
                        if (! $cellValue instanceof \PhpOffice\PhpSpreadsheet\RichText\RichText) {
                            $phpExcelSharedXMLWriter->writeElement('t', \PhpOffice\PhpSpreadsheet\Shared\StringHelper::ControlCharacterPHP2OOXML(\htmlspecialchars($cellValue)));
                        } elseif ($cellValue instanceof \PhpOffice\PhpSpreadsheet\RichText\RichText) {
                            $phpExcelSharedXMLWriter->startElement('is');
                            $this->getParentWriter()->getWriterPart('stringtable')->writeRichText($phpExcelSharedXMLWriter, $cellValue);
                            $phpExcelSharedXMLWriter->endElement();
                        }

                        break;
                    case 's':            // String
                        if (! $cellValue instanceof \PhpOffice\PhpSpreadsheet\RichText\RichText) {
                            if (isset($pFlippedStringTable[$cellValue])) {
                                $phpExcelSharedXMLWriter->writeElement('v', $pFlippedStringTable[$cellValue]);
                            }
                        } elseif ($cellValue instanceof \PhpOffice\PhpSpreadsheet\RichText\RichText) {
                            $phpExcelSharedXMLWriter->writeElement('v', $pFlippedStringTable[$cellValue->getHashCode()]);
                        }

                        break;
                    case 'f':            // Formula
                        $attributes = $pCell->getFormulaAttributes();
                        if ($attributes['t'] == 'array') {
                            $phpExcelSharedXMLWriter->startElement('f');
                            $phpExcelSharedXMLWriter->writeAttribute('t', 'array');
                            $phpExcelSharedXMLWriter->writeAttribute('ref', $pCellAddress);
                            $phpExcelSharedXMLWriter->writeAttribute('aca', '1');
                            $phpExcelSharedXMLWriter->writeAttribute('ca', '1');
                            $phpExcelSharedXMLWriter->text(\substr($cellValue, 1));
                            $phpExcelSharedXMLWriter->endElement();
                        } else {
                            $phpExcelSharedXMLWriter->writeElement('f', \substr($cellValue, 1));
                        }
                        if ($this->getParentWriter()->getOffice2003Compatibility() === \false) {
                            if ($this->getParentWriter()->getPreCalculateFormulas()) {
//                                $calculatedValue = $pCell->getCalculatedValue();
                                if (!\is_array($calculatedValue) && \substr($calculatedValue, 0, 1) != '#') {
                                    $phpExcelSharedXMLWriter->writeElement('v', \PhpOffice\PhpSpreadsheet\Shared\StringHelper::FormatNumber($calculatedValue));
                                } else {
                                    $phpExcelSharedXMLWriter->writeElement('v', '0');
                                }
                            } else {
                                $phpExcelSharedXMLWriter->writeElement('v', '0');
                            }
                        }
                        break;
                    case 'n':            // Numeric
                        // force point as decimal separator in case current locale uses comma
                        $phpExcelSharedXMLWriter->writeElement('v', \str_replace(',', '.', $cellValue));
                        break;
                    case 'b':            // Boolean
                        $phpExcelSharedXMLWriter->writeElement('v', ($cellValue ? '1' : '0'));
                        break;
                    case 'e':            // Error
                        if (\substr($cellValue, 0, 1) == '=') {
                            $phpExcelSharedXMLWriter->writeElement('f', \substr($cellValue, 1));
                            $phpExcelSharedXMLWriter->writeElement('v', \substr($cellValue, 1));
                        } else {
                            $phpExcelSharedXMLWriter->writeElement('v', $cellValue);
                        }

                        break;
                }
            }

            $phpExcelSharedXMLWriter->endElement();
        } else {
            throw new \PhpOffice\PhpSpreadsheet\Writer\Exception("Invalid parameters passed.");
        }
    }

    /**
     * Write Drawings
     *
     * @param \PhpOffice\PhpSpreadsheet\Shared\XMLWriter    $phpExcelSharedXMLWriter        XML Writer
     * @param \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet            $phpExcelWorksheet            Worksheet
     * @param    boolean                        $includeCharts    Flag indicating if we should include drawing details for charts
     * @throws    \PhpOffice\PhpSpreadsheet\Writer\Exception
     */
    private function writeDrawings(\PhpOffice\PhpSpreadsheet\Shared\XMLWriter $phpExcelSharedXMLWriter = \null, \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet $phpExcelWorksheet = \null, $includeCharts = \false)
    {
        $chartCount = ($includeCharts) ? $phpExcelWorksheet->getChartCollection()->count() : 0;
        // If sheet contains drawings, add the relationships
        if (($phpExcelWorksheet->getDrawingCollection()->count() > 0) ||
            ($chartCount > 0)) {
            $phpExcelSharedXMLWriter->startElement('drawing');
            $phpExcelSharedXMLWriter->writeAttribute('r:id', 'rId1');
            $phpExcelSharedXMLWriter->endElement();
        }
    }

    /**
     * Write LegacyDrawing
     *
     * @param \PhpOffice\PhpSpreadsheet\Shared\XMLWriter        $phpExcelSharedXMLWriter        XML Writer
     * @param \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet                $phpExcelWorksheet            Worksheet
     * @throws    \PhpOffice\PhpSpreadsheet\Writer\Exception
     */
    private function writeLegacyDrawing(\PhpOffice\PhpSpreadsheet\Shared\XMLWriter $phpExcelSharedXMLWriter = \null, \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet $phpExcelWorksheet = \null)
    {
        // If sheet contains comments, add the relationships
        if (\count($phpExcelWorksheet->getComments()) > 0) {
            $phpExcelSharedXMLWriter->startElement('legacyDrawing');
            $phpExcelSharedXMLWriter->writeAttribute('r:id', 'rId_comments_vml1');
            $phpExcelSharedXMLWriter->endElement();
        }
    }

    /**
     * Write LegacyDrawingHF
     *
     * @param \PhpOffice\PhpSpreadsheet\Shared\XMLWriter        $phpExcelSharedXMLWriter        XML Writer
     * @param \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet                $phpExcelWorksheet            Worksheet
     * @throws    \PhpOffice\PhpSpreadsheet\Writer\Exception
     */
    private function writeLegacyDrawingHF(\PhpOffice\PhpSpreadsheet\Shared\XMLWriter $phpExcelSharedXMLWriter = \null, \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet $phpExcelWorksheet = \null)
    {
        // If sheet contains images, add the relationships
        if (\count($phpExcelWorksheet->getHeaderFooter()->getImages()) > 0) {
            $phpExcelSharedXMLWriter->startElement('legacyDrawingHF');
            $phpExcelSharedXMLWriter->writeAttribute('r:id', 'rId_headerfooter_vml1');
            $phpExcelSharedXMLWriter->endElement();
        }
    }
}

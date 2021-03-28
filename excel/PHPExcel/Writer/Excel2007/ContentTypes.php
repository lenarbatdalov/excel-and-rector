<?php

namespace PhpOffice\PhpSpreadsheet\Writer\Xlsx;

/**
 * PHPExcel_Writer_Excel2007_ContentTypes
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
class ContentTypes extends \PhpOffice\PhpSpreadsheet\Writer\Xlsx\WriterPart
{
    /**
     * Write content types to XML format
     *
     * @param    boolean        $includeCharts    Flag indicating if we should include drawing details for charts
     * @return     string                         XML Output
     * @throws     \PhpOffice\PhpSpreadsheet\Writer\Exception
     */
    public function writeContentTypes(\PhpOffice\PhpSpreadsheet\Spreadsheet $phpExcel = \null, $includeCharts = \false)
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

        // Types
        $objWriter->startElement('Types');
        $objWriter->writeAttribute('xmlns', 'http://schemas.openxmlformats.org/package/2006/content-types');

        // Theme
        $this->writeOverrideContentType($objWriter, '/xl/theme/theme1.xml', 'application/vnd.openxmlformats-officedocument.theme+xml');

        // Styles
        $this->writeOverrideContentType($objWriter, '/xl/styles.xml', 'application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml');

        // Rels
        $this->writeDefaultContentType($objWriter, 'rels', 'application/vnd.openxmlformats-package.relationships+xml');

        // XML
        $this->writeDefaultContentType($objWriter, 'xml', 'application/xml');

        // VML
        $this->writeDefaultContentType($objWriter, 'vml', 'application/vnd.openxmlformats-officedocument.vmlDrawing');

        // Workbook
        if ($phpExcel->hasMacros()) { //Macros in workbook ?
            // Yes : not standard content but "macroEnabled"
            $this->writeOverrideContentType($objWriter, '/xl/workbook.xml', 'application/vnd.ms-excel.sheet.macroEnabled.main+xml');
            //... and define a new type for the VBA project
            $this->writeDefaultContentType($objWriter, 'bin', 'application/vnd.ms-office.vbaProject');
            if ($phpExcel->hasMacrosCertificate()) {// signed macros ?
                // Yes : add needed information
                $this->writeOverrideContentType($objWriter, '/xl/vbaProjectSignature.bin', 'application/vnd.ms-office.vbaProjectSignature');
            }
        } else {// no macros in workbook, so standard type
            $this->writeOverrideContentType($objWriter, '/xl/workbook.xml', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml');
        }

        // DocProps
        $this->writeOverrideContentType($objWriter, '/docProps/app.xml', 'application/vnd.openxmlformats-officedocument.extended-properties+xml');

        $this->writeOverrideContentType($objWriter, '/docProps/core.xml', 'application/vnd.openxmlformats-package.core-properties+xml');

        $customPropertyList = $phpExcel->getProperties()->getCustomProperties();
        if (!empty($customPropertyList)) {
            $this->writeOverrideContentType($objWriter, '/docProps/custom.xml', 'application/vnd.openxmlformats-officedocument.custom-properties+xml');
        }

        // Worksheets
        $sheetCount = $phpExcel->getSheetCount();
        for ($i = 0; $i < $sheetCount; ++$i) {
            $this->writeOverrideContentType($objWriter, '/xl/worksheets/sheet' . ($i + 1) . '.xml', 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml');
        }

        // Shared strings
        $this->writeOverrideContentType($objWriter, '/xl/sharedStrings.xml', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml');

        // Add worksheet relationship content types
        $chart = 1;
        for ($i = 0; $i < $sheetCount; ++$i) {
            $drawings = $phpExcel->getSheet($i)->getDrawingCollection();
            $drawingCount = \count($drawings);
            $chartCount = ($includeCharts) ? $phpExcel->getSheet($i)->getChartCount() : 0;

            //    We need a drawing relationship for the worksheet if we have either drawings or charts
            if (($drawingCount > 0) || ($chartCount > 0)) {
                $this->writeOverrideContentType($objWriter, '/xl/drawings/drawing' . ($i + 1) . '.xml', 'application/vnd.openxmlformats-officedocument.drawing+xml');
            }

            //    If we have charts, then we need a chart relationship for every individual chart
            if ($chartCount > 0) {
                for ($c = 0; $c < $chartCount; ++$c) {
                    $this->writeOverrideContentType($objWriter, '/xl/charts/chart' . $chart++ . '.xml', 'application/vnd.openxmlformats-officedocument.drawingml.chart+xml');
                }
            }
        }

        // Comments
        for ($i = 0; $i < $sheetCount; ++$i) {
            if (\count($phpExcel->getSheet($i)->getComments()) > 0) {
                $this->writeOverrideContentType($objWriter, '/xl/comments' . ($i + 1) . '.xml', 'application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml');
            }
        }

        // Add media content-types
        $aMediaContentTypes = array();
        $mediaCount = $this->getParentWriter()->getDrawingHashTable()->count();
        for ($i = 0; $i < $mediaCount; ++$i) {
            $extension     = '';
            $mimeType     = '';

            if ($this->getParentWriter()->getDrawingHashTable()->getByIndex($i) instanceof \PhpOffice\PhpSpreadsheet\Worksheet\Drawing) {
                $extension = \strtolower($this->getParentWriter()->getDrawingHashTable()->getByIndex($i)->getExtension());
                $mimeType = $this->getImageMimeType($this->getParentWriter()->getDrawingHashTable()->getByIndex($i)->getPath());
            } elseif ($this->getParentWriter()->getDrawingHashTable()->getByIndex($i) instanceof \PhpOffice\PhpSpreadsheet\Worksheet\MemoryDrawing) {
                $extension = \strtolower($this->getParentWriter()->getDrawingHashTable()->getByIndex($i)->getMimeType());
                $extension = \explode('/', $extension);
                $extension = $extension[1];

                $mimeType = $this->getParentWriter()->getDrawingHashTable()->getByIndex($i)->getMimeType();
            }

            if (!isset( $aMediaContentTypes[$extension])) {
                $aMediaContentTypes[$extension] = $mimeType;

                $this->writeDefaultContentType($objWriter, $extension, $mimeType);
            }
        }
        if ($phpExcel->hasRibbonBinObjects()) {
            // Some additional objects in the ribbon ?
            // we need to write "Extension" but not already write for media content
            $tabRibbonTypes=\array_diff($phpExcel->getRibbonBinObjects('types'), \array_keys($aMediaContentTypes));
            foreach ($tabRibbonTypes as $tabRibbonType) {
                $mimeType='image/.'.$tabRibbonType;//we wrote $mimeType like customUI Editor
                $this->writeDefaultContentType($objWriter, $tabRibbonType, $mimeType);
            }
        }
        $sheetCount = $phpExcel->getSheetCount();
        for ($i = 0; $i < $sheetCount; ++$i) {
            foreach ($phpExcel->getSheet()->getHeaderFooter()->getImages() as $phpExcelWorksheetHeaderFooterDrawing) {
                if (!isset( $aMediaContentTypes[\strtolower($phpExcelWorksheetHeaderFooterDrawing->getExtension())])) {
                    $aMediaContentTypes[\strtolower($phpExcelWorksheetHeaderFooterDrawing->getExtension())] = $this->getImageMimeType($phpExcelWorksheetHeaderFooterDrawing->getPath());

                    $this->writeDefaultContentType($objWriter, \strtolower($phpExcelWorksheetHeaderFooterDrawing->getExtension()), $aMediaContentTypes[\strtolower($phpExcelWorksheetHeaderFooterDrawing->getExtension())]);
                }
            }
        }

        $objWriter->endElement();

        // Return
        return $objWriter->getData();
    }

    /**
     * Get image mime type
     *
     * @param     string    $pFile    Filename
     * @return     string    Mime Type
     * @throws     \PhpOffice\PhpSpreadsheet\Writer\Exception
     */
    private function getImageMimeType($pFile = '')
    {
        if (\PhpOffice\PhpSpreadsheet\Shared\File::file_exists($pFile)) {
            $image = \getimagesize($pFile);
            return \image_type_to_mime_type($image[2]);
        } else {
            throw new \PhpOffice\PhpSpreadsheet\Writer\Exception("File $pFile does not exist");
        }
    }

    /**
     * Write Default content type
     *
     * @param \PhpOffice\PhpSpreadsheet\Shared\XMLWriter     $phpExcelSharedXMLWriter         XML Writer
     * @param     string                         $pPartname         Part name
     * @param     string                         $pContentType     Content type
     * @throws     \PhpOffice\PhpSpreadsheet\Writer\Exception
     */
    private function writeDefaultContentType(\PhpOffice\PhpSpreadsheet\Shared\XMLWriter $phpExcelSharedXMLWriter = \null, $pPartname = '', $pContentType = '')
    {
        if ($pPartname != '' && $pContentType != '') {
            // Write content type
            $phpExcelSharedXMLWriter->startElement('Default');
            $phpExcelSharedXMLWriter->writeAttribute('Extension', $pPartname);
            $phpExcelSharedXMLWriter->writeAttribute('ContentType', $pContentType);
            $phpExcelSharedXMLWriter->endElement();
        } else {
            throw new \PhpOffice\PhpSpreadsheet\Writer\Exception("Invalid parameters passed.");
        }
    }

    /**
     * Write Override content type
     *
     * @param \PhpOffice\PhpSpreadsheet\Shared\XMLWriter     $phpExcelSharedXMLWriter         XML Writer
     * @param     string                         $pPartname         Part name
     * @param     string                         $pContentType     Content type
     * @throws     \PhpOffice\PhpSpreadsheet\Writer\Exception
     */
    private function writeOverrideContentType(\PhpOffice\PhpSpreadsheet\Shared\XMLWriter $phpExcelSharedXMLWriter = \null, $pPartname = '', $pContentType = '')
    {
        if ($pPartname != '' && $pContentType != '') {
            // Write content type
            $phpExcelSharedXMLWriter->startElement('Override');
            $phpExcelSharedXMLWriter->writeAttribute('PartName', $pPartname);
            $phpExcelSharedXMLWriter->writeAttribute('ContentType', $pContentType);
            $phpExcelSharedXMLWriter->endElement();
        } else {
            throw new \PhpOffice\PhpSpreadsheet\Writer\Exception("Invalid parameters passed.");
        }
    }
}

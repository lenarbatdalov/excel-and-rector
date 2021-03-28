<?php
namespace PhpOffice\PhpSpreadsheet\Writer\Ods;

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
 * @package    PHPExcel_Writer_OpenDocument
 * @copyright  Copyright (c) 2006 - 2015 PHPExcel (http://www.codeplex.com/PHPExcel)
 * @license    http://www.gnu.org/licenses/old-licenses/lgpl-2.1.txt    LGPL
 * @version    ##VERSION##, ##DATE##
 */
/**
 * PHPExcel_Writer_OpenDocument_Content
 *
 * @category   PHPExcel
 * @package    PHPExcel_Writer_OpenDocument
 * @copyright  Copyright (c) 2006 - 2015 PHPExcel (http://www.codeplex.com/PHPExcel)
 * @author     Alexander Pervakov <frost-nzcr4@jagmort.com>
 */
class Content extends \PhpOffice\PhpSpreadsheet\Writer\Ods\WriterPart
{
    const NUMBER_COLS_REPEATED_MAX = 1024;
    const NUMBER_ROWS_REPEATED_MAX = 1048576;

    /**
     * Write content.xml to XML format
     *
     * @return  string                     XML Output
     * @throws  \PhpOffice\PhpSpreadsheet\Writer\Exception
     */
    public function write(\PhpOffice\PhpSpreadsheet\Spreadsheet $phpExcel = \null)
    {
        if (!$phpExcel) {
            $phpExcel = $this->getParentWriter()->getPHPExcel(); /* @var $pPHPExcel PHPExcel */
        }

        $objWriter = \null;
        if ($this->getParentWriter()->getUseDiskCaching()) {
            $objWriter = new \PhpOffice\PhpSpreadsheet\Shared\XMLWriter(\PhpOffice\PhpSpreadsheet\Shared\XMLWriter::STORAGE_DISK, $this->getParentWriter()->getDiskCachingDirectory());
        } else {
            $objWriter = new \PhpOffice\PhpSpreadsheet\Shared\XMLWriter(\PhpOffice\PhpSpreadsheet\Shared\XMLWriter::STORAGE_MEMORY);
        }

        // XML header
        $objWriter->startDocument('1.0', 'UTF-8');

        // Content
        $objWriter->startElement('office:document-content');
            $objWriter->writeAttribute('xmlns:office', 'urn:oasis:names:tc:opendocument:xmlns:office:1.0');
            $objWriter->writeAttribute('xmlns:style', 'urn:oasis:names:tc:opendocument:xmlns:style:1.0');
            $objWriter->writeAttribute('xmlns:text', 'urn:oasis:names:tc:opendocument:xmlns:text:1.0');
            $objWriter->writeAttribute('xmlns:table', 'urn:oasis:names:tc:opendocument:xmlns:table:1.0');
            $objWriter->writeAttribute('xmlns:draw', 'urn:oasis:names:tc:opendocument:xmlns:drawing:1.0');
            $objWriter->writeAttribute('xmlns:fo', 'urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0');
            $objWriter->writeAttribute('xmlns:xlink', 'http://www.w3.org/1999/xlink');
            $objWriter->writeAttribute('xmlns:dc', 'http://purl.org/dc/elements/1.1/');
            $objWriter->writeAttribute('xmlns:meta', 'urn:oasis:names:tc:opendocument:xmlns:meta:1.0');
            $objWriter->writeAttribute('xmlns:number', 'urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0');
            $objWriter->writeAttribute('xmlns:presentation', 'urn:oasis:names:tc:opendocument:xmlns:presentation:1.0');
            $objWriter->writeAttribute('xmlns:svg', 'urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0');
            $objWriter->writeAttribute('xmlns:chart', 'urn:oasis:names:tc:opendocument:xmlns:chart:1.0');
            $objWriter->writeAttribute('xmlns:dr3d', 'urn:oasis:names:tc:opendocument:xmlns:dr3d:1.0');
            $objWriter->writeAttribute('xmlns:math', 'http://www.w3.org/1998/Math/MathML');
            $objWriter->writeAttribute('xmlns:form', 'urn:oasis:names:tc:opendocument:xmlns:form:1.0');
            $objWriter->writeAttribute('xmlns:script', 'urn:oasis:names:tc:opendocument:xmlns:script:1.0');
            $objWriter->writeAttribute('xmlns:ooo', 'http://openoffice.org/2004/office');
            $objWriter->writeAttribute('xmlns:ooow', 'http://openoffice.org/2004/writer');
            $objWriter->writeAttribute('xmlns:oooc', 'http://openoffice.org/2004/calc');
            $objWriter->writeAttribute('xmlns:dom', 'http://www.w3.org/2001/xml-events');
            $objWriter->writeAttribute('xmlns:xforms', 'http://www.w3.org/2002/xforms');
            $objWriter->writeAttribute('xmlns:xsd', 'http://www.w3.org/2001/XMLSchema');
            $objWriter->writeAttribute('xmlns:xsi', 'http://www.w3.org/2001/XMLSchema-instance');
            $objWriter->writeAttribute('xmlns:rpt', 'http://openoffice.org/2005/report');
            $objWriter->writeAttribute('xmlns:of', 'urn:oasis:names:tc:opendocument:xmlns:of:1.2');
            $objWriter->writeAttribute('xmlns:xhtml', 'http://www.w3.org/1999/xhtml');
            $objWriter->writeAttribute('xmlns:grddl', 'http://www.w3.org/2003/g/data-view#');
            $objWriter->writeAttribute('xmlns:tableooo', 'http://openoffice.org/2009/table');
            $objWriter->writeAttribute('xmlns:field', 'urn:openoffice:names:experimental:ooo-ms-interop:xmlns:field:1.0');
            $objWriter->writeAttribute('xmlns:formx', 'urn:openoffice:names:experimental:ooxml-odf-interop:xmlns:form:1.0');
            $objWriter->writeAttribute('xmlns:css3t', 'http://www.w3.org/TR/css3-text/');
            $objWriter->writeAttribute('office:version', '1.2');

            $objWriter->writeElement('office:scripts');
            $objWriter->writeElement('office:font-face-decls');
            $objWriter->writeElement('office:automatic-styles');

            $objWriter->startElement('office:body');
                $objWriter->startElement('office:spreadsheet');
                    $objWriter->writeElement('table:calculation-settings');
                    $this->writeSheets($objWriter);
                    $objWriter->writeElement('table:named-expressions');
                $objWriter->endElement();
            $objWriter->endElement();
        $objWriter->endElement();

        return $objWriter->getData();
    }

    /**
     * Write sheets
     */
    private function writeSheets(\PhpOffice\PhpSpreadsheet\Shared\XMLWriter $phpExcelSharedXMLWriter)
    {
        $pPHPExcel = $this->getParentWriter()->getPHPExcel(); /* @var $pPHPExcel PHPExcel */

        $sheet_count = $pPHPExcel->getSheetCount();
        for ($i = 0; $i < $sheet_count; $i++) {
            //$this->getWriterPart('Worksheet')->writeWorksheet());
            $phpExcelSharedXMLWriter->startElement('table:table');
                $phpExcelSharedXMLWriter->writeAttribute('table:name', $pPHPExcel->getSheet($i)->getTitle());
                $phpExcelSharedXMLWriter->writeElement('office:forms');
                $phpExcelSharedXMLWriter->startElement('table:table-column');
                    $phpExcelSharedXMLWriter->writeAttribute('table:number-columns-repeated', self::NUMBER_COLS_REPEATED_MAX);
                $phpExcelSharedXMLWriter->endElement();
                $this->writeRows($phpExcelSharedXMLWriter, $pPHPExcel->getSheet($i));
            $phpExcelSharedXMLWriter->endElement();
        }
    }

    /**
     * Write rows of the specified sheet
     */
    private function writeRows(\PhpOffice\PhpSpreadsheet\Shared\XMLWriter $phpExcelSharedXMLWriter, \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet $phpExcelWorksheet)
    {
        $number_rows_repeated = self::NUMBER_ROWS_REPEATED_MAX;
        $span_row = 0;
        $rows = $phpExcelWorksheet->getRowIterator();
        while ($rows->valid()) {
            $number_rows_repeated--;
            $row = $rows->current();
            if ($row->getCellIterator()->valid()) {
                if ($span_row !== 0) {
                    $phpExcelSharedXMLWriter->startElement('table:table-row');
                    if ($span_row > 1) {
                        $phpExcelSharedXMLWriter->writeAttribute('table:number-rows-repeated', $span_row);
                    }
                    $phpExcelSharedXMLWriter->startElement('table:table-cell');
                        $phpExcelSharedXMLWriter->writeAttribute('table:number-columns-repeated', self::NUMBER_COLS_REPEATED_MAX);
                    $phpExcelSharedXMLWriter->endElement();
                    $phpExcelSharedXMLWriter->endElement();
                    $span_row = 0;
                }
                $phpExcelSharedXMLWriter->startElement('table:table-row');
                $this->writeCells($phpExcelSharedXMLWriter, $row);
                $phpExcelSharedXMLWriter->endElement();
            } else {
                $span_row++;
            }
            $rows->next();
        }
    }

    /**
     * Write cells of the specified row
     *
     * @throws \PhpOffice\PhpSpreadsheet\Writer\Exception
     */
    private function writeCells(\PhpOffice\PhpSpreadsheet\Shared\XMLWriter $phpExcelSharedXMLWriter, \PhpOffice\PhpSpreadsheet\Worksheet\Row $phpExcelWorksheetRow)
    {
        $number_cols_repeated = self::NUMBER_COLS_REPEATED_MAX;
        $prev_column = -1;
        $cells = $phpExcelWorksheetRow->getCellIterator();
        while ($cells->valid()) {
            $cell = $cells->current();
            $column = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::columnIndexFromString($cell->getColumn()) - 1;

            $this->writeCellSpan($phpExcelSharedXMLWriter, $column, $prev_column);
            $phpExcelSharedXMLWriter->startElement('table:table-cell');

            switch ($cell->getDataType()) {
                case \PhpOffice\PhpSpreadsheet\Cell\DataType::TYPE_BOOL:
                    $phpExcelSharedXMLWriter->writeAttribute('office:value-type', 'boolean');
                    $phpExcelSharedXMLWriter->writeAttribute('office:value', $cell->getValue());
                    $phpExcelSharedXMLWriter->writeElement('text:p', $cell->getValue());
                    break;

                case \PhpOffice\PhpSpreadsheet\Cell\DataType::TYPE_ERROR:
                    throw new \PhpOffice\PhpSpreadsheet\Writer\Exception('Writing of error not implemented yet.');

                case \PhpOffice\PhpSpreadsheet\Cell\DataType::TYPE_FORMULA:
                    try {
                        $formula_value = $cell->getCalculatedValue();
                    } catch (\Exception $e) {
                        $formula_value = $cell->getValue();
                    }
                    $phpExcelSharedXMLWriter->writeAttribute('table:formula', 'of:' . $cell->getValue());
                    if (\is_numeric($formula_value)) {
                        $phpExcelSharedXMLWriter->writeAttribute('office:value-type', 'float');
                    } else {
                        $phpExcelSharedXMLWriter->writeAttribute('office:value-type', 'string');
                    }
                    $phpExcelSharedXMLWriter->writeAttribute('office:value', $formula_value);
                    $phpExcelSharedXMLWriter->writeElement('text:p', $formula_value);
                    break;

                case \PhpOffice\PhpSpreadsheet\Cell\DataType::TYPE_INLINE:
                    throw new \PhpOffice\PhpSpreadsheet\Writer\Exception('Writing of inline not implemented yet.');

                case \PhpOffice\PhpSpreadsheet\Cell\DataType::TYPE_NUMERIC:
                    $phpExcelSharedXMLWriter->writeAttribute('office:value-type', 'float');
                    $phpExcelSharedXMLWriter->writeAttribute('office:value', $cell->getValue());
                    $phpExcelSharedXMLWriter->writeElement('text:p', $cell->getValue());
                    break;

                case \PhpOffice\PhpSpreadsheet\Cell\DataType::TYPE_STRING:
                    $phpExcelSharedXMLWriter->writeAttribute('office:value-type', 'string');
                    $phpExcelSharedXMLWriter->writeElement('text:p', $cell->getValue());
                    break;
            }
            \PhpOffice\PhpSpreadsheet\Writer\Ods\Cell\Comment::write($phpExcelSharedXMLWriter, $cell);
            $phpExcelSharedXMLWriter->endElement();
            $prev_column = $column;
            $cells->next();
        }
        $number_cols_repeated = $number_cols_repeated - $prev_column - 1;
        if ($number_cols_repeated > 0) {
            if ($number_cols_repeated > 1) {
                $phpExcelSharedXMLWriter->startElement('table:table-cell');
                $phpExcelSharedXMLWriter->writeAttribute('table:number-columns-repeated', $number_cols_repeated);
                $phpExcelSharedXMLWriter->endElement();
            } else {
                $phpExcelSharedXMLWriter->writeElement('table:table-cell');
            }
        }
    }

    /**
     * Write span
     *
     * @param integer $curColumn
     * @param integer $prevColumn
     */
    private function writeCellSpan(\PhpOffice\PhpSpreadsheet\Shared\XMLWriter $phpExcelSharedXMLWriter, $curColumn, $prevColumn)
    {
        $diff = $curColumn - $prevColumn - 1;
        if (1 === $diff) {
            $phpExcelSharedXMLWriter->writeElement('table:table-cell');
        } elseif ($diff > 1) {
            $phpExcelSharedXMLWriter->startElement('table:table-cell');
                $phpExcelSharedXMLWriter->writeAttribute('table:number-columns-repeated', $diff);
            $phpExcelSharedXMLWriter->endElement();
        }
    }
}

<?php

namespace PhpOffice\PhpSpreadsheet\Writer\Xlsx;

/**
 * PHPExcel_Writer_Excel2007_StringTable
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
class StringTable extends \PhpOffice\PhpSpreadsheet\Writer\Xlsx\WriterPart
{
    /**
     * Create worksheet stringtable
     *
     * @param     \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet     $pSheet                Worksheet
     * @param     string[]                 $pExistingTable     Existing table to eventually merge with
     * @return     string[]                 String table for worksheet
     * @throws     \PhpOffice\PhpSpreadsheet\Writer\Exception
     */
    public function createStringTable($pSheet = \null, $pExistingTable = \null)
    {
        if ($pSheet !== \null) {
            // Create string lookup table
            $aStringTable = array();
            $cellCollection = \null;
            $aFlippedStringTable = \null;    // For faster lookup

            // Is an existing table given?
            if (($pExistingTable !== \null) && \is_array($pExistingTable)) {
                $aStringTable = $pExistingTable;
            }

            // Fill index array
            $aFlippedStringTable = $this->flipStringTable($aStringTable);

            // Loop through cells
            foreach ($pSheet->getCoordinates() as $phpExcelCell) {
                $cell = $pSheet->getCell($phpExcelCell, true);
                $cellValue = $cell->getValue();
                if (!\is_object($cellValue) &&
                    ($cellValue !== \null) &&
                    $cellValue !== '' &&
                    !isset($aFlippedStringTable[$cellValue]) &&
                    ($cell->getDataType() == \PhpOffice\PhpSpreadsheet\Cell\DataType::TYPE_STRING || $cell->getDataType() == \PhpOffice\PhpSpreadsheet\Cell\DataType::TYPE_STRING2 || $cell->getDataType() == \PhpOffice\PhpSpreadsheet\Cell\DataType::TYPE_NULL)) {
                        $aStringTable[] = $cellValue;
                        $aFlippedStringTable[$cellValue] = \true;
                } elseif ($cellValue instanceof \PhpOffice\PhpSpreadsheet\RichText\RichText &&
                          ($cellValue !== \null) &&
                          !isset($aFlippedStringTable[$cellValue->getHashCode()])) {
                                $aStringTable[] = $cellValue;
                                $aFlippedStringTable[$cellValue->getHashCode()] = \true;
                }
            }

            return $aStringTable;
        } else {
            throw new \PhpOffice\PhpSpreadsheet\Writer\Exception("Invalid PHPExcel_Worksheet object passed.");
        }
    }

    /**
     * Write string table to XML format
     *
     * @param     string[]     $pStringTable
     * @return     string         XML Output
     * @throws     \PhpOffice\PhpSpreadsheet\Writer\Exception
     */
    public function writeStringTable($pStringTable = \null)
    {
        if ($pStringTable !== \null) {
            // Create XML writer
            $objWriter = \null;
            if ($this->getParentWriter()->getUseDiskCaching()) {
                $objWriter = new \PhpOffice\PhpSpreadsheet\Shared\XMLWriter(\PhpOffice\PhpSpreadsheet\Shared\XMLWriter::STORAGE_DISK, $this->getParentWriter()->getDiskCachingDirectory());
            } else {
                $objWriter = new \PhpOffice\PhpSpreadsheet\Shared\XMLWriter(\PhpOffice\PhpSpreadsheet\Shared\XMLWriter::STORAGE_MEMORY);
            }

            // XML header
            $objWriter->startDocument('1.0', 'UTF-8', 'yes');

            // String table
            $objWriter->startElement('sst');
            $objWriter->writeAttribute('xmlns', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main');
            $objWriter->writeAttribute('uniqueCount', \count($pStringTable));

            // Loop through string table
            foreach ($pStringTable as $singlePStringTable) {
                $objWriter->startElement('si');

                if (! $singlePStringTable instanceof \PhpOffice\PhpSpreadsheet\RichText\RichText) {
                    $textToWrite = \PhpOffice\PhpSpreadsheet\Shared\StringHelper::ControlCharacterPHP2OOXML($singlePStringTable);
                    $objWriter->startElement('t');
                    if ($textToWrite !== \trim($textToWrite)) {
                        $objWriter->writeAttribute('xml:space', 'preserve');
                    }
                    $objWriter->writeRawData($textToWrite);
                    $objWriter->endElement();
                } elseif ($singlePStringTable instanceof \PhpOffice\PhpSpreadsheet\RichText\RichText) {
                    $this->writeRichText($objWriter, $singlePStringTable);
                }

                $objWriter->endElement();
            }

            $objWriter->endElement();

            return $objWriter->getData();
        } else {
            throw new \PhpOffice\PhpSpreadsheet\Writer\Exception("Invalid string table array passed.");
        }
    }

    /**
     * Write Rich Text
     *
     * @param \PhpOffice\PhpSpreadsheet\Shared\XMLWriter    $phpExcelSharedXMLWriter         XML Writer
     * @param \PhpOffice\PhpSpreadsheet\RichText\RichText            $phpExcelRichText        Rich text
     * @param     string                        $prefix            Optional Namespace prefix
     * @throws     \PhpOffice\PhpSpreadsheet\Writer\Exception
     */
    public function writeRichText(\PhpOffice\PhpSpreadsheet\Shared\XMLWriter $phpExcelSharedXMLWriter = \null, \PhpOffice\PhpSpreadsheet\RichText\RichText $phpExcelRichText = \null, $prefix = \null)
    {
        if ($prefix !== \null) {
            $prefix .= ':';
        }
            
        // Loop through rich text elements
        $elements = $phpExcelRichText->getRichTextElements();
        foreach ($elements as $element) {
            // r
            $phpExcelSharedXMLWriter->startElement($prefix.'r');

            // rPr
            if ($element instanceof \PhpOffice\PhpSpreadsheet\RichText\Run) {
                // rPr
                $phpExcelSharedXMLWriter->startElement($prefix.'rPr');

                // rFont
                $phpExcelSharedXMLWriter->startElement($prefix.'rFont');
                $phpExcelSharedXMLWriter->writeAttribute('val', $element->getFont()->getName());
                $phpExcelSharedXMLWriter->endElement();

                // Bold
                $phpExcelSharedXMLWriter->startElement($prefix.'b');
                $phpExcelSharedXMLWriter->writeAttribute('val', ($element->getFont()->getBold() ? 'true' : 'false'));
                $phpExcelSharedXMLWriter->endElement();

                // Italic
                $phpExcelSharedXMLWriter->startElement($prefix.'i');
                $phpExcelSharedXMLWriter->writeAttribute('val', ($element->getFont()->getItalic() ? 'true' : 'false'));
                $phpExcelSharedXMLWriter->endElement();

                // Superscript / subscript
                if ($element->getFont()->getSuperScript() || $element->getFont()->getSubScript()) {
                    $phpExcelSharedXMLWriter->startElement($prefix.'vertAlign');
                    if ($element->getFont()->getSuperScript()) {
                        $phpExcelSharedXMLWriter->writeAttribute('val', 'superscript');
                    } elseif ($element->getFont()->getSubScript()) {
                        $phpExcelSharedXMLWriter->writeAttribute('val', 'subscript');
                    }
                    $phpExcelSharedXMLWriter->endElement();
                }

                // Strikethrough
                $phpExcelSharedXMLWriter->startElement($prefix.'strike');
                $phpExcelSharedXMLWriter->writeAttribute('val', ($element->getFont()->getStrikethrough() ? 'true' : 'false'));
                $phpExcelSharedXMLWriter->endElement();

                // Color
                $phpExcelSharedXMLWriter->startElement($prefix.'color');
                $phpExcelSharedXMLWriter->writeAttribute('rgb', $element->getFont()->getColor()->getARGB());
                $phpExcelSharedXMLWriter->endElement();

                // Size
                $phpExcelSharedXMLWriter->startElement($prefix.'sz');
                $phpExcelSharedXMLWriter->writeAttribute('val', $element->getFont()->getSize());
                $phpExcelSharedXMLWriter->endElement();

                // Underline
                $phpExcelSharedXMLWriter->startElement($prefix.'u');
                $phpExcelSharedXMLWriter->writeAttribute('val', $element->getFont()->getUnderline());
                $phpExcelSharedXMLWriter->endElement();

                $phpExcelSharedXMLWriter->endElement();
            }

            // t
            $phpExcelSharedXMLWriter->startElement($prefix.'t');
            $phpExcelSharedXMLWriter->writeAttribute('xml:space', 'preserve');
            $phpExcelSharedXMLWriter->writeRawData(\PhpOffice\PhpSpreadsheet\Shared\StringHelper::ControlCharacterPHP2OOXML($element->getText()));
            $phpExcelSharedXMLWriter->endElement();

            $phpExcelSharedXMLWriter->endElement();
        }
    }

    /**
     * Write Rich Text
     *
     * @param \PhpOffice\PhpSpreadsheet\Shared\XMLWriter    $phpExcelSharedXMLWriter         XML Writer
     * @param     string|\PhpOffice\PhpSpreadsheet\RichText\RichText    $pRichText        text string or Rich text
     * @param     string                        $prefix            Optional Namespace prefix
     * @throws     \PhpOffice\PhpSpreadsheet\Writer\Exception
     */
    public function writeRichTextForCharts(\PhpOffice\PhpSpreadsheet\Shared\XMLWriter $phpExcelSharedXMLWriter = \null, $pRichText = \null, $prefix = \null)
    {
        if (!$pRichText instanceof \PhpOffice\PhpSpreadsheet\RichText\RichText) {
            $textRun = $pRichText;
            $pRichText = new \PhpOffice\PhpSpreadsheet\RichText\RichText();
            $pRichText->createTextRun($textRun);
        }

        if ($prefix !== \null) {
            $prefix .= ':';
        }
            
        // Loop through rich text elements
        $elements = $pRichText->getRichTextElements();
        foreach ($elements as $element) {
            // r
            $phpExcelSharedXMLWriter->startElement($prefix.'r');

            // rPr
            $phpExcelSharedXMLWriter->startElement($prefix.'rPr');

            // Bold
            $phpExcelSharedXMLWriter->writeAttribute('b', ($element->getFont()->getBold() ? 1 : 0));
            // Italic
            $phpExcelSharedXMLWriter->writeAttribute('i', ($element->getFont()->getItalic() ? 1 : 0));
            // Underline
            $underlineType = $element->getFont()->getUnderline();
            switch ($underlineType) {
                case 'single':
                    $underlineType = 'sng';
                    break;
                case 'double':
                    $underlineType = 'dbl';
                    break;
            }
            $phpExcelSharedXMLWriter->writeAttribute('u', $underlineType);
            // Strikethrough
            $phpExcelSharedXMLWriter->writeAttribute('strike', ($element->getFont()->getStrikethrough() ? 'sngStrike' : 'noStrike'));

            // rFont
            $phpExcelSharedXMLWriter->startElement($prefix.'latin');
                $phpExcelSharedXMLWriter->writeAttribute('typeface', $element->getFont()->getName());
            $phpExcelSharedXMLWriter->endElement();

                // Superscript / subscript
//                    if ($element->getFont()->getSuperScript() || $element->getFont()->getSubScript()) {
//                        $objWriter->startElement($prefix.'vertAlign');
//                        if ($element->getFont()->getSuperScript()) {
//                            $objWriter->writeAttribute('val', 'superscript');
//                        } elseif ($element->getFont()->getSubScript()) {
//                            $objWriter->writeAttribute('val', 'subscript');
//                        }
//                        $objWriter->endElement();
//                    }
//
            $phpExcelSharedXMLWriter->endElement();

            // t
            $phpExcelSharedXMLWriter->startElement($prefix.'t');
//                    $objWriter->writeAttribute('xml:space', 'preserve');    //    Excel2010 accepts, Excel2007 complains
            $phpExcelSharedXMLWriter->writeRawData(\PhpOffice\PhpSpreadsheet\Shared\StringHelper::ControlCharacterPHP2OOXML($element->getText()));
            $phpExcelSharedXMLWriter->endElement();

            $phpExcelSharedXMLWriter->endElement();
        }
    }

    /**
     * Flip string table (for index searching)
     *
     * @param     array    $stringTable    Stringtable
     * @return     array
     */
    public function flipStringTable($stringTable = array())
    {
        // Return value
        $returnValue = array();

        // Loop through stringtable and add flipped items to $returnValue
        foreach ($stringTable as $key => $value) {
            if (! $value instanceof \PhpOffice\PhpSpreadsheet\RichText\RichText) {
                $returnValue[$value] = $key;
            } elseif ($value instanceof \PhpOffice\PhpSpreadsheet\RichText\RichText) {
                $returnValue[$value->getHashCode()] = $key;
            }
        }

        return $returnValue;
    }
}

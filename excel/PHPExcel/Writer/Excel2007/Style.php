<?php

namespace PhpOffice\PhpSpreadsheet\Writer\Xlsx;

/**
 * PHPExcel_Writer_Excel2007_Style
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
class Style extends \PhpOffice\PhpSpreadsheet\Writer\Xlsx\WriterPart
{
    /**
     * Write styles to XML format
     *
     * @return     string         XML Output
     * @throws     \PhpOffice\PhpSpreadsheet\Writer\Exception
     */
    public function writeStyles(\PhpOffice\PhpSpreadsheet\Spreadsheet $phpExcel = \null)
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

        // styleSheet
        $objWriter->startElement('styleSheet');
        $objWriter->writeAttribute('xml:space', 'preserve');
        $objWriter->writeAttribute('xmlns', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main');

        // numFmts
        $objWriter->startElement('numFmts');
        $objWriter->writeAttribute('count', $this->getParentWriter()->getNumFmtHashTable()->count());

        // numFmt
        for ($i = 0; $i < $this->getParentWriter()->getNumFmtHashTable()->count(); ++$i) {
            $this->writeNumFmt($objWriter, $this->getParentWriter()->getNumFmtHashTable()->getByIndex($i), $i);
        }

        $objWriter->endElement();

        // fonts
        $objWriter->startElement('fonts');
        $objWriter->writeAttribute('count', $this->getParentWriter()->getFontHashTable()->count());

        // font
        for ($i = 0; $i < $this->getParentWriter()->getFontHashTable()->count(); ++$i) {
            $this->writeFont($objWriter, $this->getParentWriter()->getFontHashTable()->getByIndex($i));
        }

        $objWriter->endElement();

        // fills
        $objWriter->startElement('fills');
        $objWriter->writeAttribute('count', $this->getParentWriter()->getFillHashTable()->count());

        // fill
        for ($i = 0; $i < $this->getParentWriter()->getFillHashTable()->count(); ++$i) {
            $this->writeFill($objWriter, $this->getParentWriter()->getFillHashTable()->getByIndex($i));
        }

        $objWriter->endElement();

        // borders
        $objWriter->startElement('borders');
        $objWriter->writeAttribute('count', $this->getParentWriter()->getBordersHashTable()->count());

        // border
        for ($i = 0; $i < $this->getParentWriter()->getBordersHashTable()->count(); ++$i) {
            $this->writeBorder($objWriter, $this->getParentWriter()->getBordersHashTable()->getByIndex($i));
        }

        $objWriter->endElement();

        // cellStyleXfs
        $objWriter->startElement('cellStyleXfs');
        $objWriter->writeAttribute('count', 1);

        // xf
        $objWriter->startElement('xf');
        $objWriter->writeAttribute('numFmtId', 0);
        $objWriter->writeAttribute('fontId', 0);
        $objWriter->writeAttribute('fillId', 0);
        $objWriter->writeAttribute('borderId', 0);
        $objWriter->endElement();

        $objWriter->endElement();

        // cellXfs
        $objWriter->startElement('cellXfs');
        $objWriter->writeAttribute('count', \count($phpExcel->getCellXfCollection()));

        // xf
        foreach ($phpExcel->getCellXfCollection() as $cellXf) {
            $this->writeCellStyleXf($objWriter, $cellXf, $phpExcel);
        }

        $objWriter->endElement();

        // cellStyles
        $objWriter->startElement('cellStyles');
        $objWriter->writeAttribute('count', 1);

        // cellStyle
        $objWriter->startElement('cellStyle');
        $objWriter->writeAttribute('name', 'Normal');
        $objWriter->writeAttribute('xfId', 0);
        $objWriter->writeAttribute('builtinId', 0);
        $objWriter->endElement();

        $objWriter->endElement();

        // dxfs
        $objWriter->startElement('dxfs');
        $objWriter->writeAttribute('count', $this->getParentWriter()->getStylesConditionalHashTable()->count());

        // dxf
        for ($i = 0; $i < $this->getParentWriter()->getStylesConditionalHashTable()->count(); ++$i) {
            $this->writeCellStyleDxf($objWriter, $this->getParentWriter()->getStylesConditionalHashTable()->getByIndex($i)->getStyle());
        }

        $objWriter->endElement();

        // tableStyles
        $objWriter->startElement('tableStyles');
        $objWriter->writeAttribute('defaultTableStyle', 'TableStyleMedium9');
        $objWriter->writeAttribute('defaultPivotStyle', 'PivotTableStyle1');
        $objWriter->endElement();

        $objWriter->endElement();

        // Return
        return $objWriter->getData();
    }

    /**
     * Write Fill
     *
     * @param \PhpOffice\PhpSpreadsheet\Shared\XMLWriter     $phpExcelSharedXMLWriter         XML Writer
     * @param \PhpOffice\PhpSpreadsheet\Style\Fill            $phpExcelStyleFill            Fill style
     * @throws     \PhpOffice\PhpSpreadsheet\Writer\Exception
     */
    private function writeFill(\PhpOffice\PhpSpreadsheet\Shared\XMLWriter $phpExcelSharedXMLWriter = \null, \PhpOffice\PhpSpreadsheet\Style\Fill $phpExcelStyleFill = \null)
    {
        // Check if this is a pattern type or gradient type
        if ($phpExcelStyleFill->getFillType() === \PhpOffice\PhpSpreadsheet\Style\Fill::FILL_GRADIENT_LINEAR ||
            $phpExcelStyleFill->getFillType() === \PhpOffice\PhpSpreadsheet\Style\Fill::FILL_GRADIENT_PATH) {
            // Gradient fill
            $this->writeGradientFill($phpExcelSharedXMLWriter, $phpExcelStyleFill);
        } elseif ($phpExcelStyleFill->getFillType() !== \null) {
            // Pattern fill
            $this->writePatternFill($phpExcelSharedXMLWriter, $phpExcelStyleFill);
        }
    }

    /**
     * Write Gradient Fill
     *
     * @param \PhpOffice\PhpSpreadsheet\Shared\XMLWriter     $phpExcelSharedXMLWriter         XML Writer
     * @param \PhpOffice\PhpSpreadsheet\Style\Fill            $phpExcelStyleFill            Fill style
     * @throws     \PhpOffice\PhpSpreadsheet\Writer\Exception
     */
    private function writeGradientFill(\PhpOffice\PhpSpreadsheet\Shared\XMLWriter $phpExcelSharedXMLWriter = \null, \PhpOffice\PhpSpreadsheet\Style\Fill $phpExcelStyleFill = \null)
    {
        // fill
        $phpExcelSharedXMLWriter->startElement('fill');

        // gradientFill
        $phpExcelSharedXMLWriter->startElement('gradientFill');
        $phpExcelSharedXMLWriter->writeAttribute('type', $phpExcelStyleFill->getFillType());
        $phpExcelSharedXMLWriter->writeAttribute('degree', $phpExcelStyleFill->getRotation());

        // stop
        $phpExcelSharedXMLWriter->startElement('stop');
        $phpExcelSharedXMLWriter->writeAttribute('position', '0');

        // color
        $phpExcelSharedXMLWriter->startElement('color');
        $phpExcelSharedXMLWriter->writeAttribute('rgb', $phpExcelStyleFill->getStartColor()->getARGB());
        $phpExcelSharedXMLWriter->endElement();

        $phpExcelSharedXMLWriter->endElement();

        // stop
        $phpExcelSharedXMLWriter->startElement('stop');
        $phpExcelSharedXMLWriter->writeAttribute('position', '1');

        // color
        $phpExcelSharedXMLWriter->startElement('color');
        $phpExcelSharedXMLWriter->writeAttribute('rgb', $phpExcelStyleFill->getEndColor()->getARGB());
        $phpExcelSharedXMLWriter->endElement();

        $phpExcelSharedXMLWriter->endElement();

        $phpExcelSharedXMLWriter->endElement();

        $phpExcelSharedXMLWriter->endElement();
    }

    /**
     * Write Pattern Fill
     *
     * @param \PhpOffice\PhpSpreadsheet\Shared\XMLWriter            $phpExcelSharedXMLWriter         XML Writer
     * @param \PhpOffice\PhpSpreadsheet\Style\Fill                    $phpExcelStyleFill            Fill style
     * @throws     \PhpOffice\PhpSpreadsheet\Writer\Exception
     */
    private function writePatternFill(\PhpOffice\PhpSpreadsheet\Shared\XMLWriter $phpExcelSharedXMLWriter = \null, \PhpOffice\PhpSpreadsheet\Style\Fill $phpExcelStyleFill = \null)
    {
        // fill
        $phpExcelSharedXMLWriter->startElement('fill');

        // patternFill
        $phpExcelSharedXMLWriter->startElement('patternFill');
        $phpExcelSharedXMLWriter->writeAttribute('patternType', $phpExcelStyleFill->getFillType());

        // fgColor
        if ($phpExcelStyleFill->getFillType() !== \PhpOffice\PhpSpreadsheet\Style\Fill::FILL_NONE && $phpExcelStyleFill->getStartColor()->getARGB()) {
            $phpExcelSharedXMLWriter->startElement('fgColor');
            $phpExcelSharedXMLWriter->writeAttribute('rgb', $phpExcelStyleFill->getStartColor()->getARGB());
            $phpExcelSharedXMLWriter->endElement();
        }
        // bgColor
        if ($phpExcelStyleFill->getFillType() !== \PhpOffice\PhpSpreadsheet\Style\Fill::FILL_NONE && $phpExcelStyleFill->getEndColor()->getARGB()) {
            $phpExcelSharedXMLWriter->startElement('bgColor');
            $phpExcelSharedXMLWriter->writeAttribute('rgb', $phpExcelStyleFill->getEndColor()->getARGB());
            $phpExcelSharedXMLWriter->endElement();
        }

        $phpExcelSharedXMLWriter->endElement();

        $phpExcelSharedXMLWriter->endElement();
    }

    /**
     * Write Font
     *
     * @param \PhpOffice\PhpSpreadsheet\Shared\XMLWriter        $phpExcelSharedXMLWriter         XML Writer
     * @param \PhpOffice\PhpSpreadsheet\Style\Font                $phpExcelStyleFont            Font style
     * @throws     \PhpOffice\PhpSpreadsheet\Writer\Exception
     */
    private function writeFont(\PhpOffice\PhpSpreadsheet\Shared\XMLWriter $phpExcelSharedXMLWriter = \null, \PhpOffice\PhpSpreadsheet\Style\Font $phpExcelStyleFont = \null)
    {
        // font
        $phpExcelSharedXMLWriter->startElement('font');
        //    Weird! The order of these elements actually makes a difference when opening Excel2007
        //        files in Excel2003 with the compatibility pack. It's not documented behaviour,
        //        and makes for a real WTF!

        // Bold. We explicitly write this element also when false (like MS Office Excel 2007 does
        // for conditional formatting). Otherwise it will apparently not be picked up in conditional
        // formatting style dialog
        if ($phpExcelStyleFont->getBold() !== \null) {
            $phpExcelSharedXMLWriter->startElement('b');
            $phpExcelSharedXMLWriter->writeAttribute('val', $phpExcelStyleFont->getBold() ? '1' : '0');
            $phpExcelSharedXMLWriter->endElement();
        }

        // Italic
        if ($phpExcelStyleFont->getItalic() !== \null) {
            $phpExcelSharedXMLWriter->startElement('i');
            $phpExcelSharedXMLWriter->writeAttribute('val', $phpExcelStyleFont->getItalic() ? '1' : '0');
            $phpExcelSharedXMLWriter->endElement();
        }

        // Strikethrough
        if ($phpExcelStyleFont->getStrikethrough() !== \null) {
            $phpExcelSharedXMLWriter->startElement('strike');
            $phpExcelSharedXMLWriter->writeAttribute('val', $phpExcelStyleFont->getStrikethrough() ? '1' : '0');
            $phpExcelSharedXMLWriter->endElement();
        }

        // Underline
        if ($phpExcelStyleFont->getUnderline() !== \null) {
            $phpExcelSharedXMLWriter->startElement('u');
            $phpExcelSharedXMLWriter->writeAttribute('val', $phpExcelStyleFont->getUnderline());
            $phpExcelSharedXMLWriter->endElement();
        }

        // Superscript / subscript
        if ($phpExcelStyleFont->getSuperScript() === \true || $phpExcelStyleFont->getSubScript() === \true) {
            $phpExcelSharedXMLWriter->startElement('vertAlign');
            if ($phpExcelStyleFont->getSuperScript() === \true) {
                $phpExcelSharedXMLWriter->writeAttribute('val', 'superscript');
            } elseif ($phpExcelStyleFont->getSubScript() === \true) {
                $phpExcelSharedXMLWriter->writeAttribute('val', 'subscript');
            }
            $phpExcelSharedXMLWriter->endElement();
        }

        // Size
        if ($phpExcelStyleFont->getSize() !== \null) {
            $phpExcelSharedXMLWriter->startElement('sz');
            $phpExcelSharedXMLWriter->writeAttribute('val', $phpExcelStyleFont->getSize());
            $phpExcelSharedXMLWriter->endElement();
        }

        // Foreground color
        if ($phpExcelStyleFont->getColor()->getARGB() !== \null) {
            $phpExcelSharedXMLWriter->startElement('color');
            $phpExcelSharedXMLWriter->writeAttribute('rgb', $phpExcelStyleFont->getColor()->getARGB());
            $phpExcelSharedXMLWriter->endElement();
        }

        // Name
        if ($phpExcelStyleFont->getName() !== \null) {
            $phpExcelSharedXMLWriter->startElement('name');
            $phpExcelSharedXMLWriter->writeAttribute('val', $phpExcelStyleFont->getName());
            $phpExcelSharedXMLWriter->endElement();
        }

        $phpExcelSharedXMLWriter->endElement();
    }

    /**
     * Write Border
     *
     * @param \PhpOffice\PhpSpreadsheet\Shared\XMLWriter            $phpExcelSharedXMLWriter         XML Writer
     * @param \PhpOffice\PhpSpreadsheet\Style\Borders                $phpExcelStyleBorders        Borders style
     * @throws     \PhpOffice\PhpSpreadsheet\Writer\Exception
     */
    private function writeBorder(\PhpOffice\PhpSpreadsheet\Shared\XMLWriter $phpExcelSharedXMLWriter = \null, \PhpOffice\PhpSpreadsheet\Style\Borders $phpExcelStyleBorders = \null)
    {
        // Write border
        $phpExcelSharedXMLWriter->startElement('border');
        // Diagonal?
        switch ($phpExcelStyleBorders->getDiagonalDirection()) {
            case \PhpOffice\PhpSpreadsheet\Style\Borders::DIAGONAL_UP:
                $phpExcelSharedXMLWriter->writeAttribute('diagonalUp', 'true');
                $phpExcelSharedXMLWriter->writeAttribute('diagonalDown', 'false');
                break;
            case \PhpOffice\PhpSpreadsheet\Style\Borders::DIAGONAL_DOWN:
                $phpExcelSharedXMLWriter->writeAttribute('diagonalUp', 'false');
                $phpExcelSharedXMLWriter->writeAttribute('diagonalDown', 'true');
                break;
            case \PhpOffice\PhpSpreadsheet\Style\Borders::DIAGONAL_BOTH:
                $phpExcelSharedXMLWriter->writeAttribute('diagonalUp', 'true');
                $phpExcelSharedXMLWriter->writeAttribute('diagonalDown', 'true');
                break;
        }

        // BorderPr
        $this->writeBorderPr($phpExcelSharedXMLWriter, 'left', $phpExcelStyleBorders->getLeft());
        $this->writeBorderPr($phpExcelSharedXMLWriter, 'right', $phpExcelStyleBorders->getRight());
        $this->writeBorderPr($phpExcelSharedXMLWriter, 'top', $phpExcelStyleBorders->getTop());
        $this->writeBorderPr($phpExcelSharedXMLWriter, 'bottom', $phpExcelStyleBorders->getBottom());
        $this->writeBorderPr($phpExcelSharedXMLWriter, 'diagonal', $phpExcelStyleBorders->getDiagonal());
        $phpExcelSharedXMLWriter->endElement();
    }

    /**
     * Write Cell Style Xf
     *
     * @param \PhpOffice\PhpSpreadsheet\Shared\XMLWriter            $phpExcelSharedXMLWriter         XML Writer
     * @param \PhpOffice\PhpSpreadsheet\Style\Style                        $phpExcelStyle            Style
     * @param \PhpOffice\PhpSpreadsheet\Spreadsheet                            $phpExcel        Workbook
     * @throws     \PhpOffice\PhpSpreadsheet\Writer\Exception
     */
    private function writeCellStyleXf(\PhpOffice\PhpSpreadsheet\Shared\XMLWriter $phpExcelSharedXMLWriter = \null, \PhpOffice\PhpSpreadsheet\Style\Style $phpExcelStyle = \null, \PhpOffice\PhpSpreadsheet\Spreadsheet $phpExcel = \null)
    {
        // xf
        $phpExcelSharedXMLWriter->startElement('xf');
        $phpExcelSharedXMLWriter->writeAttribute('xfId', 0);
        $phpExcelSharedXMLWriter->writeAttribute('fontId', (int)$this->getParentWriter()->getFontHashTable()->getIndexForHashCode($phpExcelStyle->getFont()->getHashCode()));
        if ($phpExcelStyle->getQuotePrefix()) {
            $phpExcelSharedXMLWriter->writeAttribute('quotePrefix', 1);
        }

        if ($phpExcelStyle->getNumberFormat()->getBuiltInFormatCode() === \false) {
            $phpExcelSharedXMLWriter->writeAttribute('numFmtId', (int)($this->getParentWriter()->getNumFmtHashTable()->getIndexForHashCode($phpExcelStyle->getNumberFormat()->getHashCode()) + 164));
        } else {
            $phpExcelSharedXMLWriter->writeAttribute('numFmtId', (int)$phpExcelStyle->getNumberFormat()->getBuiltInFormatCode());
        }

        $phpExcelSharedXMLWriter->writeAttribute('fillId', (int)$this->getParentWriter()->getFillHashTable()->getIndexForHashCode($phpExcelStyle->getFill()->getHashCode()));
        $phpExcelSharedXMLWriter->writeAttribute('borderId', (int)$this->getParentWriter()->getBordersHashTable()->getIndexForHashCode($phpExcelStyle->getBorders()->getHashCode()));

        // Apply styles?
        $phpExcelSharedXMLWriter->writeAttribute('applyFont', ($phpExcel->getDefaultStyle()->getFont()->getHashCode() != $phpExcelStyle->getFont()->getHashCode()) ? '1' : '0');
        $phpExcelSharedXMLWriter->writeAttribute('applyNumberFormat', ($phpExcel->getDefaultStyle()->getNumberFormat()->getHashCode() != $phpExcelStyle->getNumberFormat()->getHashCode()) ? '1' : '0');
        $phpExcelSharedXMLWriter->writeAttribute('applyFill', ($phpExcel->getDefaultStyle()->getFill()->getHashCode() != $phpExcelStyle->getFill()->getHashCode()) ? '1' : '0');
        $phpExcelSharedXMLWriter->writeAttribute('applyBorder', ($phpExcel->getDefaultStyle()->getBorders()->getHashCode() != $phpExcelStyle->getBorders()->getHashCode()) ? '1' : '0');
        $phpExcelSharedXMLWriter->writeAttribute('applyAlignment', ($phpExcel->getDefaultStyle()->getAlignment()->getHashCode() != $phpExcelStyle->getAlignment()->getHashCode()) ? '1' : '0');
        if ($phpExcelStyle->getProtection()->getLocked() != \PhpOffice\PhpSpreadsheet\Style\Protection::PROTECTION_INHERIT || $phpExcelStyle->getProtection()->getHidden() != \PhpOffice\PhpSpreadsheet\Style\Protection::PROTECTION_INHERIT) {
            $phpExcelSharedXMLWriter->writeAttribute('applyProtection', 'true');
        }

        // alignment
        $phpExcelSharedXMLWriter->startElement('alignment');
        $phpExcelSharedXMLWriter->writeAttribute('horizontal', $phpExcelStyle->getAlignment()->getHorizontal());
        $phpExcelSharedXMLWriter->writeAttribute('vertical', $phpExcelStyle->getAlignment()->getVertical());

        $textRotation = 0;
        if ($phpExcelStyle->getAlignment()->getTextRotation() >= 0) {
            $textRotation = $phpExcelStyle->getAlignment()->getTextRotation();
        } elseif ($phpExcelStyle->getAlignment()->getTextRotation() < 0) {
            $textRotation = 90 - $phpExcelStyle->getAlignment()->getTextRotation();
        }
        $phpExcelSharedXMLWriter->writeAttribute('textRotation', $textRotation);

        $phpExcelSharedXMLWriter->writeAttribute('wrapText', ($phpExcelStyle->getAlignment()->getWrapText() ? 'true' : 'false'));
        $phpExcelSharedXMLWriter->writeAttribute('shrinkToFit', ($phpExcelStyle->getAlignment()->getShrinkToFit() ? 'true' : 'false'));

        if ($phpExcelStyle->getAlignment()->getIndent() > 0) {
            $phpExcelSharedXMLWriter->writeAttribute('indent', $phpExcelStyle->getAlignment()->getIndent());
        }
        if ($phpExcelStyle->getAlignment()->getReadorder() > 0) {
            $phpExcelSharedXMLWriter->writeAttribute('readingOrder', $phpExcelStyle->getAlignment()->getReadorder());
        }
        $phpExcelSharedXMLWriter->endElement();

        // protection
        if ($phpExcelStyle->getProtection()->getLocked() != \PhpOffice\PhpSpreadsheet\Style\Protection::PROTECTION_INHERIT || $phpExcelStyle->getProtection()->getHidden() != \PhpOffice\PhpSpreadsheet\Style\Protection::PROTECTION_INHERIT) {
            $phpExcelSharedXMLWriter->startElement('protection');
            if ($phpExcelStyle->getProtection()->getLocked() != \PhpOffice\PhpSpreadsheet\Style\Protection::PROTECTION_INHERIT) {
                $phpExcelSharedXMLWriter->writeAttribute('locked', ($phpExcelStyle->getProtection()->getLocked() == \PhpOffice\PhpSpreadsheet\Style\Protection::PROTECTION_PROTECTED ? 'true' : 'false'));
            }
            if ($phpExcelStyle->getProtection()->getHidden() != \PhpOffice\PhpSpreadsheet\Style\Protection::PROTECTION_INHERIT) {
                $phpExcelSharedXMLWriter->writeAttribute('hidden', ($phpExcelStyle->getProtection()->getHidden() == \PhpOffice\PhpSpreadsheet\Style\Protection::PROTECTION_PROTECTED ? 'true' : 'false'));
            }
            $phpExcelSharedXMLWriter->endElement();
        }

        $phpExcelSharedXMLWriter->endElement();
    }

    /**
     * Write Cell Style Dxf
     *
     * @param \PhpOffice\PhpSpreadsheet\Shared\XMLWriter         $phpExcelSharedXMLWriter         XML Writer
     * @param \PhpOffice\PhpSpreadsheet\Style\Style                    $phpExcelStyle            Style
     * @throws     \PhpOffice\PhpSpreadsheet\Writer\Exception
     */
    private function writeCellStyleDxf(\PhpOffice\PhpSpreadsheet\Shared\XMLWriter $phpExcelSharedXMLWriter = \null, \PhpOffice\PhpSpreadsheet\Style\Style $phpExcelStyle = \null)
    {
        // dxf
        $phpExcelSharedXMLWriter->startElement('dxf');

        // font
        $this->writeFont($phpExcelSharedXMLWriter, $phpExcelStyle->getFont());

        // numFmt
        $this->writeNumFmt($phpExcelSharedXMLWriter, $phpExcelStyle->getNumberFormat());

        // fill
        $this->writeFill($phpExcelSharedXMLWriter, $phpExcelStyle->getFill());

        // alignment
        $phpExcelSharedXMLWriter->startElement('alignment');
        if ($phpExcelStyle->getAlignment()->getHorizontal() !== \null) {
            $phpExcelSharedXMLWriter->writeAttribute('horizontal', $phpExcelStyle->getAlignment()->getHorizontal());
        }
        if ($phpExcelStyle->getAlignment()->getVertical() !== \null) {
            $phpExcelSharedXMLWriter->writeAttribute('vertical', $phpExcelStyle->getAlignment()->getVertical());
        }

        if ($phpExcelStyle->getAlignment()->getTextRotation() !== \null) {
            $textRotation = 0;
            if ($phpExcelStyle->getAlignment()->getTextRotation() >= 0) {
                $textRotation = $phpExcelStyle->getAlignment()->getTextRotation();
            } elseif ($phpExcelStyle->getAlignment()->getTextRotation() < 0) {
                $textRotation = 90 - $phpExcelStyle->getAlignment()->getTextRotation();
            }
            $phpExcelSharedXMLWriter->writeAttribute('textRotation', $textRotation);
        }
        $phpExcelSharedXMLWriter->endElement();

        // border
        $this->writeBorder($phpExcelSharedXMLWriter, $phpExcelStyle->getBorders());

        // protection
        if ((($phpExcelStyle->getProtection()->getLocked() !== \null) || ($phpExcelStyle->getProtection()->getHidden() !== \null)) && ($phpExcelStyle->getProtection()->getLocked() !== \PhpOffice\PhpSpreadsheet\Style\Protection::PROTECTION_INHERIT ||
            $phpExcelStyle->getProtection()->getHidden() !== \PhpOffice\PhpSpreadsheet\Style\Protection::PROTECTION_INHERIT)) {
            $phpExcelSharedXMLWriter->startElement('protection');
            if (($phpExcelStyle->getProtection()->getLocked() !== \null) &&
                ($phpExcelStyle->getProtection()->getLocked() !== \PhpOffice\PhpSpreadsheet\Style\Protection::PROTECTION_INHERIT)) {
                $phpExcelSharedXMLWriter->writeAttribute('locked', ($phpExcelStyle->getProtection()->getLocked() == \PhpOffice\PhpSpreadsheet\Style\Protection::PROTECTION_PROTECTED ? 'true' : 'false'));
            }
            if (($phpExcelStyle->getProtection()->getHidden() !== \null) &&
                ($phpExcelStyle->getProtection()->getHidden() !== \PhpOffice\PhpSpreadsheet\Style\Protection::PROTECTION_INHERIT)) {
                $phpExcelSharedXMLWriter->writeAttribute('hidden', ($phpExcelStyle->getProtection()->getHidden() == \PhpOffice\PhpSpreadsheet\Style\Protection::PROTECTION_PROTECTED ? 'true' : 'false'));
            }
            $phpExcelSharedXMLWriter->endElement();
        }

        $phpExcelSharedXMLWriter->endElement();
    }

    /**
     * Write BorderPr
     *
     * @param \PhpOffice\PhpSpreadsheet\Shared\XMLWriter        $phpExcelSharedXMLWriter         XML Writer
     * @param     string                            $pName            Element name
     * @param \PhpOffice\PhpSpreadsheet\Style\Border            $phpExcelStyleBorder        Border style
     * @throws     \PhpOffice\PhpSpreadsheet\Writer\Exception
     */
    private function writeBorderPr(\PhpOffice\PhpSpreadsheet\Shared\XMLWriter $phpExcelSharedXMLWriter = \null, $pName = 'left', \PhpOffice\PhpSpreadsheet\Style\Border $phpExcelStyleBorder = \null)
    {
        // Write BorderPr
        if ($phpExcelStyleBorder->getBorderStyle() != \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_NONE) {
            $phpExcelSharedXMLWriter->startElement($pName);
            $phpExcelSharedXMLWriter->writeAttribute('style', $phpExcelStyleBorder->getBorderStyle());

            // color
            $phpExcelSharedXMLWriter->startElement('color');
            $phpExcelSharedXMLWriter->writeAttribute('rgb', $phpExcelStyleBorder->getColor()->getARGB());
            $phpExcelSharedXMLWriter->endElement();

            $phpExcelSharedXMLWriter->endElement();
        }
    }

    /**
     * Write NumberFormat
     *
     * @param \PhpOffice\PhpSpreadsheet\Shared\XMLWriter            $phpExcelSharedXMLWriter         XML Writer
     * @param \PhpOffice\PhpSpreadsheet\Style\NumberFormat            $phpExcelStyleNumberFormat    Number Format
     * @param     int                                    $pId            Number Format identifier
     * @throws     \PhpOffice\PhpSpreadsheet\Writer\Exception
     */
    private function writeNumFmt(\PhpOffice\PhpSpreadsheet\Shared\XMLWriter $phpExcelSharedXMLWriter = \null, \PhpOffice\PhpSpreadsheet\Style\NumberFormat $phpExcelStyleNumberFormat = \null, $pId = 0)
    {
        // Translate formatcode
        $formatCode = $phpExcelStyleNumberFormat->getFormatCode();

        // numFmt
        if ($formatCode !== \null) {
            $phpExcelSharedXMLWriter->startElement('numFmt');
            $phpExcelSharedXMLWriter->writeAttribute('numFmtId', ($pId + 164));
            $phpExcelSharedXMLWriter->writeAttribute('formatCode', $formatCode);
            $phpExcelSharedXMLWriter->endElement();
        }
    }

    /**
     * Get an array of all styles
     *
     * @return     \PhpOffice\PhpSpreadsheet\Style\Style[]        All styles in PHPExcel
     * @throws     \PhpOffice\PhpSpreadsheet\Writer\Exception
     */
    public function allStyles(\PhpOffice\PhpSpreadsheet\Spreadsheet $phpExcel = \null)
    {
        return $phpExcel->getCellXfCollection();
    }

    /**
     * Get an array of all conditional styles
     *
     * @return     \PhpOffice\PhpSpreadsheet\Style\Conditional[]        All conditional styles in PHPExcel
     * @throws     \PhpOffice\PhpSpreadsheet\Writer\Exception
     */
    public function allConditionalStyles(\PhpOffice\PhpSpreadsheet\Spreadsheet $phpExcel = \null)
    {
        // Get an array of all styles
        $aStyles = array();

        $sheetCount = $phpExcel->getSheetCount();
        for ($i = 0; $i < $sheetCount; ++$i) {
            foreach ($phpExcel->getSheet($i)->getConditionalStylesCollection() as $conditionalStyles) {
                foreach ($conditionalStyles as $conditionalStyle) {
                    $aStyles[] = $conditionalStyle;
                }
            }
        }

        return $aStyles;
    }

    /**
     * Get an array of all fills
     *
     * @return     \PhpOffice\PhpSpreadsheet\Style\Fill[]        All fills in PHPExcel
     * @throws     \PhpOffice\PhpSpreadsheet\Writer\Exception
     */
    public function allFills(\PhpOffice\PhpSpreadsheet\Spreadsheet $phpExcel = \null)
    {
        // Get an array of unique fills
        $aFills = array();

        // Two first fills are predefined
        $fill0 = new \PhpOffice\PhpSpreadsheet\Style\Fill();
        $fill0->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_NONE);
        $aFills[] = $fill0;

        $fill1 = new \PhpOffice\PhpSpreadsheet\Style\Fill();
        $fill1->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_PATTERN_GRAY125);
        $aFills[] = $fill1;
        // The remaining fills
        $aStyles = $this->allStyles($phpExcel);
        foreach ($aStyles as $aStyle) {
            if (!\array_key_exists($aStyle->getFill()->getHashCode(), $aFills)) {
                $aFills[ $aStyle->getFill()->getHashCode() ] = $aStyle->getFill();
            }
        }

        return $aFills;
    }

    /**
     * Get an array of all fonts
     *
     * @return     \PhpOffice\PhpSpreadsheet\Style\Font[]        All fonts in PHPExcel
     * @throws     \PhpOffice\PhpSpreadsheet\Writer\Exception
     */
    public function allFonts(\PhpOffice\PhpSpreadsheet\Spreadsheet $phpExcel = \null)
    {
        // Get an array of unique fonts
        $aFonts = array();
        $aStyles = $this->allStyles($phpExcel);

        foreach ($aStyles as $aStyle) {
            if (!\array_key_exists($aStyle->getFont()->getHashCode(), $aFonts)) {
                $aFonts[ $aStyle->getFont()->getHashCode() ] = $aStyle->getFont();
            }
        }

        return $aFonts;
    }

    /**
     * Get an array of all borders
     *
     * @return     \PhpOffice\PhpSpreadsheet\Style\Borders[]        All borders in PHPExcel
     * @throws     \PhpOffice\PhpSpreadsheet\Writer\Exception
     */
    public function allBorders(\PhpOffice\PhpSpreadsheet\Spreadsheet $phpExcel = \null)
    {
        // Get an array of unique borders
        $aBorders = array();
        $aStyles = $this->allStyles($phpExcel);

        foreach ($aStyles as $aStyle) {
            if (!\array_key_exists($aStyle->getBorders()->getHashCode(), $aBorders)) {
                $aBorders[ $aStyle->getBorders()->getHashCode() ] = $aStyle->getBorders();
            }
        }

        return $aBorders;
    }

    /**
     * Get an array of all number formats
     *
     * @return     \PhpOffice\PhpSpreadsheet\Style\NumberFormat[]        All number formats in PHPExcel
     * @throws     \PhpOffice\PhpSpreadsheet\Writer\Exception
     */
    public function allNumberFormats(\PhpOffice\PhpSpreadsheet\Spreadsheet $phpExcel = \null)
    {
        // Get an array of unique number formats
        $aNumFmts = array();
        $aStyles = $this->allStyles($phpExcel);

        foreach ($aStyles as $aStyle) {
            if ($aStyle->getNumberFormat()->getBuiltInFormatCode() === \false && !\array_key_exists($aStyle->getNumberFormat()->getHashCode(), $aNumFmts)) {
                $aNumFmts[ $aStyle->getNumberFormat()->getHashCode() ] = $aStyle->getNumberFormat();
            }
        }

        return $aNumFmts;
    }
}

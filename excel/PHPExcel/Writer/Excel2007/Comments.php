<?php

namespace PhpOffice\PhpSpreadsheet\Writer\Xlsx;

/**
 * PHPExcel_Writer_Excel2007_Comments
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
class Comments extends \PhpOffice\PhpSpreadsheet\Writer\Xlsx\WriterPart
{
    /**
     * Write comments to XML format
     *
     * @return     string                                 XML Output
     * @throws     \PhpOffice\PhpSpreadsheet\Writer\Exception
     */
    public function writeComments(\PhpOffice\PhpSpreadsheet\Worksheet\Worksheet $phpExcelWorksheet = \null)
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

          // Comments cache
          $comments    = $phpExcelWorksheet->getComments();

          // Authors cache
          $authors    = array();
          $authorId    = 0;
        foreach ($comments as $comment) {
            if (!isset($authors[$comment->getAuthor()])) {
                $authors[$comment->getAuthor()] = $authorId++;
            }
        }

        // comments
        $objWriter->startElement('comments');
        $objWriter->writeAttribute('xmlns', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main');

        // Loop through authors
        $objWriter->startElement('authors');
        foreach (array_keys($authors) as $author) {
            $objWriter->writeElement('author', $author);
        }
        $objWriter->endElement();

        // Loop through comments
        $objWriter->startElement('commentList');
        foreach ($comments as $key => $value) {
            $this->writeComment($objWriter, $key, $value, $authors);
        }
        $objWriter->endElement();

        $objWriter->endElement();

        // Return
        return $objWriter->getData();
    }

    /**
     * Write comment to XML format
     *
     * @param \PhpOffice\PhpSpreadsheet\Shared\XMLWriter        $phpExcelSharedXMLWriter             XML Writer
     * @param    string                            $pCellReference        Cell reference
     * @param \PhpOffice\PhpSpreadsheet\Comment                $phpExcelComment            Comment
     * @param    array                            $pAuthors            Array of authors
     * @throws     \PhpOffice\PhpSpreadsheet\Writer\Exception
     */
    private function writeComment(\PhpOffice\PhpSpreadsheet\Shared\XMLWriter $phpExcelSharedXMLWriter = \null, $pCellReference = 'A1', \PhpOffice\PhpSpreadsheet\Comment $phpExcelComment = \null, $pAuthors = \null)
    {
        // comment
        $phpExcelSharedXMLWriter->startElement('comment');
        $phpExcelSharedXMLWriter->writeAttribute('ref', $pCellReference);
        $phpExcelSharedXMLWriter->writeAttribute('authorId', $pAuthors[$phpExcelComment->getAuthor()]);

        // text
        $phpExcelSharedXMLWriter->startElement('text');
        $this->getParentWriter()->getWriterPart('stringtable')->writeRichText($phpExcelSharedXMLWriter, $phpExcelComment->getText());
        $phpExcelSharedXMLWriter->endElement();

        $phpExcelSharedXMLWriter->endElement();
    }

    /**
     * Write VML comments to XML format
     *
     * @return     string                                 XML Output
     * @throws     \PhpOffice\PhpSpreadsheet\Writer\Exception
     */
    public function writeVMLComments(\PhpOffice\PhpSpreadsheet\Worksheet\Worksheet $phpExcelWorksheet = \null)
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

          // Comments cache
          $comments    = $phpExcelWorksheet->getComments();

        // xml
        $objWriter->startElement('xml');
        $objWriter->writeAttribute('xmlns:v', 'urn:schemas-microsoft-com:vml');
        $objWriter->writeAttribute('xmlns:o', 'urn:schemas-microsoft-com:office:office');
        $objWriter->writeAttribute('xmlns:x', 'urn:schemas-microsoft-com:office:excel');

        // o:shapelayout
        $objWriter->startElement('o:shapelayout');
        $objWriter->writeAttribute('v:ext', 'edit');

            // o:idmap
            $objWriter->startElement('o:idmap');
            $objWriter->writeAttribute('v:ext', 'edit');
            $objWriter->writeAttribute('data', '1');
            $objWriter->endElement();

        $objWriter->endElement();

        // v:shapetype
        $objWriter->startElement('v:shapetype');
        $objWriter->writeAttribute('id', '_x0000_t202');
        $objWriter->writeAttribute('coordsize', '21600,21600');
        $objWriter->writeAttribute('o:spt', '202');
        $objWriter->writeAttribute('path', 'm,l,21600r21600,l21600,xe');

            // v:stroke
            $objWriter->startElement('v:stroke');
            $objWriter->writeAttribute('joinstyle', 'miter');
            $objWriter->endElement();

            // v:path
            $objWriter->startElement('v:path');
            $objWriter->writeAttribute('gradientshapeok', 't');
            $objWriter->writeAttribute('o:connecttype', 'rect');
            $objWriter->endElement();

        $objWriter->endElement();

        // Loop through comments
        foreach ($comments as $key => $value) {
            $this->writeVMLComment($objWriter, $key, $value);
        }

        $objWriter->endElement();

        // Return
        return $objWriter->getData();
    }

    /**
     * Write VML comment to XML format
     *
     * @param \PhpOffice\PhpSpreadsheet\Shared\XMLWriter        $phpExcelSharedXMLWriter             XML Writer
     * @param    string                            $pCellReference        Cell reference
     * @param \PhpOffice\PhpSpreadsheet\Comment                $phpExcelComment            Comment
     * @throws     \PhpOffice\PhpSpreadsheet\Writer\Exception
     */
    private function writeVMLComment(\PhpOffice\PhpSpreadsheet\Shared\XMLWriter $phpExcelSharedXMLWriter = \null, $pCellReference = 'A1', \PhpOffice\PhpSpreadsheet\Comment $phpExcelComment = \null)
    {
         // Metadata
         list($column, $row) = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::coordinateFromString($pCellReference);
         $column = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::columnIndexFromString($column);
         $id = 1024 + $column + $row;
         $id = \substr($id, 0, 4);

        // v:shape
        $phpExcelSharedXMLWriter->startElement('v:shape');
        $phpExcelSharedXMLWriter->writeAttribute('id', '_x0000_s' . $id);
        $phpExcelSharedXMLWriter->writeAttribute('type', '#_x0000_t202');
        $phpExcelSharedXMLWriter->writeAttribute('style', 'position:absolute;margin-left:' . $phpExcelComment->getMarginLeft() . ';margin-top:' . $phpExcelComment->getMarginTop() . ';width:' . $phpExcelComment->getWidth() . ';height:' . $phpExcelComment->getHeight() . ';z-index:1;visibility:' . ($phpExcelComment->isVisible() ? 'visible' : 'hidden'));
        $phpExcelSharedXMLWriter->writeAttribute('fillcolor', '#' . $phpExcelComment->getFillColor()->getRGB());
        $phpExcelSharedXMLWriter->writeAttribute('o:insetmode', 'auto');

            // v:fill
            $phpExcelSharedXMLWriter->startElement('v:fill');
            $phpExcelSharedXMLWriter->writeAttribute('color2', '#' . $phpExcelComment->getFillColor()->getRGB());
            $phpExcelSharedXMLWriter->endElement();

            // v:shadow
            $phpExcelSharedXMLWriter->startElement('v:shadow');
            $phpExcelSharedXMLWriter->writeAttribute('on', 't');
            $phpExcelSharedXMLWriter->writeAttribute('color', 'black');
            $phpExcelSharedXMLWriter->writeAttribute('obscured', 't');
            $phpExcelSharedXMLWriter->endElement();

            // v:path
            $phpExcelSharedXMLWriter->startElement('v:path');
            $phpExcelSharedXMLWriter->writeAttribute('o:connecttype', 'none');
            $phpExcelSharedXMLWriter->endElement();

            // v:textbox
            $phpExcelSharedXMLWriter->startElement('v:textbox');
            $phpExcelSharedXMLWriter->writeAttribute('style', 'mso-direction-alt:auto');

                // div
                $phpExcelSharedXMLWriter->startElement('div');
                $phpExcelSharedXMLWriter->writeAttribute('style', 'text-align:left');
                $phpExcelSharedXMLWriter->endElement();

            $phpExcelSharedXMLWriter->endElement();

            // x:ClientData
            $phpExcelSharedXMLWriter->startElement('x:ClientData');
            $phpExcelSharedXMLWriter->writeAttribute('ObjectType', 'Note');

                // x:MoveWithCells
                $phpExcelSharedXMLWriter->writeElement('x:MoveWithCells', '');

                // x:SizeWithCells
                $phpExcelSharedXMLWriter->writeElement('x:SizeWithCells', '');

                // x:Anchor
                //$objWriter->writeElement('x:Anchor', $column . ', 15, ' . ($row - 2) . ', 10, ' . ($column + 4) . ', 15, ' . ($row + 5) . ', 18');

                // x:AutoFill
                $phpExcelSharedXMLWriter->writeElement('x:AutoFill', 'False');

                // x:Row
                $phpExcelSharedXMLWriter->writeElement('x:Row', ($row - 1));

                // x:Column
                $phpExcelSharedXMLWriter->writeElement('x:Column', ($column - 1));

            $phpExcelSharedXMLWriter->endElement();

        $phpExcelSharedXMLWriter->endElement();
    }
}

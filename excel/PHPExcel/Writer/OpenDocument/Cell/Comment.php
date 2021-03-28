<?php
namespace PhpOffice\PhpSpreadsheet\Writer\Ods\Cell;

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
 * PHPExcel_Writer_OpenDocument_Cell_Comment
 *
 * @category   PHPExcel
 * @package    PHPExcel_Writer_OpenDocument
 * @copyright  Copyright (c) 2006 - 2015 PHPExcel (http://www.codeplex.com/PHPExcel)
 * @author     Alexander Pervakov <frost-nzcr4@jagmort.com>
 */
class Comment
{
    public static function write(\PhpOffice\PhpSpreadsheet\Shared\XMLWriter $phpExcelSharedXMLWriter, \PhpOffice\PhpSpreadsheet\Cell\Cell $phpExcelCell)
    {
        $comments = $phpExcelCell->getWorksheet()->getComments();
        if (!isset($comments[$phpExcelCell->getCoordinate()])) {
            return;
        }
        $comment = $comments[$phpExcelCell->getCoordinate()];

        $phpExcelSharedXMLWriter->startElement('office:annotation');
            //$objWriter->writeAttribute('draw:style-name', 'gr1');
            //$objWriter->writeAttribute('draw:text-style-name', 'P1');
            $phpExcelSharedXMLWriter->writeAttribute('svg:width', $comment->getWidth());
            $phpExcelSharedXMLWriter->writeAttribute('svg:height', $comment->getHeight());
            $phpExcelSharedXMLWriter->writeAttribute('svg:x', $comment->getMarginLeft());
            $phpExcelSharedXMLWriter->writeAttribute('svg:y', $comment->getMarginTop());
            //$objWriter->writeAttribute('draw:caption-point-x', $comment->getMarginLeft());
            //$objWriter->writeAttribute('draw:caption-point-y', $comment->getMarginTop());
                $phpExcelSharedXMLWriter->writeElement('dc:creator', $comment->getAuthor());
                // TODO: Not realized in PHPExcel_Comment yet.
                //$objWriter->writeElement('dc:date', $comment->getDate());
                $phpExcelSharedXMLWriter->writeElement('text:p', $comment->getText()->getPlainText());
                    //$objWriter->writeAttribute('draw:text-style-name', 'P1');
        $phpExcelSharedXMLWriter->endElement();
    }
}

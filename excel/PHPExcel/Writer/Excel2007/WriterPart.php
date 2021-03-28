<?php

namespace PhpOffice\PhpSpreadsheet\Writer\Xlsx;

/**
 * PHPExcel_Writer_Excel2007_WriterPart
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
abstract class WriterPart
{
    /**
     * Parent IWriter object
     *
     * @var \PhpOffice\PhpSpreadsheet\Writer\IWriter
     */
    private $phpExcelWriterIWriter;

    /**
     * Set parent IWriter object
     *
     * @throws \PhpOffice\PhpSpreadsheet\Writer\Exception
     */
    public function setParentWriter(\PhpOffice\PhpSpreadsheet\Writer\IWriter $phpExcelWriterIWriter = \null)
    {
        $this->phpExcelWriterIWriter = $phpExcelWriterIWriter;
    }

    /**
     * Get parent IWriter object
     *
     * @return \PhpOffice\PhpSpreadsheet\Writer\IWriter
     * @throws \PhpOffice\PhpSpreadsheet\Writer\Exception
     */
    public function getParentWriter()
    {
        if (!\is_null($this->phpExcelWriterIWriter)) {
            return $this->phpExcelWriterIWriter;
        } else {
            throw new \PhpOffice\PhpSpreadsheet\Writer\Exception("No parent PHPExcel_Writer_IWriter assigned.");
        }
    }

    /**
     * Set parent IWriter object
     *
     * @throws \PhpOffice\PhpSpreadsheet\Writer\Exception
     */
    public function __construct(\PhpOffice\PhpSpreadsheet\Writer\IWriter $phpExcelWriterIWriter = \null)
    {
        if (!\is_null($phpExcelWriterIWriter)) {
            $this->phpExcelWriterIWriter = $phpExcelWriterIWriter;
        }
    }
}

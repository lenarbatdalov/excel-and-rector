<?php

namespace PhpOffice\PhpSpreadsheet\Worksheet;

/**
 * PHPExcel_Worksheet_Column
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
 * @package    PHPExcel_Worksheet
 * @copyright  Copyright (c) 2006 - 2015 PHPExcel (http://www.codeplex.com/PHPExcel)
 * @license    http://www.gnu.org/licenses/old-licenses/lgpl-2.1.txt    LGPL
 * @version    ##VERSION##, ##DATE##
 */
class Column
{
    /**
     * PHPExcel_Worksheet
     *
     * @var \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet
     */
    private $phpExcelWorksheet;

    /**
     * Column index
     *
     * @var string
     */
    private $columnIndex;

    /**
     * Create a new column
     *
     * @param string                $columnIndex
     */
    public function __construct(\PhpOffice\PhpSpreadsheet\Worksheet\Worksheet $phpExcelWorksheet = \null, $columnIndex = 'A')
    {
        // Set parent and column index
        $this->phpExcelWorksheet         = $phpExcelWorksheet;
        $this->columnIndex = $columnIndex;
    }

    /**
     * Destructor
     */
    public function __destruct()
    {
        unset($this->phpExcelWorksheet);
    }

    /**
     * Get column index
     *
     * @return string
     */
    public function getColumnIndex()
    {
        return $this->columnIndex;
    }

    /**
     * Get cell iterator
     *
     * @param    integer                $startRow        The row number at which to start iterating
     * @param    integer                $endRow            Optionally, the row number at which to stop iterating
     * @return \PhpOffice\PhpSpreadsheet\Worksheet\CellIterator
     */
    public function getCellIterator($startRow = 1, $endRow = \null)
    {
        return new \PhpOffice\PhpSpreadsheet\Worksheet\ColumnCellIterator($this->phpExcelWorksheet, $this->columnIndex, $startRow, $endRow);
    }
}

<?php

namespace PhpOffice\PhpSpreadsheet\Worksheet;

/**
 * PHPExcel_WorksheetIterator
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
 * @package    PHPExcel
 * @copyright  Copyright (c) 2006 - 2015 PHPExcel (http://www.codeplex.com/PHPExcel)
 * @license    http://www.gnu.org/licenses/old-licenses/lgpl-2.1.txt    LGPL
 * @version    ##VERSION##, ##DATE##
 */
class Iterator implements \Iterator
{
    /**
     * Spreadsheet to iterate
     *
     * @var \PhpOffice\PhpSpreadsheet\Spreadsheet
     */
    private $phpExcel;

    /**
     * Current iterator position
     *
     * @var int
     */
    private $position = 0;

    /**
     * Create a new worksheet iterator
     */
    public function __construct(\PhpOffice\PhpSpreadsheet\Spreadsheet $phpExcel = \null)
    {
        // Set subject
        $this->phpExcel = $phpExcel;
    }

    /**
     * Destructor
     */
    public function __destruct()
    {
        unset($this->phpExcel);
    }

    /**
     * Rewind iterator
     */
    public function rewind()
    {
        $this->position = 0;
    }

    /**
     * Current PHPExcel_Worksheet
     *
     * @return \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet
     */
    public function current()
    {
        return $this->phpExcel->getSheet($this->position);
    }

    /**
     * Current key
     *
     * @return int
     */
    public function getPosition()
    {
        return $this->position;
    }

    /**
     * Next value
     */
    public function next()
    {
        ++$this->position;
    }

    /**
     * More PHPExcel_Worksheet instances available?
     *
     * @return boolean
     */
    public function valid()
    {
        return $this->position < $this->phpExcel->getSheetCount();
    }
}

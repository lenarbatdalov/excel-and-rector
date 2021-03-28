<?php

namespace PhpOffice\PhpSpreadsheet;

/**
 * PHPExcel_HashTable
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
class HashTable
{
    /**
     * HashTable elements
     *
     * @var array
     */
    protected $items = array();

    /**
     * HashTable key map
     *
     * @var array
     */
    protected $keyMap = array();

    /**
     * Create a new PHPExcel_HashTable
     *
     * @param    \PhpOffice\PhpSpreadsheet\IComparable[] $pSource    Optional source array to create HashTable from
     * @throws    \PhpOffice\PhpSpreadsheet\Exception
     */
    public function __construct($pSource = \null)
    {
        if ($pSource !== \null) {
            // Create HashTable
            $this->addFromSource($pSource);
        }
    }

    /**
     * Add HashTable items from source
     *
     * @param    \PhpOffice\PhpSpreadsheet\IComparable[] $pSource    Source array to create HashTable from
     * @throws    \PhpOffice\PhpSpreadsheet\Exception
     */
    public function addFromSource($pSource = \null)
    {
        // Check if an array was passed
        if ($pSource == \null) {
            return;
        } elseif (!\is_array($pSource)) {
            throw new \PhpOffice\PhpSpreadsheet\Exception('Invalid array parameter passed.');
        }

        foreach ($pSource as $singlePSource) {
            $this->add($singlePSource);
        }
    }

    /**
     * Add HashTable item
     *
     * @param \PhpOffice\PhpSpreadsheet\IComparable $phpExcelIComparable    Item to add
     * @throws    \PhpOffice\PhpSpreadsheet\Exception
     */
    public function add(\PhpOffice\PhpSpreadsheet\IComparable $phpExcelIComparable = \null)
    {
        $hash = $phpExcelIComparable->getHashCode();
        if (!isset($this->items[$hash])) {
            $this->items[$hash] = $phpExcelIComparable;
            $this->keyMap[\count($this->items) - 1] = $hash;
        }
    }

    /**
     * Remove HashTable item
     *
     * @param \PhpOffice\PhpSpreadsheet\IComparable $phpExcelIComparable    Item to remove
     * @throws    \PhpOffice\PhpSpreadsheet\Exception
     */
    public function remove(\PhpOffice\PhpSpreadsheet\IComparable $phpExcelIComparable = \null)
    {
        $hash = $phpExcelIComparable->getHashCode();
        if (isset($this->items[$hash])) {
            unset($this->items[$hash]);

            $deleteKey = -1;
            foreach ($this->keyMap as $key => $value) {
                if ($deleteKey >= 0) {
                    $this->keyMap[$key - 1] = $value;
                }

                if ($value == $hash) {
                    $deleteKey = $key;
                }
            }
            unset($this->keyMap[\count($this->keyMap) - 1]);
        }
    }

    /**
     * Clear HashTable
     *
     */
    public function clear()
    {
        $this->items = array();
        $this->keyMap = array();
    }

    /**
     * Count
     *
     * @return int
     */
    public function count()
    {
        return \count($this->items);
    }

    /**
     * Get index for hash code
     *
     * @param    string    $pHashCode
     * @return    int    Index
     */
    public function getIndexForHashCode($pHashCode = '')
    {
        return \array_search($pHashCode, $this->keyMap);
    }

    /**
     * Get by index
     *
     * @param    int    $pIndex
     * @return    \PhpOffice\PhpSpreadsheet\IComparable
     */
    public function getByIndex($pIndex = 0)
    {
        if (isset($this->keyMap[$pIndex])) {
            return $this->getByHashCode($this->keyMap[$pIndex]);
        }

        return \null;
    }

    /**
     * Get by hashcode
     *
     * @param    string    $pHashCode
     * @return    \PhpOffice\PhpSpreadsheet\IComparable
     */
    public function getByHashCode($pHashCode = '')
    {
        if (isset($this->items[$pHashCode])) {
            return $this->items[$pHashCode];
        }

        return \null;
    }

    /**
     * HashTable to array
     *
     * @return \PhpOffice\PhpSpreadsheet\IComparable[]
     */
    public function getItems()
    {
        return $this->items;
    }

    /**
     * Implement PHP __clone to create a deep clone, not just a shallow copy.
     */
    public function __clone()
    {
        $vars = \get_object_vars($this);
        foreach ($vars as $key => $value) {
            if (\is_object($value)) {
                $this->$key = clone $value;
            }
        }
    }
}

<?php
namespace PhpOffice\PhpSpreadsheet\Worksheet;

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
 * @package    PHPExcel_Worksheet
 * @copyright  Copyright (c) 2006 - 2015 PHPExcel (http://www.codeplex.com/PHPExcel)
 * @license    http://www.gnu.org/licenses/old-licenses/lgpl-2.1.txt    LGPL
 * @version    ##VERSION##, ##DATE##
 */
/**
 * PHPExcel_Worksheet_Protection
 *
 * @category   PHPExcel
 * @package    PHPExcel_Worksheet
 * @copyright  Copyright (c) 2006 - 2015 PHPExcel (http://www.codeplex.com/PHPExcel)
 */
class Protection
{
    /**
     * Sheet
     *
     * @var boolean
     */
    private $sheet                    = \false;

    /**
     * Objects
     *
     * @var boolean
     */
    private $objects                = \false;

    /**
     * Scenarios
     *
     * @var boolean
     */
    private $scenarios                = \false;

    /**
     * Format cells
     *
     * @var boolean
     */
    private $formatCells            = \false;

    /**
     * Format columns
     *
     * @var boolean
     */
    private $formatColumns            = \false;

    /**
     * Format rows
     *
     * @var boolean
     */
    private $formatRows            = \false;

    /**
     * Insert columns
     *
     * @var boolean
     */
    private $insertColumns            = \false;

    /**
     * Insert rows
     *
     * @var boolean
     */
    private $insertRows            = \false;

    /**
     * Insert hyperlinks
     *
     * @var boolean
     */
    private $insertHyperlinks        = \false;

    /**
     * Delete columns
     *
     * @var boolean
     */
    private $deleteColumns            = \false;

    /**
     * Delete rows
     *
     * @var boolean
     */
    private $deleteRows            = \false;

    /**
     * Select locked cells
     *
     * @var boolean
     */
    private $selectLockedCells        = \false;

    /**
     * Sort
     *
     * @var boolean
     */
    private $sort                    = \false;

    /**
     * AutoFilter
     *
     * @var boolean
     */
    private $autoFilter            = \false;

    /**
     * Pivot tables
     *
     * @var boolean
     */
    private $pivotTables            = \false;

    /**
     * Select unlocked cells
     *
     * @var boolean
     */
    private $selectUnlockedCells    = \false;

    /**
     * Password
     *
     * @var string
     */
    private $password                = '';

    /**
     * Is some sort of protection enabled?
     *
     * @return boolean
     */
    public function isProtectionEnabled()
    {
        return $this->sheet ||
            $this->objects ||
            $this->scenarios ||
            $this->formatCells ||
            $this->formatColumns ||
            $this->formatRows ||
            $this->insertColumns ||
            $this->insertRows ||
            $this->insertHyperlinks ||
            $this->deleteColumns ||
            $this->deleteRows ||
            $this->selectLockedCells ||
            $this->sort ||
            $this->autoFilter ||
            $this->pivotTables ||
            $this->selectUnlockedCells;
    }

    /**
     * Get Sheet
     *
     * @return boolean
     */
    public function isSheet()
    {
        return $this->sheet;
    }

    /**
     * Set Sheet
     *
     * @param boolean $pValue
     * @return \PhpOffice\PhpSpreadsheet\Worksheet\Protection
     */
    public function setSheet($pValue = \false)
    {
        $this->sheet = $pValue;
        return $this;
    }

    /**
     * Get Objects
     *
     * @return boolean
     */
    public function isObjects()
    {
        return $this->objects;
    }

    /**
     * Set Objects
     *
     * @param boolean $pValue
     * @return \PhpOffice\PhpSpreadsheet\Worksheet\Protection
     */
    public function setObjects($pValue = \false)
    {
        $this->objects = $pValue;
        return $this;
    }

    /**
     * Get Scenarios
     *
     * @return boolean
     */
    public function isScenarios()
    {
        return $this->scenarios;
    }

    /**
     * Set Scenarios
     *
     * @param boolean $pValue
     * @return \PhpOffice\PhpSpreadsheet\Worksheet\Protection
     */
    public function setScenarios($pValue = \false)
    {
        $this->scenarios = $pValue;
        return $this;
    }

    /**
     * Get FormatCells
     *
     * @return boolean
     */
    public function isFormatCells()
    {
        return $this->formatCells;
    }

    /**
     * Set FormatCells
     *
     * @param boolean $pValue
     * @return \PhpOffice\PhpSpreadsheet\Worksheet\Protection
     */
    public function setFormatCells($pValue = \false)
    {
        $this->formatCells = $pValue;
        return $this;
    }

    /**
     * Get FormatColumns
     *
     * @return boolean
     */
    public function isFormatColumns()
    {
        return $this->formatColumns;
    }

    /**
     * Set FormatColumns
     *
     * @param boolean $pValue
     * @return \PhpOffice\PhpSpreadsheet\Worksheet\Protection
     */
    public function setFormatColumns($pValue = \false)
    {
        $this->formatColumns = $pValue;
        return $this;
    }

    /**
     * Get FormatRows
     *
     * @return boolean
     */
    public function isFormatRows()
    {
        return $this->formatRows;
    }

    /**
     * Set FormatRows
     *
     * @param boolean $pValue
     * @return \PhpOffice\PhpSpreadsheet\Worksheet\Protection
     */
    public function setFormatRows($pValue = \false)
    {
        $this->formatRows = $pValue;
        return $this;
    }

    /**
     * Get InsertColumns
     *
     * @return boolean
     */
    public function isInsertColumns()
    {
        return $this->insertColumns;
    }

    /**
     * Set InsertColumns
     *
     * @param boolean $pValue
     * @return \PhpOffice\PhpSpreadsheet\Worksheet\Protection
     */
    public function setInsertColumns($pValue = \false)
    {
        $this->insertColumns = $pValue;
        return $this;
    }

    /**
     * Get InsertRows
     *
     * @return boolean
     */
    public function isInsertRows()
    {
        return $this->insertRows;
    }

    /**
     * Set InsertRows
     *
     * @param boolean $pValue
     * @return \PhpOffice\PhpSpreadsheet\Worksheet\Protection
     */
    public function setInsertRows($pValue = \false)
    {
        $this->insertRows = $pValue;
        return $this;
    }

    /**
     * Get InsertHyperlinks
     *
     * @return boolean
     */
    public function isInsertHyperlinks()
    {
        return $this->insertHyperlinks;
    }

    /**
     * Set InsertHyperlinks
     *
     * @param boolean $pValue
     * @return \PhpOffice\PhpSpreadsheet\Worksheet\Protection
     */
    public function setInsertHyperlinks($pValue = \false)
    {
        $this->insertHyperlinks = $pValue;
        return $this;
    }

    /**
     * Get DeleteColumns
     *
     * @return boolean
     */
    public function isDeleteColumns()
    {
        return $this->deleteColumns;
    }

    /**
     * Set DeleteColumns
     *
     * @param boolean $pValue
     * @return \PhpOffice\PhpSpreadsheet\Worksheet\Protection
     */
    public function setDeleteColumns($pValue = \false)
    {
        $this->deleteColumns = $pValue;
        return $this;
    }

    /**
     * Get DeleteRows
     *
     * @return boolean
     */
    public function isDeleteRows()
    {
        return $this->deleteRows;
    }

    /**
     * Set DeleteRows
     *
     * @param boolean $pValue
     * @return \PhpOffice\PhpSpreadsheet\Worksheet\Protection
     */
    public function setDeleteRows($pValue = \false)
    {
        $this->deleteRows = $pValue;
        return $this;
    }

    /**
     * Get SelectLockedCells
     *
     * @return boolean
     */
    public function isSelectLockedCells()
    {
        return $this->selectLockedCells;
    }

    /**
     * Set SelectLockedCells
     *
     * @param boolean $pValue
     * @return \PhpOffice\PhpSpreadsheet\Worksheet\Protection
     */
    public function setSelectLockedCells($pValue = \false)
    {
        $this->selectLockedCells = $pValue;
        return $this;
    }

    /**
     * Get Sort
     *
     * @return boolean
     */
    public function isSort()
    {
        return $this->sort;
    }

    /**
     * Set Sort
     *
     * @param boolean $pValue
     * @return \PhpOffice\PhpSpreadsheet\Worksheet\Protection
     */
    public function setSort($pValue = \false)
    {
        $this->sort = $pValue;
        return $this;
    }

    /**
     * Get AutoFilter
     *
     * @return boolean
     */
    public function isAutoFilter()
    {
        return $this->autoFilter;
    }

    /**
     * Set AutoFilter
     *
     * @param boolean $pValue
     * @return \PhpOffice\PhpSpreadsheet\Worksheet\Protection
     */
    public function setAutoFilter($pValue = \false)
    {
        $this->autoFilter = $pValue;
        return $this;
    }

    /**
     * Get PivotTables
     *
     * @return boolean
     */
    public function isPivotTables()
    {
        return $this->pivotTables;
    }

    /**
     * Set PivotTables
     *
     * @param boolean $pValue
     * @return \PhpOffice\PhpSpreadsheet\Worksheet\Protection
     */
    public function setPivotTables($pValue = \false)
    {
        $this->pivotTables = $pValue;
        return $this;
    }

    /**
     * Get SelectUnlockedCells
     *
     * @return boolean
     */
    public function isSelectUnlockedCells()
    {
        return $this->selectUnlockedCells;
    }

    /**
     * Set SelectUnlockedCells
     *
     * @param boolean $pValue
     * @return \PhpOffice\PhpSpreadsheet\Worksheet\Protection
     */
    public function setSelectUnlockedCells($pValue = \false)
    {
        $this->selectUnlockedCells = $pValue;
        return $this;
    }

    /**
     * Get Password (hashed)
     *
     * @return string
     */
    public function getPassword()
    {
        return $this->password;
    }

    /**
     * Set Password
     *
     * @param string     $pValue
     * @param boolean     $pAlreadyHashed If the password has already been hashed, set this to true
     * @return \PhpOffice\PhpSpreadsheet\Worksheet\Protection
     */
    public function setPassword($pValue = '', $pAlreadyHashed = \false)
    {
        if (!$pAlreadyHashed) {
            $pValue = \PhpOffice\PhpSpreadsheet\Shared\PasswordHasher::hashPassword($pValue);
        }
        $this->password = $pValue;
        return $this;
    }

    /**
     * Implement PHP __clone to create a deep clone, not just a shallow copy.
     */
    public function __clone()
    {
        $vars = \get_object_vars($this);
        foreach ($vars as $key => $value) {
            $this->$key = \is_object($value) ? clone $value : $value;
        }
    }
}

<?php

namespace PhpOffice\PhpSpreadsheet;

/**
 * PHPExcel_NamedRange
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
class NamedRange
{
    /**
     * Range name
     *
     * @var string
     */
    private $name;

    /**
     * Worksheet on which the named range can be resolved
     *
     * @var \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet
     */
    private $worksheet;

    /**
     * Range of the referenced cells
     *
     * @var string
     */
    private $range;

    /**
     * Is the named range local? (i.e. can only be used on $this->worksheet)
     *
     * @var bool
     */
    private $localOnly;

    /**
     * Scope
     *
     * @var \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet
     */
    private $scope;

    /**
     * Create a new NamedRange
     *
     * @param string $pName
     * @param string $pRange
     * @param bool $pLocalOnly
     * @param \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet|null $pScope    Scope. Only applies when $pLocalOnly = true. Null for global scope.
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     */
    public function __construct($pName = \null, \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet $phpExcelWorksheet, $pRange = 'A1', $pLocalOnly = \false, $pScope = \null)
    {
        // Validate data
        if (($pName === \null) || ($phpExcelWorksheet === \null) || ($pRange === \null)) {
            throw new \PhpOffice\PhpSpreadsheet\Exception('Parameters can not be null.');
        }

        // Set local members
        $this->name       = $pName;
        $this->worksheet  = $phpExcelWorksheet;
        $this->range      = $pRange;
        $this->localOnly  = $pLocalOnly;
        $this->scope      = ($pLocalOnly == \true) ? (($pScope == \null) ? $phpExcelWorksheet : $pScope) : \null;
    }

    /**
     * Get name
     *
     * @return string
     */
    public function getName()
    {
        return $this->name;
    }

    /**
     * Set name
     *
     * @param string $value
     * @return \PhpOffice\PhpSpreadsheet\NamedRange
     */
    public function setName($value = \null)
    {
        if ($value !== \null) {
            // Old title
            $oldTitle = $this->name;

            // Re-attach
            if ($this->worksheet !== \null) {
                $this->worksheet->getParent()->removeNamedRange($this->name, $this->worksheet);
            }
            $this->name = $value;

            if ($this->worksheet !== \null) {
                $this->worksheet->getParent()->addNamedRange($this);
            }

            // New title
            $newTitle = $this->name;
            \PhpOffice\PhpSpreadsheet\ReferenceHelper::getInstance()->updateNamedFormulas($this->worksheet->getParent(), $oldTitle, $newTitle);
        }
        return $this;
    }

    /**
     * Get worksheet
     *
     * @return \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet
     */
    public function getWorksheet()
    {
        return $this->worksheet;
    }

    /**
     * Set worksheet
     *
     * @return \PhpOffice\PhpSpreadsheet\NamedRange
     */
    public function setWorksheet(\PhpOffice\PhpSpreadsheet\Worksheet\Worksheet $phpExcelWorksheet = \null)
    {
        if ($phpExcelWorksheet !== \null) {
            $this->worksheet = $phpExcelWorksheet;
        }
        return $this;
    }

    /**
     * Get range
     *
     * @return string
     */
    public function getRange()
    {
        return $this->range;
    }

    /**
     * Set range
     *
     * @param string $value
     * @return \PhpOffice\PhpSpreadsheet\NamedRange
     */
    public function setRange($value = \null)
    {
        if ($value !== \null) {
            $this->range = $value;
        }
        return $this;
    }

    /**
     * Get localOnly
     *
     * @return bool
     */
    public function isLocalOnly()
    {
        return $this->localOnly;
    }

    /**
     * Set localOnly
     *
     * @param bool $value
     * @return \PhpOffice\PhpSpreadsheet\NamedRange
     */
    public function setLocalOnly($value = \false)
    {
        $this->localOnly = $value;
        $this->scope = $value ? $this->worksheet : \null;
        return $this;
    }

    /**
     * Get scope
     *
     * @return \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet|null
     */
    public function getScope()
    {
        return $this->scope;
    }

    /**
     * Set scope
     *
     * @param \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet|null $phpExcelWorksheet
     * @return \PhpOffice\PhpSpreadsheet\NamedRange
     */
    public function setScope(\PhpOffice\PhpSpreadsheet\Worksheet\Worksheet $phpExcelWorksheet = \null)
    {
        $this->scope = $phpExcelWorksheet;
        $this->localOnly = $phpExcelWorksheet != \null;
        return $this;
    }

    /**
     * Resolve a named range to a regular cell range
     *
     * @param string $pNamedRange Named range
     * @param \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet|null $phpExcelWorksheet Scope. Use null for global scope
     * @return \PhpOffice\PhpSpreadsheet\NamedRange
     */
    public static function resolveRange($pNamedRange = '', \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet $phpExcelWorksheet)
    {
        return $phpExcelWorksheet->getParent()->getNamedRange($pNamedRange, $phpExcelWorksheet);
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

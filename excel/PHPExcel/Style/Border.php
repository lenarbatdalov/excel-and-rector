<?php

namespace PhpOffice\PhpSpreadsheet\Style;

/**
 * PHPExcel_Style_Border
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
 * @package    PHPExcel_Style
 * @copyright  Copyright (c) 2006 - 2015 PHPExcel (http://www.codeplex.com/PHPExcel)
 * @license    http://www.gnu.org/licenses/old-licenses/lgpl-2.1.txt    LGPL
 * @version    ##VERSION##, ##DATE##
 */
class Border extends \PhpOffice\PhpSpreadsheet\Style\Supervisor implements \PhpOffice\PhpSpreadsheet\IComparable
{
    /* Border style */
    const BORDER_NONE             = 'none';
    const BORDER_DASHDOT          = 'dashDot';
    const BORDER_DASHDOTDOT       = 'dashDotDot';
    const BORDER_DASHED           = 'dashed';
    const BORDER_DOTTED           = 'dotted';
    const BORDER_DOUBLE           = 'double';
    const BORDER_HAIR             = 'hair';
    const BORDER_MEDIUM           = 'medium';
    const BORDER_MEDIUMDASHDOT    = 'mediumDashDot';
    const BORDER_MEDIUMDASHDOTDOT = 'mediumDashDotDot';
    const BORDER_MEDIUMDASHED     = 'mediumDashed';
    const BORDER_SLANTDASHDOT     = 'slantDashDot';
    const BORDER_THICK            = 'thick';
    const BORDER_THIN             = 'thin';

    /**
     * Border style
     *
     * @var string
     */
    protected $borderStyle = \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_NONE;

    /**
     * Border color
     *
     * @var \PhpOffice\PhpSpreadsheet\Style\Color
     */
    protected $color;

    /**
     * Parent property name
     *
     * @var string
     */
    protected $parentPropertyName;

    /**
     * Create a new PHPExcel_Style_Border
     *
     * @param    boolean    $isSupervisor    Flag indicating if this is a supervisor or not
     *                                    Leave this value at default unless you understand exactly what
     *                                        its ramifications are
     * @param    boolean    $isConditional    Flag indicating if this is a conditional style or not
     *                                    Leave this value at default unless you understand exactly what
     *                                        its ramifications are
     */
    public function __construct($isSupervisor = \false)
    {
        // Supervisor?
        parent::__construct($isSupervisor);

        // Initialise values
        $this->color    = new \PhpOffice\PhpSpreadsheet\Style\Color(\PhpOffice\PhpSpreadsheet\Style\Color::COLOR_BLACK, $isSupervisor);

        // bind parent if we are a supervisor
        if ($isSupervisor) {
            $this->color->bindParent($this, 'color');
        }
    }

    /**
     * Bind parent. Only used for supervisor
     *
     * @param \PhpOffice\PhpSpreadsheet\Style\Borders $parent
     * @param string $parentPropertyName
     * @return \PhpOffice\PhpSpreadsheet\Style\Border
     */
    public function bindParent($parent, $parentPropertyName = \null)
    {
        $this->parent = $parent;
        $this->parentPropertyName = $parentPropertyName;
        return $this;
    }

    /**
     * Get the shared style component for the currently active cell in currently active sheet.
     * Only used for style supervisor
     *
     * @return \PhpOffice\PhpSpreadsheet\Style\Border
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     */
    public function getSharedComponent()
    {
        switch ($this->parentPropertyName) {
            case 'allBorders':
            case 'horizontal':
            case 'inside':
            case 'outline':
            case 'vertical':
                throw new \PhpOffice\PhpSpreadsheet\Exception('Cannot get shared component for a pseudo-border.');
            case 'bottom':
                return $this->parent->getSharedComponent()->getBottom();
            case 'diagonal':
                return $this->parent->getSharedComponent()->getDiagonal();
            case 'left':
                return $this->parent->getSharedComponent()->getLeft();
            case 'right':
                return $this->parent->getSharedComponent()->getRight();
            case 'top':
                return $this->parent->getSharedComponent()->getTop();
        }
    }

    /**
     * Build style array from subcomponents
     *
     * @param array $array
     * @return array
     */
    public function getStyleArray($array)
    {
        switch ($this->parentPropertyName) {
            case 'allBorders':
            case 'bottom':
            case 'diagonal':
            case 'horizontal':
            case 'inside':
            case 'left':
            case 'outline':
            case 'right':
            case 'top':
            case 'vertical':
                $key = \strtolower('vertical');
                break;
        }
        return $this->parent->getStyleArray(array($key => $array));
    }

    /**
     * Apply styles from array
     *
     * <code>
     * $objPHPExcel->getActiveSheet()->getStyle('B2')->getBorders()->getTop()->applyFromArray(
     *        array(
     *            'style' => PHPExcel_Style_Border::BORDER_DASHDOT,
     *            'color' => array(
     *                'rgb' => '808080'
     *            )
     *        )
     * );
     * </code>
     *
     * @param    array    $pStyles    Array containing style information
     * @throws    \PhpOffice\PhpSpreadsheet\Exception
     * @return \PhpOffice\PhpSpreadsheet\Style\Border
     */
    public function applyFromArray($pStyles = \null)
    {
        if (\is_array($pStyles)) {
            if ($this->isSupervisor) {
                $this->getActiveSheet()->getStyle($this->getSelectedCells())->applyFromArray($this->getStyleArray($pStyles), true);
            } else {
                if (isset($pStyles['style'])) {
                    $this->setBorderStyle($pStyles['style']);
                }
                if (isset($pStyles['color'])) {
                    $this->getColor()->applyFromArray($pStyles['color']);
                }
            }
        } else {
            throw new \PhpOffice\PhpSpreadsheet\Exception("Invalid style array passed.");
        }
        return $this;
    }

    /**
     * Get Border style
     *
     * @return string
     */
    public function getBorderStyle()
    {
        if ($this->isSupervisor) {
            return $this->getSharedComponent()->getBorderStyle();
        }
        return $this->borderStyle;
    }

    /**
     * Set Border style
     *
     * @param string|boolean    $pValue
     *                            When passing a boolean, FALSE equates PHPExcel_Style_Border::BORDER_NONE
     *                                and TRUE to PHPExcel_Style_Border::BORDER_MEDIUM
     * @return \PhpOffice\PhpSpreadsheet\Style\Border
     */
    public function setBorderStyle($pValue = \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_NONE)
    {

        if (empty($pValue)) {
            $pValue = \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_NONE;
        } elseif (\is_bool($pValue) && $pValue) {
            $pValue = \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_MEDIUM;
        }
        if ($this->isSupervisor) {
            $styleArray = $this->getStyleArray(array('style' => $pValue));
            $this->getActiveSheet()->getStyle($this->getSelectedCells())->applyFromArray($styleArray, true);
        } else {
            $this->borderStyle = $pValue;
        }
        return $this;
    }

    /**
     * Get Border Color
     *
     * @return \PhpOffice\PhpSpreadsheet\Style\Color
     */
    public function getColor()
    {
        return $this->color;
    }

    /**
     * Set Border Color
     *
     * @throws    \PhpOffice\PhpSpreadsheet\Exception
     * @return \PhpOffice\PhpSpreadsheet\Style\Border
     */
    public function setColor(\PhpOffice\PhpSpreadsheet\Style\Color $phpExcelStyleColor = \null)
    {
        // make sure parameter is a real color and not a supervisor
        $color = $phpExcelStyleColor->getIsSupervisor() ? $phpExcelStyleColor->getSharedComponent() : $phpExcelStyleColor;

        if ($this->isSupervisor) {
            $styleArray = $this->getColor()->getStyleArray(array('argb' => $color->getARGB()));
            $this->getActiveSheet()->getStyle($this->getSelectedCells())->applyFromArray($styleArray, true);
        } else {
            $this->color = $color;
        }
        return $this;
    }

    /**
     * Get hash code
     *
     * @return string    Hash code
     */
    public function getHashCode()
    {
        if ($this->isSupervisor) {
            return $this->getSharedComponent()->getHashCode();
        }
        return \md5(
            $this->borderStyle .
            $this->color->getHashCode() .
            __CLASS__
        );
    }
}

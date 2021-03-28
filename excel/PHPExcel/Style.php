<?php

namespace PhpOffice\PhpSpreadsheet\Style;

/**
 * PHPExcel_Style
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
class Style extends \PhpOffice\PhpSpreadsheet\Style\Supervisor implements \PhpOffice\PhpSpreadsheet\IComparable
{
    /**
     * Font
     *
     * @var \PhpOffice\PhpSpreadsheet\Style\Font
     */
    protected $font;

    /**
     * Fill
     *
     * @var \PhpOffice\PhpSpreadsheet\Style\Fill
     */
    protected $fill;

    /**
     * Borders
     *
     * @var \PhpOffice\PhpSpreadsheet\Style\Borders
     */
    protected $borders;

    /**
     * Alignment
     *
     * @var \PhpOffice\PhpSpreadsheet\Style\Alignment
     */
    protected $alignment;

    /**
     * Number Format
     *
     * @var \PhpOffice\PhpSpreadsheet\Style\NumberFormat
     */
    protected $numberFormat;

    /**
     * Conditional styles
     *
     * @var \PhpOffice\PhpSpreadsheet\Style\Conditional[]
     */
    protected $conditionalStyles;

    /**
     * Protection
     *
     * @var \PhpOffice\PhpSpreadsheet\Style\Protection
     */
    protected $protection;

    /**
     * Index of style in collection. Only used for real style.
     *
     * @var int
     */
    protected $index;

    /**
     * Use Quote Prefix when displaying in cell editor. Only used for real style.
     *
     * @var boolean
     */
    protected $quotePrefix = \false;

    /**
     * Create a new PHPExcel_Style
     *
     * @param boolean $isSupervisor Flag indicating if this is a supervisor or not
     *         Leave this value at default unless you understand exactly what
     *    its ramifications are
     * @param boolean $isConditional Flag indicating if this is a conditional style or not
     *       Leave this value at default unless you understand exactly what
     *    its ramifications are
     */
    public function __construct($isSupervisor = \false, $isConditional = \false)
    {
        // Supervisor?
        $this->isSupervisor = $isSupervisor;

        // Initialise values
        $this->conditionalStyles = array();
        $this->font         = new \PhpOffice\PhpSpreadsheet\Style\Font($isSupervisor, $isConditional);
        $this->fill         = new \PhpOffice\PhpSpreadsheet\Style\Fill($isSupervisor, $isConditional);
        $this->borders      = new \PhpOffice\PhpSpreadsheet\Style\Borders($isSupervisor, $isConditional);
        $this->alignment    = new \PhpOffice\PhpSpreadsheet\Style\Alignment($isSupervisor, $isConditional);
        $this->numberFormat = new \PhpOffice\PhpSpreadsheet\Style\NumberFormat($isSupervisor, $isConditional);
        $this->protection   = new \PhpOffice\PhpSpreadsheet\Style\Protection($isSupervisor, $isConditional);

        // bind parent if we are a supervisor
        if ($isSupervisor) {
            $this->font->bindParent($this);
            $this->fill->bindParent($this);
            $this->borders->bindParent($this);
            $this->alignment->bindParent($this);
            $this->numberFormat->bindParent($this);
            $this->protection->bindParent($this);
        }
    }

    /**
     * Get the shared style component for the currently active cell in currently active sheet.
     * Only used for style supervisor
     *
     * @return \PhpOffice\PhpSpreadsheet\Style\Style
     */
    public function getSharedComponent()
    {
        $activeSheet = $this->getActiveSheet();
        $selectedCell = $this->getActiveCell(); $xfIndex = $activeSheet->cellExists($selectedCell) ? $activeSheet->getCell($selectedCell, true)->getXfIndex() : 0;

        return $this->parent->getCellXfByIndex($xfIndex);
    }

    /**
     * Get parent. Only used for style supervisor
     *
     * @return \PhpOffice\PhpSpreadsheet\Spreadsheet
     */
    public function getParent()
    {
        return $this->parent;
    }

    /**
     * Build style array from subcomponents
     *
     * @param array $array
     * @return array
     */
    public function getStyleArray($array)
    {
        return array('quotePrefix' => $array);
    }

    /**
     * Apply styles from array
     *
     * <code>
     * $objPHPExcel->getActiveSheet()->getStyle('B2')->applyFromArray(
     *         array(
     *             'font'    => array(
     *                 'name'      => 'Arial',
     *                 'bold'      => true,
     *                 'italic'    => false,
     *                 'underline' => PHPExcel_Style_Font::UNDERLINE_DOUBLE,
     *                 'strike'    => false,
     *                 'color'     => array(
     *                     'rgb' => '808080'
     *                 )
     *             ),
     *             'borders' => array(
     *                 'bottom'     => array(
     *                     'style' => PHPExcel_Style_Border::BORDER_DASHDOT,
     *                     'color' => array(
     *                         'rgb' => '808080'
     *                     )
     *                 ),
     *                 'top'     => array(
     *                     'style' => PHPExcel_Style_Border::BORDER_DASHDOT,
     *                     'color' => array(
     *                         'rgb' => '808080'
     *                     )
     *                 )
     *             ),
     *             'quotePrefix'    => true
     *         )
     * );
     * </code>
     *
     * @param    array    $pStyles    Array containing style information
     * @param     boolean        $pAdvanced    Advanced mode for setting borders.
     * @throws    \PhpOffice\PhpSpreadsheet\Exception
     * @return \PhpOffice\PhpSpreadsheet\Style\Style
     */
    public function applyFromArray($pStyles = \null, $pAdvanced = \true)
    {
        if (\is_array($pStyles)) {
            if ($this->isSupervisor) {
                $pRange = $this->getSelectedCells();

                // Uppercase coordinate
                $pRange = \strtoupper($pRange);

                // Is it a cell range or a single cell?
                if (\strpos($pRange, ':') === \false) {
                    $rangeA = $pRange;
                    $rangeB = $pRange;
                } else {
                    list($rangeA, $rangeB) = \explode(':', $pRange);
                }

                // Calculate range outer borders
                $rangeStart = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::coordinateFromString($rangeA);
                $rangeEnd   = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::coordinateFromString($rangeB);

                // Translate column into index
                $rangeStart[0]    = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::columnIndexFromString($rangeStart[0]) - 1;
                $rangeEnd[0]    = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::columnIndexFromString($rangeEnd[0]) - 1;

                // Make sure we can loop upwards on rows and columns
                if ($rangeStart[0] > $rangeEnd[0] && $rangeStart[1] > $rangeEnd[1]) {
                    $tmp = $rangeStart;
                    $rangeStart = $rangeEnd;
                    $rangeEnd = $tmp;
                }

                // ADVANCED MODE:
                if ($pAdvanced && isset($pStyles['borders'])) {
                    // 'allborders' is a shorthand property for 'outline' and 'inside' and
                    //        it applies to components that have not been set explicitly
                    if (isset($pStyles['borders']['allborders'])) {
                        foreach (array('outline', 'inside') as $component) {
                            if (!isset($pStyles['borders'][$component])) {
                                $pStyles['borders'][$component] = $pStyles['borders']['allborders'];
                            }
                        }
                        unset($pStyles['borders']['allborders']); // not needed any more
                    }
                    // 'outline' is a shorthand property for 'top', 'right', 'bottom', 'left'
                    //        it applies to components that have not been set explicitly
                    if (isset($pStyles['borders']['outline'])) {
                        foreach (array('top', 'right', 'bottom', 'left') as $component) {
                            if (!isset($pStyles['borders'][$component])) {
                                $pStyles['borders'][$component] = $pStyles['borders']['outline'];
                            }
                        }
                        unset($pStyles['borders']['outline']); // not needed any more
                    }
                    // 'inside' is a shorthand property for 'vertical' and 'horizontal'
                    //        it applies to components that have not been set explicitly
                    if (isset($pStyles['borders']['inside'])) {
                        foreach (array('vertical', 'horizontal') as $component) {
                            if (!isset($pStyles['borders'][$component])) {
                                $pStyles['borders'][$component] = $pStyles['borders']['inside'];
                            }
                        }
                        unset($pStyles['borders']['inside']); // not needed any more
                    }
                    // width and height characteristics of selection, 1, 2, or 3 (for 3 or more)
                    $xMax = \min($rangeEnd[0] - $rangeStart[0] + 1, 3);
                    $yMax = \min($rangeEnd[1] - $rangeStart[1] + 1, 3);

                    // loop through up to 3 x 3 = 9 regions
                    for ($x = 1; $x <= $xMax; ++$x) {
                        // start column index for region
                        $colStart = ($x == 3) ?
                            \PhpOffice\PhpSpreadsheet\Cell\Coordinate::stringFromColumnIndex($rangeEnd[0])
                                : \PhpOffice\PhpSpreadsheet\Cell\Coordinate::stringFromColumnIndex($rangeStart[0] + $x - 1);
                        // end column index for region
                        $colEnd = ($x == 1) ?
                            \PhpOffice\PhpSpreadsheet\Cell\Coordinate::stringFromColumnIndex($rangeStart[0])
                                : \PhpOffice\PhpSpreadsheet\Cell\Coordinate::stringFromColumnIndex($rangeEnd[0] - $xMax + $x);

                        for ($y = 1; $y <= $yMax; ++$y) {
                            // which edges are touching the region
                            $edges = array();
                            if ($x == 1) {
                                // are we at left edge
                                $edges[] = 'left';
                            }
                            if ($x == $xMax) {
                                // are we at right edge
                                $edges[] = 'right';
                            }
                            if ($y == 1) {
                                // are we at top edge?
                                $edges[] = 'top';
                            }
                            if ($y == $yMax) {
                                // are we at bottom edge?
                                $edges[] = 'bottom';
                            }

                            // start row index for region
                            $rowStart = ($y == 3) ?
                                $rangeEnd[1] : $rangeStart[1] + $y - 1;

                            // end row index for region
                            $rowEnd = ($y == 1) ?
                                $rangeStart[1] : $rangeEnd[1] - $yMax + $y;

                            // build range for region
                            $range = $colStart . $rowStart . ':' . $colEnd . $rowEnd;

                            // retrieve relevant style array for region
                            $regionStyles = $pStyles;
                            unset($regionStyles['borders']['inside']);

                            // what are the inner edges of the region when looking at the selection
                            $innerEdges = \array_diff(array('top', 'right', 'bottom', 'left'), $edges);

                            // inner edges that are not touching the region should take the 'inside' border properties if they have been set
                            foreach ($innerEdges as $innerEdge) {
                                switch ($innerEdge) {
                                    case 'top':
                                    case 'bottom':
                                        // should pick up 'horizontal' border property if set
                                        if (isset($pStyles['borders']['horizontal'])) {
                                            $regionStyles['borders'][$innerEdge] = $pStyles['borders']['horizontal'];
                                        } else {
                                            unset($regionStyles['borders'][$innerEdge]);
                                        }
                                        break;
                                    case 'left':
                                    case 'right':
                                        // should pick up 'vertical' border property if set
                                        if (isset($pStyles['borders']['vertical'])) {
                                            $regionStyles['borders'][$innerEdge] = $pStyles['borders']['vertical'];
                                        } else {
                                            unset($regionStyles['borders'][$innerEdge]);
                                        }
                                        break;
                                }
                            }

                            // apply region style to region by calling applyFromArray() in simple mode
                            $this->getActiveSheet()->getStyle($range)->applyFromArray($regionStyles, \false);
                        }
                    }
                    return $this;
                }

                // SIMPLE MODE:
                // Selection type, inspect
                if (\preg_match('/^[A-Z]+1:[A-Z]+1048576$/', $pRange)) {
                    $selectionType = 'COLUMN';
                } elseif (\preg_match('/^A\d+:XFD\d+$/', $pRange)) {
                    $selectionType = 'ROW';
                } else {
                    $selectionType = 'CELL';
                }

                // First loop through columns, rows, or cells to find out which styles are affected by this operation
                switch ($selectionType) {
                    case 'COLUMN':
                        $oldXfIndexes = array();
                        for ($col = $rangeStart[0]; $col <= $rangeEnd[0]; ++$col) {
                            $oldXfIndexes[$this->getActiveSheet()->getColumnDimensionByColumn($col)->getXfIndex()] = \true;
                        }
                        break;
                    case 'ROW':
                        $oldXfIndexes = array();
                        for ($row = $rangeStart[1]; $row <= $rangeEnd[1]; ++$row) {
                            if ($this->getActiveSheet()->getRowDimension($row, true)->getXfIndex() == \null) {
                                $oldXfIndexes[0] = \true; // row without explicit style should be formatted based on default style
                            } else {
                                $oldXfIndexes[$this->getActiveSheet()->getRowDimension($row, true)->getXfIndex()] = \true;
                            }
                        }
                        break;
                    case 'CELL':
                        $oldXfIndexes = array();
                        for ($col = $rangeStart[0]; $col <= $rangeEnd[0]; ++$col) {
                            for ($row = $rangeStart[1]; $row <= $rangeEnd[1]; ++$row) {
                                $oldXfIndexes[$this->getActiveSheet()->getCellByColumnAndRow($col + 1, $row, true)->getXfIndex()] = \true;
                            }
                        }
                        break;
                }

                // clone each of the affected styles, apply the style array, and add the new styles to the workbook
                $phpExcel = $this->getActiveSheet()->getParent();
                foreach (array_keys($oldXfIndexes) as $oldXfIndex) {
                    $style = $phpExcel->getCellXfByIndex($oldXfIndex);
                    $newStyle = clone $style;
                    $newStyle->applyFromArray($pStyles, true);

                    if ($existingStyle = $phpExcel->getCellXfByHashCode($newStyle->getHashCode())) {
                        // there is already such cell Xf in our collection
                        $newXfIndexes[$oldXfIndex] = $existingStyle->getIndex();
                    } else {
                        // we don't have such a cell Xf, need to add
                        $phpExcel->addCellXf($newStyle);
                        $newXfIndexes[$oldXfIndex] = $newStyle->getIndex();
                    }
                }

                // Loop through columns, rows, or cells again and update the XF index
                switch ($selectionType) {
                    case 'COLUMN':
                        for ($col = $rangeStart[0]; $col <= $rangeEnd[0]; ++$col) {
                            $columnDimension = $this->getActiveSheet()->getColumnDimensionByColumn($col);
                            $oldXfIndex = $columnDimension->getXfIndex();
                            $columnDimension->setXfIndex($newXfIndexes[$oldXfIndex]);
                        }
                        break;

                    case 'ROW':
                        for ($row = $rangeStart[1]; $row <= $rangeEnd[1]; ++$row) {
                            $rowDimension = $this->getActiveSheet()->getRowDimension($row, true);
                            $oldXfIndex = $rowDimension->getXfIndex() === \null ?
                                0 : $rowDimension->getXfIndex(); // row without explicit style should be formatted based on default style
                            $rowDimension->setXfIndex($newXfIndexes[$oldXfIndex]);
                        }
                        break;

                    case 'CELL':
                        for ($col = $rangeStart[0]; $col <= $rangeEnd[0]; ++$col) {
                            for ($row = $rangeStart[1]; $row <= $rangeEnd[1]; ++$row) {
                                $cell = $this->getActiveSheet()->getCellByColumnAndRow($col + 1, $row, true);
                                $oldXfIndex = $cell->getXfIndex();
                                $cell->setXfIndex($newXfIndexes[$oldXfIndex]);
                            }
                        }
                        break;
                }

            } else {
                // not a supervisor, just apply the style array directly on style object
                if (\array_key_exists('fill', $pStyles)) {
                    $this->getFill()->applyFromArray($pStyles['fill']);
                }
                if (\array_key_exists('font', $pStyles)) {
                    $this->getFont()->applyFromArray($pStyles['font']);
                }
                if (\array_key_exists('borders', $pStyles)) {
                    $this->getBorders()->applyFromArray($pStyles['borders']);
                }
                if (\array_key_exists('alignment', $pStyles)) {
                    $this->getAlignment()->applyFromArray($pStyles['alignment']);
                }
                if (\array_key_exists('numberformat', $pStyles)) {
                    $this->getNumberFormat()->applyFromArray($pStyles['numberformat']);
                }
                if (\array_key_exists('protection', $pStyles)) {
                    $this->getProtection()->applyFromArray($pStyles['protection']);
                }
                if (\array_key_exists('quotePrefix', $pStyles)) {
                    $this->quotePrefix = $pStyles['quotePrefix'];
                }
            }
        } else {
            throw new \PhpOffice\PhpSpreadsheet\Exception("Invalid style array passed.");
        }
        return $this;
    }

    /**
     * Get Fill
     *
     * @return \PhpOffice\PhpSpreadsheet\Style\Fill
     */
    public function getFill()
    {
        return $this->fill;
    }

    /**
     * Get Font
     *
     * @return \PhpOffice\PhpSpreadsheet\Style\Font
     */
    public function getFont()
    {
        return $this->font;
    }

    /**
     * Set font
     *
     * @return \PhpOffice\PhpSpreadsheet\Style\Style
     */
    public function setFont(\PhpOffice\PhpSpreadsheet\Style\Font $phpExcelStyleFont)
    {
        $this->font = $phpExcelStyleFont;
        return $this;
    }

    /**
     * Get Borders
     *
     * @return \PhpOffice\PhpSpreadsheet\Style\Borders
     */
    public function getBorders()
    {
        return $this->borders;
    }

    /**
     * Get Alignment
     *
     * @return \PhpOffice\PhpSpreadsheet\Style\Alignment
     */
    public function getAlignment()
    {
        return $this->alignment;
    }

    /**
     * Get Number Format
     *
     * @return \PhpOffice\PhpSpreadsheet\Style\NumberFormat
     */
    public function getNumberFormat()
    {
        return $this->numberFormat;
    }

    /**
     * Get Conditional Styles. Only used on supervisor.
     *
     * @return \PhpOffice\PhpSpreadsheet\Style\Conditional[]
     */
    public function getConditionalStyles()
    {
        return $this->getActiveSheet()->getConditionalStyles($this->getActiveCell());
    }

    /**
     * Set Conditional Styles. Only used on supervisor.
     *
     * @param \PhpOffice\PhpSpreadsheet\Style\Conditional[] $pValue Array of condtional styles
     * @return \PhpOffice\PhpSpreadsheet\Style\Style
     */
    public function setConditionalStyles($pValue = \null)
    {
        if (\is_array($pValue)) {
            $this->getActiveSheet()->setConditionalStyles($this->getSelectedCells(), $pValue);
        }
        return $this;
    }

    /**
     * Get Protection
     *
     * @return \PhpOffice\PhpSpreadsheet\Style\Protection
     */
    public function getProtection()
    {
        return $this->protection;
    }

    /**
     * Get quote prefix
     *
     * @return boolean
     */
    public function getQuotePrefix()
    {
        if ($this->isSupervisor) {
            return $this->getSharedComponent()->getQuotePrefix();
        }
        return $this->quotePrefix;
    }

    /**
     * Set quote prefix
     *
     * @param boolean $pValue
     */
    public function setQuotePrefix($pValue)
    {
        if ($pValue == '') {
            $pValue = \false;
        }
        if ($this->isSupervisor) {
            $styleArray = array('quotePrefix' => $pValue);
            $this->getActiveSheet()->getStyle($this->getSelectedCells())->applyFromArray($styleArray, true);
        } else {
            $this->quotePrefix = $pValue;
        }
        return $this;
    }

    /**
     * Get hash code
     *
     * @return string Hash code
     */
    public function getHashCode()
    {
        $hashConditionals = '';
        foreach ($this->conditionalStyles as $conditionalStyle) {
            $hashConditionals .= $conditionalStyle->getHashCode();
        }

        return \md5(
            $this->fill->getHashCode() .
            $this->font->getHashCode() .
            $this->borders->getHashCode() .
            $this->alignment->getHashCode() .
            $this->numberFormat->getHashCode() .
            $hashConditionals .
            $this->protection->getHashCode() .
            ($this->quotePrefix  ? 't' : 'f') .
            __CLASS__
        );
    }

    /**
     * Get own index in style collection
     *
     * @return int
     */
    public function getIndex()
    {
        return $this->index;
    }

    /**
     * Set own index in style collection
     *
     * @param int $pValue
     */
    public function setIndex($pValue)
    {
        $this->index = $pValue;
    }
}

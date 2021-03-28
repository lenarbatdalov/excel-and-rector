<?php

namespace PhpOffice\PhpSpreadsheet;

/**
 * PHPExcel_ReferenceHelper (Singleton)
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
class ReferenceHelper
{
    /**    Constants                */
    /**    Regular Expressions      */
    const REFHELPER_REGEXP_CELLREF      = '((\w*|\'[^!]*\')!)?(?<![:a-z\$])(\$?[a-z]{1,3}\$?\d+)(?=[^:!\d\'])';
    const REFHELPER_REGEXP_CELLRANGE    = '((\w*|\'[^!]*\')!)?(\$?[a-z]{1,3}\$?\d+):(\$?[a-z]{1,3}\$?\d+)';
    const REFHELPER_REGEXP_ROWRANGE     = '((\w*|\'[^!]*\')!)?(\$?\d+):(\$?\d+)';
    const REFHELPER_REGEXP_COLRANGE     = '((\w*|\'[^!]*\')!)?(\$?[a-z]{1,3}):(\$?[a-z]{1,3})';

    /**
     * Instance of this class
     *
     * @var \PhpOffice\PhpSpreadsheet\ReferenceHelper
     */
    private static $phpExcelReferenceHelper;

    /**
     * Get an instance of this class
     *
     * @return \PhpOffice\PhpSpreadsheet\ReferenceHelper
     */
    public static function getInstance()
    {
        if (!isset(self::$phpExcelReferenceHelper) || (self::$phpExcelReferenceHelper === \null)) {
            self::$phpExcelReferenceHelper = new \PhpOffice\PhpSpreadsheet\ReferenceHelper();
        }

        return self::$phpExcelReferenceHelper;
    }

    /**
     * Create a new PHPExcel_ReferenceHelper
     */
    protected function __construct()
    {
    }

    /**
     * Compare two column addresses
     * Intended for use as a Callback function for sorting column addresses by column
     *
     * @param   string   $a  First column to test (e.g. 'AA')
     * @param   string   $b  Second column to test (e.g. 'Z')
     * @return  integer
     */
    public static function columnSort($a, $b)
    {
        return \strcasecmp(\strlen($a) . $a, \strlen($b) . $b);
    }

    /**
     * Compare two column addresses
     * Intended for use as a Callback function for reverse sorting column addresses by column
     *
     * @param   string   $a  First column to test (e.g. 'AA')
     * @param   string   $b  Second column to test (e.g. 'Z')
     * @return  integer
     */
    public static function columnReverseSort($a, $b)
    {
        return 1 - \strcasecmp(\strlen($a) . $a, \strlen($b) . $b);
    }

    /**
     * Compare two cell addresses
     * Intended for use as a Callback function for sorting cell addresses by column and row
     *
     * @param   string   $a  First cell to test (e.g. 'AA1')
     * @param   string   $b  Second cell to test (e.g. 'Z1')
     * @return  integer
     */
    public static function cellSort($a, $b)
    {
        \sscanf($a, '%[A-Z]%d', $ac, $ar);
        \sscanf($b, '%[A-Z]%d', $bc, $br);

        if ($ar == $br) {
            return \strcasecmp(\strlen($ac) . $ac, \strlen($bc) . $bc);
        }
        return ($ar < $br) ? -1 : 1;
    }

    /**
     * Compare two cell addresses
     * Intended for use as a Callback function for sorting cell addresses by column and row
     *
     * @param   string   $a  First cell to test (e.g. 'AA1')
     * @param   string   $b  Second cell to test (e.g. 'Z1')
     * @return  integer
     */
    public static function cellReverseSort($a, $b)
    {
        \sscanf($a, '%[A-Z]%d', $ac, $ar);
        \sscanf($b, '%[A-Z]%d', $bc, $br);

        if ($ar == $br) {
            return 1 - \strcasecmp(\strlen($ac) . $ac, \strlen($bc) . $bc);
        }
        return ($ar < $br) ? 1 : -1;
    }

    /**
     * Test whether a cell address falls within a defined range of cells
     *
     * @param   string     $cellAddress        Address of the cell we're testing
     * @param   integer    $beforeRow          Number of the row we're inserting/deleting before
     * @param   integer    $pNumRows           Number of rows to insert/delete (negative values indicate deletion)
     * @param   integer    $beforeColumnIndex  Index number of the column we're inserting/deleting before
     * @param   integer    $pNumCols           Number of columns to insert/delete (negative values indicate deletion)
     * @return  boolean
     */
    private static function cellAddressInDeleteRange($cellAddress, $beforeRow, $pNumRows, $beforeColumnIndex, $pNumCols)
    {
        list($cellColumn, $cellRow) = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::coordinateFromString($cellAddress);
        $cellColumnIndex = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::columnIndexFromString($cellColumn);
        //    Is cell within the range of rows/columns if we're deleting
        if ($pNumRows < 0 &&
            ($cellRow >= ($beforeRow + $pNumRows)) &&
            ($cellRow < $beforeRow)) {
            return \true;
        } elseif ($pNumCols < 0 &&
            ($cellColumnIndex >= ($beforeColumnIndex + $pNumCols)) &&
            ($cellColumnIndex < $beforeColumnIndex)) {
            return \true;
        }
        return \false;
    }

    /**
     * Update page breaks when inserting/deleting rows/columns
     *
     * @param \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet  $phpExcelWorksheet             The worksheet that we're editing
     * @param   string              $pBefore            Insert/Delete before this cell address (e.g. 'A1')
     * @param   integer             $beforeColumnIndex  Index number of the column we're inserting/deleting before
     * @param   integer             $pNumCols           Number of columns to insert/delete (negative values indicate deletion)
     * @param   integer             $beforeRow          Number of the row we're inserting/deleting before
     * @param   integer             $pNumRows           Number of rows to insert/delete (negative values indicate deletion)
     */
    protected function adjustPageBreaks(\PhpOffice\PhpSpreadsheet\Worksheet\Worksheet $phpExcelWorksheet, $pBefore, $beforeColumnIndex, $pNumCols, $beforeRow, $pNumRows)
    {
        $aBreaks = $phpExcelWorksheet->getBreaks();
        ($pNumCols > 0 || $pNumRows > 0) ?
            \uksort($aBreaks, array('PHPExcel_ReferenceHelper','cellReverseSort')) :
            \uksort($aBreaks, array('PHPExcel_ReferenceHelper','cellSort'));

        foreach ($aBreaks as $key => $value) {
            if (self::cellAddressInDeleteRange($key, $beforeRow, $pNumRows, $beforeColumnIndex, $pNumCols)) {
                //    If we're deleting, then clear any defined breaks that are within the range
                //        of rows/columns that we're deleting
                $phpExcelWorksheet->setBreak($key, \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet::BREAK_NONE);
            } else {
                //    Otherwise update any affected breaks by inserting a new break at the appropriate point
                //        and removing the old affected break
                $newReference = $this->updateCellReference($key, $pBefore, $pNumCols, $pNumRows);
                if ($key != $newReference) {
                    $phpExcelWorksheet->setBreak($newReference, $value)
                        ->setBreak($key, \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet::BREAK_NONE);
                }
            }
        }
    }

    /**
     * Update cell comments when inserting/deleting rows/columns
     *
     * @param   \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet  $pSheet             The worksheet that we're editing
     * @param   string              $pBefore            Insert/Delete before this cell address (e.g. 'A1')
     * @param   integer             $beforeColumnIndex  Index number of the column we're inserting/deleting before
     * @param   integer             $pNumCols           Number of columns to insert/delete (negative values indicate deletion)
     * @param   integer             $beforeRow          Number of the row we're inserting/deleting before
     * @param   integer             $pNumRows           Number of rows to insert/delete (negative values indicate deletion)
     */
    protected function adjustComments($pSheet, $pBefore, $beforeColumnIndex, $pNumCols, $beforeRow, $pNumRows)
    {
        $aComments = $pSheet->getComments();
        $aNewComments = array(); // the new array of all comments

        foreach ($aComments as $key => &$value) {
            // Any comments inside a deleted range will be ignored
            if (!self::cellAddressInDeleteRange($key, $beforeRow, $pNumRows, $beforeColumnIndex, $pNumCols)) {
                // Otherwise build a new array of comments indexed by the adjusted cell reference
                $newReference = $this->updateCellReference($key, $pBefore, $pNumCols, $pNumRows);
                $aNewComments[$newReference] = $value;
            }
        }
        //    Replace the comments array with the new set of comments
        $pSheet->setComments($aNewComments);
    }

    /**
     * Update hyperlinks when inserting/deleting rows/columns
     *
     * @param   \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet  $pSheet             The worksheet that we're editing
     * @param   string              $pBefore            Insert/Delete before this cell address (e.g. 'A1')
     * @param   integer             $beforeColumnIndex  Index number of the column we're inserting/deleting before
     * @param   integer             $pNumCols           Number of columns to insert/delete (negative values indicate deletion)
     * @param   integer             $beforeRow          Number of the row we're inserting/deleting before
     * @param   integer             $pNumRows           Number of rows to insert/delete (negative values indicate deletion)
     */
    protected function adjustHyperlinks($pSheet, $pBefore, $beforeColumnIndex, $pNumCols, $beforeRow, $pNumRows)
    {
        $aHyperlinkCollection = $pSheet->getHyperlinkCollection();
        ($pNumCols > 0 || $pNumRows > 0) ? \uksort($aHyperlinkCollection, array('PHPExcel_ReferenceHelper','cellReverseSort')) : \uksort($aHyperlinkCollection, array('PHPExcel_ReferenceHelper','cellSort'));

        foreach ($aHyperlinkCollection as $key => $value) {
            $newReference = $this->updateCellReference($key, $pBefore, $pNumCols, $pNumRows);
            if ($key != $newReference) {
                $pSheet->setHyperlink($newReference, $value);
                $pSheet->setHyperlink($key, \null);
            }
        }
    }

    /**
     * Update data validations when inserting/deleting rows/columns
     *
     * @param   \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet  $pSheet             The worksheet that we're editing
     * @param   string              $pBefore            Insert/Delete before this cell address (e.g. 'A1')
     * @param   integer             $beforeColumnIndex  Index number of the column we're inserting/deleting before
     * @param   integer             $pNumCols           Number of columns to insert/delete (negative values indicate deletion)
     * @param   integer             $beforeRow          Number of the row we're inserting/deleting before
     * @param   integer             $pNumRows           Number of rows to insert/delete (negative values indicate deletion)
     */
    protected function adjustDataValidations($pSheet, $pBefore, $beforeColumnIndex, $pNumCols, $beforeRow, $pNumRows)
    {
        $aDataValidationCollection = $pSheet->getDataValidationCollection();
        ($pNumCols > 0 || $pNumRows > 0) ? \uksort($aDataValidationCollection, array('PHPExcel_ReferenceHelper','cellReverseSort')) : \uksort($aDataValidationCollection, array('PHPExcel_ReferenceHelper','cellSort'));
        
        foreach ($aDataValidationCollection as $key => $value) {
            $newReference = $this->updateCellReference($key, $pBefore, $pNumCols, $pNumRows);
            if ($key != $newReference) {
                $pSheet->setDataValidation($newReference, $value);
                $pSheet->setDataValidation($key, \null);
            }
        }
    }

    /**
     * Update merged cells when inserting/deleting rows/columns
     *
     * @param   \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet  $pSheet             The worksheet that we're editing
     * @param   string              $pBefore            Insert/Delete before this cell address (e.g. 'A1')
     * @param   integer             $beforeColumnIndex  Index number of the column we're inserting/deleting before
     * @param   integer             $pNumCols           Number of columns to insert/delete (negative values indicate deletion)
     * @param   integer             $beforeRow          Number of the row we're inserting/deleting before
     * @param   integer             $pNumRows           Number of rows to insert/delete (negative values indicate deletion)
     */
    protected function adjustMergeCells($pSheet, $pBefore, $beforeColumnIndex, $pNumCols, $beforeRow, $pNumRows)
    {
        $aMergeCells = $pSheet->getMergeCells();
        $aNewMergeCells = array(); // the new array of all merge cells
        foreach (array_keys($aMergeCells) as &$key) {
            $newReference = $this->updateCellReference($key, $pBefore, $pNumCols, $pNumRows);
            $aNewMergeCells[$newReference] = $newReference;
        }
        $pSheet->setMergeCells($aNewMergeCells); // replace the merge cells array
    }

    /**
     * Update protected cells when inserting/deleting rows/columns
     *
     * @param   \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet  $pSheet             The worksheet that we're editing
     * @param   string              $pBefore            Insert/Delete before this cell address (e.g. 'A1')
     * @param   integer             $beforeColumnIndex  Index number of the column we're inserting/deleting before
     * @param   integer             $pNumCols           Number of columns to insert/delete (negative values indicate deletion)
     * @param   integer             $beforeRow          Number of the row we're inserting/deleting before
     * @param   integer             $pNumRows           Number of rows to insert/delete (negative values indicate deletion)
     */
    protected function adjustProtectedCells($pSheet, $pBefore, $beforeColumnIndex, $pNumCols, $beforeRow, $pNumRows)
    {
        $aProtectedCells = $pSheet->getProtectedCells();
        ($pNumCols > 0 || $pNumRows > 0) ?
            \uksort($aProtectedCells, array('PHPExcel_ReferenceHelper','cellReverseSort')) :
            \uksort($aProtectedCells, array('PHPExcel_ReferenceHelper','cellSort'));
        foreach ($aProtectedCells as $key => $value) {
            $newReference = $this->updateCellReference($key, $pBefore, $pNumCols, $pNumRows);
            if ($key != $newReference) {
                $pSheet->protectCells($newReference, $value, \true);
                $pSheet->unprotectCells($key);
            }
        }
    }

    /**
     * Update column dimensions when inserting/deleting rows/columns
     *
     * @param   \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet  $pSheet             The worksheet that we're editing
     * @param   string              $pBefore            Insert/Delete before this cell address (e.g. 'A1')
     * @param   integer             $beforeColumnIndex  Index number of the column we're inserting/deleting before
     * @param   integer             $pNumCols           Number of columns to insert/delete (negative values indicate deletion)
     * @param   integer             $beforeRow          Number of the row we're inserting/deleting before
     * @param   integer             $pNumRows           Number of rows to insert/delete (negative values indicate deletion)
     */
    protected function adjustColumnDimensions($pSheet, $pBefore, $beforeColumnIndex, $pNumCols, $beforeRow, $pNumRows)
    {
        $aColumnDimensions = \array_reverse($pSheet->getColumnDimensions(), \true);
        if (!empty($aColumnDimensions)) {
            foreach ($aColumnDimensions as $aColumnDimension) {
                $newReference = $this->updateCellReference($aColumnDimension->getColumnIndex() . '1', $pBefore, $pNumCols, $pNumRows);
                list($newReference) = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::coordinateFromString($newReference);
                if ($aColumnDimension->getColumnIndex() != $newReference) {
                    $aColumnDimension->setColumnIndex($newReference);
                }
            }
            $pSheet->refreshColumnDimensions();
        }
    }

    /**
     * Update row dimensions when inserting/deleting rows/columns
     *
     * @param   \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet  $pSheet             The worksheet that we're editing
     * @param   string              $pBefore            Insert/Delete before this cell address (e.g. 'A1')
     * @param   integer             $beforeColumnIndex  Index number of the column we're inserting/deleting before
     * @param   integer             $pNumCols           Number of columns to insert/delete (negative values indicate deletion)
     * @param   integer             $beforeRow          Number of the row we're inserting/deleting before
     * @param   integer             $pNumRows           Number of rows to insert/delete (negative values indicate deletion)
     */
    protected function adjustRowDimensions($pSheet, $pBefore, $beforeColumnIndex, $pNumCols, $beforeRow, $pNumRows)
    {
        $aRowDimensions = \array_reverse($pSheet->getRowDimensions(), \true);
        if (!empty($aRowDimensions)) {
            foreach ($aRowDimensions as $aRowDimension) {
                $newReference = $this->updateCellReference('A' . $aRowDimension->getRowIndex(), $pBefore, $pNumCols, $pNumRows);
                list(, $newReference) = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::coordinateFromString($newReference);
                if ($aRowDimension->getRowIndex() != $newReference) {
                    $aRowDimension->setRowIndex($newReference);
                }
            }
            $pSheet->refreshRowDimensions();

            $copyDimension = $pSheet->getRowDimension($beforeRow - 1, true);
            for ($i = $beforeRow; $i <= $beforeRow - 1 + $pNumRows; ++$i) {
                $newDimension = $pSheet->getRowDimension($i, true);
                $newDimension->setRowHeight($copyDimension->getRowHeight());
                $newDimension->setVisible($copyDimension->isVisible());
                $newDimension->setOutlineLevel($copyDimension->getOutlineLevel());
                $newDimension->setCollapsed($copyDimension->isCollapsed());
            }
        }
    }

    /**
     * Insert a new column or row, updating all possible related data
     *
     * @param   string              $pBefore    Insert before this cell address (e.g. 'A1')
     * @param   integer             $pNumCols   Number of columns to insert/delete (negative values indicate deletion)
     * @param   integer             $pNumRows   Number of rows to insert/delete (negative values indicate deletion)
     * @param \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet  $phpExcelWorksheet     The worksheet that we're editing
     * @throws  \PhpOffice\PhpSpreadsheet\Exception
     */
    public function insertNewBefore($pBefore = 'A1', $pNumCols = 0, $pNumRows = 0, \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet $phpExcelWorksheet = \null)
    {
        $remove = ($pNumCols < 0 || $pNumRows < 0);
        $aCellCollection = $phpExcelWorksheet->getCellCollection();

        // Get coordinates of $pBefore
        $beforeColumn    = 'A';
        $beforeRow        = 1;
        list($beforeColumn, $beforeRow) = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::coordinateFromString($pBefore);
        $beforeColumnIndex = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::columnIndexFromString($beforeColumn);

        // Clear cells if we are removing columns or rows
        $highestColumn    = $phpExcelWorksheet->getHighestColumn();
        $highestRow    = $phpExcelWorksheet->getHighestRow();

        // 1. Clear column strips if we are removing columns
        if ($pNumCols < 0 && $beforeColumnIndex - 2 + $pNumCols > 0) {
            for ($i = 1; $i <= $highestRow - 1; ++$i) {
                for ($j = $beforeColumnIndex - 1 + $pNumCols; $j <= $beforeColumnIndex - 2; ++$j) {
                    $coordinate = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::stringFromColumnIndex($j) . $i;
                    $phpExcelWorksheet->removeConditionalStyles($coordinate);
                    if ($phpExcelWorksheet->cellExists($coordinate)) {
                        $phpExcelWorksheet->getCell($coordinate)->setValueExplicit('', \PhpOffice\PhpSpreadsheet\Cell\DataType::TYPE_NULL);
                        $phpExcelWorksheet->getCell($coordinate)->setXfIndex(0);
                    }
                }
            }
        }

        // 2. Clear row strips if we are removing rows
        if ($pNumRows < 0 && $beforeRow - 1 + $pNumRows > 0) {
            for ($i = $beforeColumnIndex - 1; $i <= \PhpOffice\PhpSpreadsheet\Cell\Coordinate::columnIndexFromString($highestColumn) - 1; ++$i) {
                for ($j = $beforeRow + $pNumRows; $j <= $beforeRow - 1; ++$j) {
                    $coordinate = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::stringFromColumnIndex($i) . $j;
                    $phpExcelWorksheet->removeConditionalStyles($coordinate);
                    if ($phpExcelWorksheet->cellExists($coordinate)) {
                        $phpExcelWorksheet->getCell($coordinate)->setValueExplicit('', \PhpOffice\PhpSpreadsheet\Cell\DataType::TYPE_NULL);
                        $phpExcelWorksheet->getCell($coordinate)->setXfIndex(0);
                    }
                }
            }
        }

        // Loop through cells, bottom-up, and change cell coordinates
        if ($remove) {
            // It's faster to reverse and pop than to use unshift, especially with large cell collections
            $aCellCollection = \array_reverse($aCellCollection);
        }
        while ($cellID = \array_pop($aCellCollection)) {
            $cell = $phpExcelWorksheet->getCell($cellID);
            $cellIndex = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::columnIndexFromString($cell->getColumn());

            if ($cellIndex-1 + $pNumCols < 0) {
                continue;
            }

            // New coordinates
            $newCoordinates = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::stringFromColumnIndex($cellIndex-1 + $pNumCols) . ($cell->getRow() + $pNumRows);

            // Should the cell be updated? Move value and cellXf index from one cell to another.
            if (($cellIndex >= $beforeColumnIndex) && ($cell->getRow() >= $beforeRow)) {
                // Update cell styles
                $phpExcelWorksheet->getCell($newCoordinates)->setXfIndex($cell->getXfIndex());
                // Insert this cell at its new location
                if ($cell->getDataType() == \PhpOffice\PhpSpreadsheet\Cell\DataType::TYPE_FORMULA) {
                    // Formula should be adjusted
                    $phpExcelWorksheet->getCell($newCoordinates)
                           ->setValue($this->updateFormulaReferences($cell->getValue(), $pBefore, $pNumCols, $pNumRows, $phpExcelWorksheet->getTitle()));
                } else {
                    // Formula should not be adjusted
                    $phpExcelWorksheet->getCell($newCoordinates)->setValue($cell->getValue());
                }
                // Clear the original cell
                $phpExcelWorksheet->getCellCacheController()->deleteCacheData($cellID);
            } elseif ($cell->getDataType() == \PhpOffice\PhpSpreadsheet\Cell\DataType::TYPE_FORMULA) {
                // Formula should be adjusted
                $cell->setValue($this->updateFormulaReferences($cell->getValue(), $pBefore, $pNumCols, $pNumRows, $phpExcelWorksheet->getTitle()));
            }
        }

        // Duplicate styles for the newly inserted cells
        $highestColumn    = $phpExcelWorksheet->getHighestColumn();
        $highestRow    = $phpExcelWorksheet->getHighestRow();

        if ($pNumCols > 0 && $beforeColumnIndex - 2 > 0) {
            for ($i = $beforeRow; $i <= $highestRow - 1; ++$i) {
                // Style
                $coordinate = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::stringFromColumnIndex($beforeColumnIndex - 2) . $i;
                if ($phpExcelWorksheet->cellExists($coordinate)) {
                    $xfIndex = $phpExcelWorksheet->getCell($coordinate)->getXfIndex();
                    $conditionalStyles = $phpExcelWorksheet->conditionalStylesExists($coordinate) ?
                        $phpExcelWorksheet->getConditionalStyles($coordinate) : \false;
                    for ($j = $beforeColumnIndex - 1; $j <= $beforeColumnIndex - 2 + $pNumCols; ++$j) {
                        $phpExcelWorksheet->getCellByColumnAndRow($j, $i)->setXfIndex($xfIndex);
                        if ($conditionalStyles) {
                            $cloned = array();
                            foreach ($conditionalStyles as $conditionalStyle) {
                                $cloned[] = clone $conditionalStyle;
                            }
                            $phpExcelWorksheet->setConditionalStyles(\PhpOffice\PhpSpreadsheet\Cell\Coordinate::stringFromColumnIndex($j) . $i, $cloned);
                        }
                    }
                }

            }
        }

        if ($pNumRows > 0 && $beforeRow - 1 > 0) {
            for ($i = $beforeColumnIndex - 1; $i <= \PhpOffice\PhpSpreadsheet\Cell\Coordinate::columnIndexFromString($highestColumn) - 1; ++$i) {
                // Style
                $coordinate = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::stringFromColumnIndex($i) . ($beforeRow - 1);
                if ($phpExcelWorksheet->cellExists($coordinate)) {
                    $xfIndex = $phpExcelWorksheet->getCell($coordinate)->getXfIndex();
                    $conditionalStyles = $phpExcelWorksheet->conditionalStylesExists($coordinate) ?
                        $phpExcelWorksheet->getConditionalStyles($coordinate) : \false;
                    for ($j = $beforeRow; $j <= $beforeRow - 1 + $pNumRows; ++$j) {
                        $phpExcelWorksheet->getCell(\PhpOffice\PhpSpreadsheet\Cell\Coordinate::stringFromColumnIndex($i) . $j)->setXfIndex($xfIndex);
                        if ($conditionalStyles) {
                            $cloned = array();
                            foreach ($conditionalStyles as $conditionalStyle) {
                                $cloned[] = clone $conditionalStyle;
                            }
                            $phpExcelWorksheet->setConditionalStyles(\PhpOffice\PhpSpreadsheet\Cell\Coordinate::stringFromColumnIndex($i) . $j, $cloned);
                        }
                    }
                }
            }
        }

        // Update worksheet: column dimensions
        $this->adjustColumnDimensions($phpExcelWorksheet, $pBefore, $beforeColumnIndex, $pNumCols, $beforeRow, $pNumRows);

        // Update worksheet: row dimensions
        $this->adjustRowDimensions($phpExcelWorksheet, $pBefore, $beforeColumnIndex, $pNumCols, $beforeRow, $pNumRows);

        //    Update worksheet: page breaks
        $this->adjustPageBreaks($phpExcelWorksheet, $pBefore, $beforeColumnIndex, $pNumCols, $beforeRow, $pNumRows);

        //    Update worksheet: comments
        $this->adjustComments($phpExcelWorksheet, $pBefore, $beforeColumnIndex, $pNumCols, $beforeRow, $pNumRows);

        // Update worksheet: hyperlinks
        $this->adjustHyperlinks($phpExcelWorksheet, $pBefore, $beforeColumnIndex, $pNumCols, $beforeRow, $pNumRows);

        // Update worksheet: data validations
        $this->adjustDataValidations($phpExcelWorksheet, $pBefore, $beforeColumnIndex, $pNumCols, $beforeRow, $pNumRows);

        // Update worksheet: merge cells
        $this->adjustMergeCells($phpExcelWorksheet, $pBefore, $beforeColumnIndex, $pNumCols, $beforeRow, $pNumRows);

        // Update worksheet: protected cells
        $this->adjustProtectedCells($phpExcelWorksheet, $pBefore, $beforeColumnIndex, $pNumCols, $beforeRow, $pNumRows);

        // Update worksheet: autofilter
        $autoFilter = $phpExcelWorksheet->getAutoFilter();
        $autoFilterRange = $autoFilter->getRange();
        if (!empty($autoFilterRange)) {
            if ($pNumCols != 0) {
                if (\count($autoFilterColumns) > 0) {
                    \sscanf($pBefore, '%[A-Z]%d', $column, $row);
                    $columnIndex = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::columnIndexFromString($column);
                    list($rangeStart, $rangeEnd) = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::rangeBoundaries($autoFilterRange);
                    if ($columnIndex <= $rangeEnd[0]) {
                        if ($pNumCols < 0) {
                            //    If we're actually deleting any columns that fall within the autofilter range,
                            //        then we delete any rules for those columns
                            $deleteColumn = $columnIndex + $pNumCols - 1;
                            $deleteCount = \abs($pNumCols);
                            for ($i = 1; $i <= $deleteCount; ++$i) {
                                if (array_key_exists(\PhpOffice\PhpSpreadsheet\Cell\Coordinate::stringFromColumnIndex($deleteColumn), $autoFilter->getColumns())) {
                                    $autoFilter->clearColumn(\PhpOffice\PhpSpreadsheet\Cell\Coordinate::stringFromColumnIndex($deleteColumn));
                                }
                                ++$deleteColumn;
                            }
                        }
                        $startCol = ($columnIndex > $rangeStart[0]) ? $columnIndex : $rangeStart[0];

                        //    Shuffle columns in autofilter range
                        if ($pNumCols > 0) {
                            //    For insert, we shuffle from end to beginning to avoid overwriting
                            $startColID = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::stringFromColumnIndex($startCol-1);
                            $toColID = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::stringFromColumnIndex($startCol+$pNumCols-1);
                            $endColID = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::stringFromColumnIndex($rangeEnd[0]);

                            $startColRef = $startCol;
                            $endColRef = $rangeEnd[0];
                            $toColRef = $rangeEnd[0]+$pNumCols;

                            do {
                                $autoFilter->shiftColumn(\PhpOffice\PhpSpreadsheet\Cell\Coordinate::stringFromColumnIndex($endColRef-1), \PhpOffice\PhpSpreadsheet\Cell\Coordinate::stringFromColumnIndex($toColRef-1));
                                --$endColRef;
                                --$toColRef;
                            } while ($startColRef <= $endColRef);
                        } else {
                            //    For delete, we shuffle from beginning to end to avoid overwriting
                            $startColID = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::stringFromColumnIndex($startCol-1);
                            $toColID = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::stringFromColumnIndex($startCol+$pNumCols-1);
                            $endColID = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::stringFromColumnIndex($rangeEnd[0]);
                            do {
                                $autoFilter->shiftColumn($startColID, $toColID);
                                ++$startColID;
                                ++$toColID;
                            } while ($startColID !== $endColID);
                        }
                    }
                }
            }
            $phpExcelWorksheet->setAutoFilter($this->updateCellReference($autoFilterRange, $pBefore, $pNumCols, $pNumRows));
        }

        // Update worksheet: freeze pane
        if ($phpExcelWorksheet->getFreezePane() != '') {
            $phpExcelWorksheet->freezePane($this->updateCellReference($phpExcelWorksheet->getFreezePane(), $pBefore, $pNumCols, $pNumRows));
        }

        // Page setup
        if ($phpExcelWorksheet->getPageSetup()->isPrintAreaSet()) {
            $phpExcelWorksheet->getPageSetup()->setPrintArea($this->updateCellReference($phpExcelWorksheet->getPageSetup()->getPrintArea(), $pBefore, $pNumCols, $pNumRows));
        }

        // Update worksheet: drawings
        $aDrawings = $phpExcelWorksheet->getDrawingCollection();
        foreach ($aDrawings as $aDrawing) {
            $newReference = $this->updateCellReference($aDrawing->getCoordinates(), $pBefore, $pNumCols, $pNumRows);
            if ($aDrawing->getCoordinates() != $newReference) {
                $aDrawing->setCoordinates($newReference);
            }
        }

        foreach ($phpExcelWorksheet->getParent()->getNamedRanges() as $phpExcelNamedRange) {
            if ($phpExcelNamedRange->getWorksheet()->getHashCode() == $phpExcelWorksheet->getHashCode()) {
                $phpExcelNamedRange->setRange($this->updateCellReference($phpExcelNamedRange->getRange(), $pBefore, $pNumCols, $pNumRows));
            }
        }

        // Garbage collect
        $phpExcelWorksheet->garbageCollect();
    }

    /**
     * Update references within formulas
     *
     * @param    string    $pFormula    Formula to update
     * @param    int        $pBefore    Insert before this one
     * @param    int        $pNumCols    Number of columns to insert
     * @param    int        $pNumRows    Number of rows to insert
     * @param   string  $sheetName  Worksheet name/title
     * @return    string    Updated formula
     * @throws    \PhpOffice\PhpSpreadsheet\Exception
     */
    public function updateFormulaReferences($pFormula = '', $pBefore = 'A1', $pNumCols = 0, $pNumRows = 0, $sheetName = '')
    {
        //    Update cell references in the formula
        $formulaBlocks = \explode('"', $pFormula);
        $i = \false;
        foreach ($formulaBlocks as &$formulaBlock) {
            //    Ignore blocks that were enclosed in quotes (alternating entries in the $formulaBlocks array after the explode)
            if ($i = !$i) {
                $adjustCount = 0;
                $newCellTokens = $cellTokens = array();
                //    Search for row ranges (e.g. 'Sheet1'!3:5 or 3:5) with or without $ absolutes (e.g. $3:5)
                $matchCount = \preg_match_all('/'.self::REFHELPER_REGEXP_ROWRANGE.'/i', ' '.$formulaBlock.' ', $matches, \PREG_SET_ORDER);
                if ($matchCount > 0) {
                    foreach ($matches as $match) {
                        $fromString = ($match[2] > '') ? $match[2].'!' : '';
                        $fromString .= $match[3].':'.$match[4];
                        $modified3 = \substr($this->updateCellReference('$A'.$match[3], $pBefore, $pNumCols, $pNumRows), 2);
                        $modified4 = \substr($this->updateCellReference('$A'.$match[4], $pBefore, $pNumCols, $pNumRows), 2);

                        if ($match[3].':'.$match[4] !== $modified3.':'.$modified4 && (($match[2] == '') || (\trim($match[2], "'") === $sheetName))) {
                            $toString = ($match[2] > '') ? $match[2].'!' : '';
                            $toString .= $modified3.':'.$modified4;
                            //    Max worksheet size is 1,048,576 rows by 16,384 columns in Excel 2007, so our adjustments need to be at least one digit more
                            $column = 100000;
                            $row = 10000000 + \trim($match[3], '$');
                            $cellIndex = $column.$row;
                            $newCellTokens[$cellIndex] = \preg_quote($toString);
                            $cellTokens[$cellIndex] = '/(?<!\d\$\!)'.\preg_quote($fromString, '/').'(?!\d)/i';
                            ++$adjustCount;
                        }
                    }
                }
                //    Search for column ranges (e.g. 'Sheet1'!C:E or C:E) with or without $ absolutes (e.g. $C:E)
                $matchCount = \preg_match_all('/'.self::REFHELPER_REGEXP_COLRANGE.'/i', ' '.$formulaBlock.' ', $matches, \PREG_SET_ORDER);
                if ($matchCount > 0) {
                    foreach ($matches as $match) {
                        $fromString = ($match[2] > '') ? $match[2].'!' : '';
                        $fromString .= $match[3].':'.$match[4];
                        $modified3 = \substr($this->updateCellReference($match[3].'$1', $pBefore, $pNumCols, $pNumRows), 0, -2);
                        $modified4 = \substr($this->updateCellReference($match[4].'$1', $pBefore, $pNumCols, $pNumRows), 0, -2);

                        if ($match[3].':'.$match[4] !== $modified3.':'.$modified4 && (($match[2] == '') || (\trim($match[2], "'") === $sheetName))) {
                            $toString = ($match[2] > '') ? $match[2].'!' : '';
                            $toString .= $modified3.':'.$modified4;
                            //    Max worksheet size is 1,048,576 rows by 16,384 columns in Excel 2007, so our adjustments need to be at least one digit more
                            $column = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::columnIndexFromString(\trim($match[3], '$')) + 100000;
                            $row = 10000000;
                            $cellIndex = $column.$row;
                            $newCellTokens[$cellIndex] = \preg_quote($toString);
                            $cellTokens[$cellIndex] = '/(?<![A-Z\$\!])'.\preg_quote($fromString, '/').'(?![A-Z])/i';
                            ++$adjustCount;
                        }
                    }
                }
                //    Search for cell ranges (e.g. 'Sheet1'!A3:C5 or A3:C5) with or without $ absolutes (e.g. $A1:C$5)
                $matchCount = \preg_match_all('/'.self::REFHELPER_REGEXP_CELLRANGE.'/i', ' '.$formulaBlock.' ', $matches, \PREG_SET_ORDER);
                if ($matchCount > 0) {
                    foreach ($matches as $match) {
                        $fromString = ($match[2] > '') ? $match[2].'!' : '';
                        $fromString .= $match[3].':'.$match[4];
                        $modified3 = $this->updateCellReference($match[3], $pBefore, $pNumCols, $pNumRows);
                        $modified4 = $this->updateCellReference($match[4], $pBefore, $pNumCols, $pNumRows);

                        if ($match[3].$match[4] !== $modified3.$modified4 && (($match[2] == '') || (\trim($match[2], "'") === $sheetName))) {
                            $toString = ($match[2] > '') ? $match[2].'!' : '';
                            $toString .= $modified3.':'.$modified4;
                            list($column, $row) = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::coordinateFromString($match[3]);
                            //    Max worksheet size is 1,048,576 rows by 16,384 columns in Excel 2007, so our adjustments need to be at least one digit more
                            $column = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::columnIndexFromString(\trim($column, '$')) + 100000;
                            $row = \trim($row, '$') + 10000000;
                            $cellIndex = $column.$row;
                            $newCellTokens[$cellIndex] = \preg_quote($toString);
                            $cellTokens[$cellIndex] = '/(?<![A-Z]\$\!)'.\preg_quote($fromString, '/').'(?!\d)/i';
                            ++$adjustCount;
                        }
                    }
                }
                //    Search for cell references (e.g. 'Sheet1'!A3 or C5) with or without $ absolutes (e.g. $A1 or C$5)
                $matchCount = \preg_match_all('/'.self::REFHELPER_REGEXP_CELLREF.'/i', ' '.$formulaBlock.' ', $matches, \PREG_SET_ORDER);

                if ($matchCount > 0) {
                    foreach ($matches as $match) {
                        $fromString = ($match[2] > '') ? $match[2].'!' : '';
                        $fromString .= $match[3];

                        $modified3 = $this->updateCellReference($match[3], $pBefore, $pNumCols, $pNumRows);
                        if ($match[3] !== $modified3 && (($match[2] == '') || (\trim($match[2], "'") === $sheetName))) {
                            $toString = ($match[2] > '') ? $match[2].'!' : '';
                            $toString .= $modified3;
                            list($column, $row) = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::coordinateFromString($match[3]);
                            //    Max worksheet size is 1,048,576 rows by 16,384 columns in Excel 2007, so our adjustments need to be at least one digit more
                            $column = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::columnIndexFromString(\trim($column, '$')) + 100000;
                            $row = \trim($row, '$') + 10000000;
                            $cellIndex = $row . $column;
                            $newCellTokens[$cellIndex] = \preg_quote($toString);
                            $cellTokens[$cellIndex] = '/(?<![A-Z\$\!])'.\preg_quote($fromString, '/').'(?!\d)/i';
                            ++$adjustCount;
                        }
                    }
                }
                if ($adjustCount > 0) {
                    if ($pNumCols > 0 || $pNumRows > 0) {
                        \krsort($cellTokens);
                        \krsort($newCellTokens);
                    } else {
                        \ksort($cellTokens);
                        \ksort($newCellTokens);
                    }   //  Update cell references in the formula
                    $formulaBlock = \str_replace('\\', '', \preg_replace($cellTokens, $newCellTokens, $formulaBlock));
                }
            }
        }
        unset($formulaBlock);

        //    Then rebuild the formula string
        return \implode('"', $formulaBlocks);
    }

    /**
     * Update cell reference
     *
     * @param    string    $pCellRange            Cell range
     * @param    int        $pBefore            Insert before this one
     * @param    int        $pNumCols            Number of columns to increment
     * @param    int        $pNumRows            Number of rows to increment
     * @return    string    Updated cell range
     * @throws    \PhpOffice\PhpSpreadsheet\Exception
     */
    public function updateCellReference($pCellRange = 'A1', $pBefore = 'A1', $pNumCols = 0, $pNumRows = 0)
    {
        // Is it in another worksheet? Will not have to update anything.
        if (\strpos($pCellRange, "!") !== \false) {
            return $pCellRange;
        // Is it a range or a single cell?
        } elseif (\strpos($pCellRange, ':') === \false && \strpos($pCellRange, ',') === \false) {
            // Single cell
            return $this->updateSingleCellReference($pCellRange, $pBefore, $pNumCols, $pNumRows);
        } elseif (\strpos($pCellRange, ':') !== \false || \strpos($pCellRange, ',') !== \false) {
            // Range
            return $this->updateCellRange($pCellRange, $pBefore, $pNumCols, $pNumRows);
        } else {
            // Return original
            return $pCellRange;
        }
    }

    /**
     * Update named formulas (i.e. containing worksheet references / named ranges)
     *
     * @param \PhpOffice\PhpSpreadsheet\Spreadsheet $pPhpExcel    Object to update
     * @param string $oldName        Old name (name to replace)
     * @param string $newName        New name
     */
    public function updateNamedFormulas(\PhpOffice\PhpSpreadsheet\Spreadsheet $pPhpExcel, $oldName = '', $newName = '')
    {
        if ($oldName == '') {
            return;
        }

        foreach ($pPhpExcel->getWorksheetIterator() as $sheet) {
            foreach ($sheet->getCoordinates(\false) as $phpExcelCell) {
                $cell = $sheet->getCell($phpExcelCell, true);
                if (($cell !== \null) && ($cell->getDataType() == \PhpOffice\PhpSpreadsheet\Cell\DataType::TYPE_FORMULA)) {
                    $formula = $cell->getValue();
                    if (\strpos($formula, $oldName) !== \false) {
                        $formula = \str_replace("'" . $oldName . "'!", "'" . $newName . "'!", $formula);
                        $formula = \str_replace($oldName . "!", $newName . "!", $formula);
                        $cell->setValueExplicit($formula, \PhpOffice\PhpSpreadsheet\Cell\DataType::TYPE_FORMULA);
                    }
                }
            }
        }
    }

    /**
     * Update cell range
     *
     * @param    string    $pCellRange            Cell range    (e.g. 'B2:D4', 'B:C' or '2:3')
     * @param    int        $pBefore            Insert before this one
     * @param    int        $pNumCols            Number of columns to increment
     * @param    int        $pNumRows            Number of rows to increment
     * @return    string    Updated cell range
     * @throws    \PhpOffice\PhpSpreadsheet\Exception
     */
    private function updateCellRange($pCellRange = 'A1:A1', $pBefore = 'A1', $pNumCols = 0, $pNumRows = 0)
    {
        if (\strpos($pCellRange, ':') !== \false || \strpos($pCellRange, ',') !== \false) {
            // Update range
            $range = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::splitRange($pCellRange);
            $ic = \count($range);
            for ($i = 0; $i < $ic; ++$i) {
                $jc = \count($range[$i]);
                for ($j = 0; $j < $jc; ++$j) {
                    if (\ctype_alpha($range[$i][$j])) {
                        $r = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::coordinateFromString($this->updateSingleCellReference($range[$i][$j].'1', $pBefore, $pNumCols, $pNumRows));
                        $range[$i][$j] = $r[0];
                    } elseif (\ctype_digit($range[$i][$j])) {
                        $r = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::coordinateFromString($this->updateSingleCellReference('A'.$range[$i][$j], $pBefore, $pNumCols, $pNumRows));
                        $range[$i][$j] = $r[1];
                    } else {
                        $range[$i][$j] = $this->updateSingleCellReference($range[$i][$j], $pBefore, $pNumCols, $pNumRows);
                    }
                }
            }

            // Recreate range string
            return \PhpOffice\PhpSpreadsheet\Cell\Coordinate::buildRange($range);
        } else {
            throw new \PhpOffice\PhpSpreadsheet\Exception("Only cell ranges may be passed to this method.");
        }
    }

    /**
     * Update single cell reference
     *
     * @param    string    $pCellReference        Single cell reference
     * @param    int        $pBefore            Insert before this one
     * @param    int        $pNumCols            Number of columns to increment
     * @param    int        $pNumRows            Number of rows to increment
     * @return    string    Updated cell reference
     * @throws    \PhpOffice\PhpSpreadsheet\Exception
     */
    private function updateSingleCellReference($pCellReference = 'A1', $pBefore = 'A1', $pNumCols = 0, $pNumRows = 0)
    {
        if (\strpos($pCellReference, ':') === \false && \strpos($pCellReference, ',') === \false) {
            // Get coordinates of $pBefore
            list($beforeColumn, $beforeRow) = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::coordinateFromString($pBefore);

            // Get coordinates of $pCellReference
            list($newColumn, $newRow) = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::coordinateFromString($pCellReference);

            // Verify which parts should be updated
            $updateColumn = (($newColumn{0} != '$') && ($beforeColumn{0} != '$') && (\PhpOffice\PhpSpreadsheet\Cell\Coordinate::columnIndexFromString($newColumn) >= \PhpOffice\PhpSpreadsheet\Cell\Coordinate::columnIndexFromString($beforeColumn)));
            $updateRow = (($newRow{0} != '$') && ($beforeRow{0} != '$') && $newRow >= $beforeRow);

            // Create new column reference
            if ($updateColumn) {
                $newColumn    = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::stringFromColumnIndex(\PhpOffice\PhpSpreadsheet\Cell\Coordinate::columnIndexFromString($newColumn) - 1 + $pNumCols);
            }

            // Create new row reference
            if ($updateRow) {
                $newRow += $pNumRows;
            }

            // Return new reference
            return $newColumn . $newRow;
        } else {
            throw new \PhpOffice\PhpSpreadsheet\Exception("Only single cell references may be passed to this method.");
        }
    }

    /**
     * __clone implementation. Cloning should not be allowed in a Singleton!
     *
     * @throws    \PhpOffice\PhpSpreadsheet\Exception
     */
    final public function __clone()
    {
        throw new \PhpOffice\PhpSpreadsheet\Exception("Cloning a Singleton is not allowed!");
    }
}

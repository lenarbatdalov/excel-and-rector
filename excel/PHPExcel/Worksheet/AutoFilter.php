<?php

namespace PhpOffice\PhpSpreadsheet\Worksheet;

/**
 * PHPExcel_Worksheet_AutoFilter
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
class AutoFilter
{
    /**
     * Autofilter Worksheet
     *
     * @var \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet
     */
    private $phpExcelWorksheet;


    /**
     * Autofilter Range
     *
     * @var string
     */
    private $range = '';


    /**
     * Autofilter Column Ruleset
     *
     * @var array of PHPExcel_Worksheet_AutoFilter_Column
     */
    private $columns = array();


    /**
     * Create a new PHPExcel_Worksheet_AutoFilter
     *
     *    @param    string        $pRange        Cell range (i.e. A1:E10)
     */
    public function __construct($pRange = '', \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet $phpExcelWorksheet = \null)
    {
        $this->range = $pRange;
        $this->phpExcelWorksheet = $phpExcelWorksheet;
    }

    /**
     * Get AutoFilter Parent Worksheet
     *
     * @return \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet
     */
    public function getParent()
    {
        return $this->phpExcelWorksheet;
    }

    /**
     * Set AutoFilter Parent Worksheet
     *
     * @return \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter
     */
    public function setParent(\PhpOffice\PhpSpreadsheet\Worksheet\Worksheet $phpExcelWorksheet = \null)
    {
        $this->phpExcelWorksheet = $phpExcelWorksheet;

        return $this;
    }

    /**
     * Get AutoFilter Range
     *
     * @return string
     */
    public function getRange()
    {
        return $this->range;
    }

    /**
     *    Set AutoFilter Range
     *
     *    @param    string        $pRange        Cell range (i.e. A1:E10)
     *    @throws    \PhpOffice\PhpSpreadsheet\Exception
     *    @return \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter
     */
    public function setRange($pRange = '')
    {
        // Uppercase coordinate
        $cellAddress = \explode('!', \strtoupper($pRange));
        if (\count($cellAddress) > 1) {
            list($worksheet, $pRange) = $cellAddress;
        }

        if (\strpos($pRange, ':') !== \false) {
            $this->range = $pRange;
        } elseif (empty($pRange)) {
            $this->range = '';
        } else {
            throw new \PhpOffice\PhpSpreadsheet\Exception('Autofilter must be set on a range of cells.');
        }

        if (empty($pRange)) {
            //    Discard all column rules
            $this->columns = array();
        } else {
            //    Discard any column rules that are no longer valid within this range
            list($rangeStart, $rangeEnd) = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::rangeBoundaries($this->range);
            foreach (array_keys($this->columns) as $key) {
                $colIndex = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::columnIndexFromString($key);
                if (($rangeStart[0] > $colIndex) || ($rangeEnd[0] < $colIndex)) {
                    unset($this->columns[$key]);
                }
            }
        }

        return $this;
    }

    /**
     * Get all AutoFilter Columns
     *
     * @throws    \PhpOffice\PhpSpreadsheet\Exception
     * @return array of PHPExcel_Worksheet_AutoFilter_Column
     */
    public function getColumns()
    {
        return $this->columns;
    }

    /**
     * Validate that the specified column is in the AutoFilter range
     *
     * @param    string    $column            Column name (e.g. A)
     * @throws    \PhpOffice\PhpSpreadsheet\Exception
     * @return    integer    The column offset within the autofilter range
     */
    public function testColumnInRange($column)
    {
        if (empty($this->range)) {
            throw new \PhpOffice\PhpSpreadsheet\Exception("No autofilter range is defined.");
        }

        $columnIndex = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::columnIndexFromString($column);
        list($rangeStart, $rangeEnd) = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::rangeBoundaries($this->range);
        if (($rangeStart[0] > $columnIndex) || ($rangeEnd[0] < $columnIndex)) {
            throw new \PhpOffice\PhpSpreadsheet\Exception("Column is outside of current autofilter range.");
        }

        return $columnIndex - $rangeStart[0];
    }

    /**
     * Get a specified AutoFilter Column Offset within the defined AutoFilter range
     *
     * @param    string    $pColumn        Column name (e.g. A)
     * @throws    \PhpOffice\PhpSpreadsheet\Exception
     * @return integer    The offset of the specified column within the autofilter range
     */
    public function getColumnOffset($pColumn)
    {
        return $this->testColumnInRange($pColumn);
    }

    /**
     * Get a specified AutoFilter Column
     *
     * @param    string    $pColumn        Column name (e.g. A)
     * @throws    \PhpOffice\PhpSpreadsheet\Exception
     * @return \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column
     */
    public function getColumn($pColumn)
    {
        $this->testColumnInRange($pColumn);

        if (!isset($this->columns[$pColumn])) {
            $this->columns[$pColumn] = new \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column($pColumn, $this);
        }

        return $this->columns[$pColumn];
    }

    /**
     * Get a specified AutoFilter Column by it's offset
     *
     * @param    integer    $pColumnOffset        Column offset within range (starting from 0)
     * @throws    \PhpOffice\PhpSpreadsheet\Exception
     * @return \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column
     */
    public function getColumnByOffset($pColumnOffset = 0)
    {
        list($rangeStart, $rangeEnd) = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::rangeBoundaries($this->range);
        $pColumn = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::stringFromColumnIndex($rangeStart[0] + $pColumnOffset - 1);

        return $this->getColumn($pColumn);
    }

    /**
     *    Set AutoFilter
     *
     *    @param    \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column|string        $pColumn
     *            A simple string containing a Column ID like 'A' is permitted
     *    @throws    \PhpOffice\PhpSpreadsheet\Exception
     *    @return \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter
     */
    public function setColumn($pColumn)
    {
        if ((\is_string($pColumn)) && (!empty($pColumn))) {
            $column = $pColumn;
        } elseif (\is_object($pColumn) && ($pColumn instanceof \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column)) {
            $column = $pColumn->getColumnIndex();
        } else {
            throw new \PhpOffice\PhpSpreadsheet\Exception("Column is not within the autofilter range.");
        }
        $this->testColumnInRange($column);

        if (\is_string($pColumn)) {
            $this->columns[$pColumn] = new \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column($pColumn, $this);
        } elseif (\is_object($pColumn) && ($pColumn instanceof \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column)) {
            $pColumn->setParent($this);
            $this->columns[$column] = $pColumn;
        }
        \ksort($this->columns);

        return $this;
    }

    /**
     * Clear a specified AutoFilter Column
     *
     * @param    string  $pColumn    Column name (e.g. A)
     * @throws    \PhpOffice\PhpSpreadsheet\Exception
     * @return \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter
     */
    public function clearColumn($pColumn)
    {
        $this->testColumnInRange($pColumn);

        if (isset($this->columns[$pColumn])) {
            unset($this->columns[$pColumn]);
        }

        return $this;
    }

    /**
     *    Shift an AutoFilter Column Rule to a different column
     *
     *    Note: This method bypasses validation of the destination column to ensure it is within this AutoFilter range.
     *        Nor does it verify whether any column rule already exists at $toColumn, but will simply overrideany existing value.
     *        Use with caution.
     *
     *    @param    string    $fromColumn        Column name (e.g. A)
     *    @param    string    $toColumn        Column name (e.g. B)
     *    @return \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter
     */
    public function shiftColumn($fromColumn = \null, $toColumn = \null)
    {
        $fromColumn = \strtoupper($fromColumn);
        $toColumn = \strtoupper($toColumn);

        if (($fromColumn !== \null) && (isset($this->columns[$fromColumn])) && ($toColumn !== \null)) {
            $this->columns[$fromColumn]->setParent();
            $this->columns[$fromColumn]->setColumnIndex($toColumn);
            $this->columns[$toColumn] = $this->columns[$fromColumn];
            $this->columns[$toColumn]->setParent($this);
            unset($this->columns[$fromColumn]);

            \ksort($this->columns);
        }

        return $this;
    }

    /**
     *    Search/Replace arrays to convert Excel wildcard syntax to a regexp syntax for preg_matching
     *
     *    @var    array
     */
    private static $fromReplace = array('\*', '\?', '~~', '~.*', '~.?');
    private static $toReplace   = array('.*', '.',  '~',  '\*',  '\?');


    /**
     *    Convert a dynamic rule daterange to a custom filter range expression for ease of calculation
     *
     *    @param    string                                        $dynamicRuleType
     *    @param    PHPExcel_Worksheet_AutoFilter_Column        &$filterColumn
     *    @return mixed[]
     */
    private function dynamicFilterDateRange($dynamicRuleType, &$filterColumn)
    {
        $returnDateType = \PhpOffice\PhpSpreadsheet\Calculation\Functions::getReturnDateType();
        \PhpOffice\PhpSpreadsheet\Calculation\Functions::setReturnDateType(\PhpOffice\PhpSpreadsheet\Calculation\Functions::RETURNDATE_PHP_NUMERIC);
        $val = $maxVal = \null;

        $ruleValues = array();
        $baseDate = \PhpOffice\PhpSpreadsheet\Calculation\DateTime::DATENOW();
        //    Calculate start/end dates for the required date range based on current date
        switch ($dynamicRuleType) {
            case \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column\Rule::AUTOFILTER_RULETYPE_DYNAMIC_LASTWEEK:
            case \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column\Rule::AUTOFILTER_RULETYPE_DYNAMIC_NEXTWEEK:
                $baseDate = \strtotime('-7 days', $baseDate);
                break;
            case \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column\Rule::AUTOFILTER_RULETYPE_DYNAMIC_LASTMONTH:
                $baseDate = \strtotime('-1 month', \gmmktime(0, 0, 0, 1, \date('m', $baseDate), \date('Y', $baseDate)));
                break;
            case \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column\Rule::AUTOFILTER_RULETYPE_DYNAMIC_NEXTMONTH:
                $baseDate = \strtotime('+1 month', \gmmktime(0, 0, 0, 1, \date('m', $baseDate), \date('Y', $baseDate)));
                break;
            case \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column\Rule::AUTOFILTER_RULETYPE_DYNAMIC_LASTQUARTER:
                $baseDate = \strtotime('-3 month', \gmmktime(0, 0, 0, 1, \date('m', $baseDate), \date('Y', $baseDate)));
                break;
            case \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column\Rule::AUTOFILTER_RULETYPE_DYNAMIC_NEXTQUARTER:
                $baseDate = \strtotime('+3 month', \gmmktime(0, 0, 0, 1, \date('m', $baseDate), \date('Y', $baseDate)));
                break;
            case \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column\Rule::AUTOFILTER_RULETYPE_DYNAMIC_LASTYEAR:
                $baseDate = \strtotime('-1 year', \gmmktime(0, 0, 0, 1, \date('m', $baseDate), \date('Y', $baseDate)));
                break;
            case \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column\Rule::AUTOFILTER_RULETYPE_DYNAMIC_NEXTYEAR:
                $baseDate = \strtotime('+1 year', \gmmktime(0, 0, 0, 1, \date('m', $baseDate), \date('Y', $baseDate)));
                break;
        }

        switch ($dynamicRuleType) {
            case \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column\Rule::AUTOFILTER_RULETYPE_DYNAMIC_TODAY:
            case \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column\Rule::AUTOFILTER_RULETYPE_DYNAMIC_YESTERDAY:
            case \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column\Rule::AUTOFILTER_RULETYPE_DYNAMIC_TOMORROW:
                $maxVal = (int) \PhpOffice\PhpSpreadsheet\Shared\Date::PHPtoExcel(\strtotime('+1 day', $baseDate));
                $val = (int) \PhpOffice\PhpSpreadsheet\Shared\Date::PHPToExcel($baseDate);
                break;
            case \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column\Rule::AUTOFILTER_RULETYPE_DYNAMIC_YEARTODATE:
                $maxVal = (int) \PhpOffice\PhpSpreadsheet\Shared\Date::PHPtoExcel(\strtotime('+1 day', $baseDate));
                $val = (int) \PhpOffice\PhpSpreadsheet\Shared\Date::PHPToExcel(\gmmktime(0, 0, 0, 1, 1, \date('Y', $baseDate)));
                break;
            case \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column\Rule::AUTOFILTER_RULETYPE_DYNAMIC_THISYEAR:
            case \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column\Rule::AUTOFILTER_RULETYPE_DYNAMIC_LASTYEAR:
            case \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column\Rule::AUTOFILTER_RULETYPE_DYNAMIC_NEXTYEAR:
                $maxVal = (int) \PhpOffice\PhpSpreadsheet\Shared\Date::PHPToExcel(\gmmktime(0, 0, 0, 31, 12, \date('Y', $baseDate)));
                ++$maxVal;
                $val = (int) \PhpOffice\PhpSpreadsheet\Shared\Date::PHPToExcel(\gmmktime(0, 0, 0, 1, 1, \date('Y', $baseDate)));
                break;
            case \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column\Rule::AUTOFILTER_RULETYPE_DYNAMIC_THISQUARTER:
            case \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column\Rule::AUTOFILTER_RULETYPE_DYNAMIC_LASTQUARTER:
            case \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column\Rule::AUTOFILTER_RULETYPE_DYNAMIC_NEXTQUARTER:
                $thisMonth = \date('m', $baseDate);
                $thisQuarter = \floor(--$thisMonth / 3);
                $maxVal = (int) \PhpOffice\PhpSpreadsheet\Shared\Date::PHPtoExcel(\gmmktime(0, 0, 0, \date('t', $baseDate), (1+$thisQuarter)*3, \date('Y', $baseDate)));
                ++$maxVal;
                $val = (int) \PhpOffice\PhpSpreadsheet\Shared\Date::PHPToExcel(\gmmktime(0, 0, 0, 1, 1+$thisQuarter*3, \date('Y', $baseDate)));
                break;
            case \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column\Rule::AUTOFILTER_RULETYPE_DYNAMIC_THISMONTH:
            case \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column\Rule::AUTOFILTER_RULETYPE_DYNAMIC_LASTMONTH:
            case \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column\Rule::AUTOFILTER_RULETYPE_DYNAMIC_NEXTMONTH:
                $maxVal = (int) \PhpOffice\PhpSpreadsheet\Shared\Date::PHPtoExcel(\gmmktime(0, 0, 0, \date('t', $baseDate), \date('m', $baseDate), \date('Y', $baseDate)));
                ++$maxVal;
                $val = (int) \PhpOffice\PhpSpreadsheet\Shared\Date::PHPToExcel(\gmmktime(0, 0, 0, 1, \date('m', $baseDate), \date('Y', $baseDate)));
                break;
            case \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column\Rule::AUTOFILTER_RULETYPE_DYNAMIC_THISWEEK:
            case \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column\Rule::AUTOFILTER_RULETYPE_DYNAMIC_LASTWEEK:
            case \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column\Rule::AUTOFILTER_RULETYPE_DYNAMIC_NEXTWEEK:
                $dayOfWeek = \date('w', $baseDate);
                $val = (int) \PhpOffice\PhpSpreadsheet\Shared\Date::PHPToExcel($baseDate) - $dayOfWeek;
                $maxVal = $val + 7;
                break;
        }

        switch ($dynamicRuleType) {
            //    Adjust Today dates for Yesterday and Tomorrow
            case \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column\Rule::AUTOFILTER_RULETYPE_DYNAMIC_YESTERDAY:
                --$maxVal;
                --$val;
                break;
            case \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column\Rule::AUTOFILTER_RULETYPE_DYNAMIC_TOMORROW:
                ++$maxVal;
                ++$val;
                break;
        }

        //    Set the filter column rule attributes ready for writing
        $filterColumn->setAttributes(array('val' => $val, 'maxVal' => $maxVal));

        //    Set the rules for identifying rows for hide/show
        $ruleValues[] = array('operator' => \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column\Rule::AUTOFILTER_COLUMN_RULE_GREATERTHANOREQUAL, 'value' => $val);
        $ruleValues[] = array('operator' => \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column\Rule::AUTOFILTER_COLUMN_RULE_LESSTHAN, 'value' => $maxVal);
        \PhpOffice\PhpSpreadsheet\Calculation\Functions::setReturnDateType($returnDateType);

        return array('method' => 'filterTestInCustomDataSet', 'arguments' => array('filterRules' => $ruleValues, 'join' => \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column::AUTOFILTER_COLUMN_JOIN_AND));
    }

    private function calculateTopTenValue($columnID, $startRow, $endRow, $ruleType, $ruleValue)
    {
        $range = $columnID.$startRow.':'.$columnID.$endRow;
        $dataValues = \PhpOffice\PhpSpreadsheet\Calculation\Functions::flattenArray($this->phpExcelWorksheet->rangeToArray($range, \null, \true, \false, false));

        $dataValues = \array_filter($dataValues);
        if ($ruleType == \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column\Rule::AUTOFILTER_COLUMN_RULE_TOPTEN_TOP) {
            \rsort($dataValues);
        } else {
            \sort($dataValues);
        }

        return \array_pop(\array_slice($dataValues, 0, $ruleValue));
    }

    /**
     *    Apply the AutoFilter rules to the AutoFilter Range
     *
     *    @throws    \PhpOffice\PhpSpreadsheet\Exception
     *    @return \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter
     */
    public function showHideRows()
    {
        list($rangeStart, $rangeEnd) = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::rangeBoundaries($this->range);

        //    The heading row should always be visible
//        echo 'AutoFilter Heading Row ', $rangeStart[1],' is always SHOWN',PHP_EOL;
        $this->phpExcelWorksheet->getRowDimension($rangeStart[1], true)->setVisible(\true);

        $columnFilterTests = array();
        foreach ($this->columns as $columnID => $filterColumn) {
            $rules = $filterColumn->getRules();
            switch ($filterColumn->getFilterType()) {
                case \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column::AUTOFILTER_FILTERTYPE_FILTER:
                    $ruleValues = array();
                    //    Build a list of the filter value selections
                    foreach ($rules as $rule) {
                        $ruleType = $rule->getRuleType();
                        $ruleValues[] = $rule->getValue();
                    }
                    //    Test if we want to include blanks in our filter criteria
                    $blanks = \false;
                    $ruleDataSet = \array_filter($ruleValues);
                    if (\count($ruleValues) !== \count($ruleDataSet)) {
                        $blanks = \true;
                    }
                    if ($ruleType == \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column\Rule::AUTOFILTER_RULETYPE_FILTER) {
                        //    Filter on absolute values
                        $columnFilterTests[$columnID] = array(
                            'method' => 'filterTestInSimpleDataSet',
                            'arguments' => array('filterValues' => $ruleDataSet, 'blanks' => $blanks)
                        );
                    } else {
                        //    Filter on date group values
                        $arguments = array(
                            'date' => array(),
                            'time' => array(),
                            'dateTime' => array(),
                        );
                        foreach ($ruleDataSet as $singleRuleDataSet) {
                            $date = $time = '';
                            if ((isset($singleRuleDataSet[\PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column\Rule::AUTOFILTER_RULETYPE_DATEGROUP_YEAR])) &&
                                ($singleRuleDataSet[\PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column\Rule::AUTOFILTER_RULETYPE_DATEGROUP_YEAR] !== '')) {
                                $date .= \sprintf('%04d', $singleRuleDataSet[\PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column\Rule::AUTOFILTER_RULETYPE_DATEGROUP_YEAR]);
                            }
                            if ((isset($singleRuleDataSet[\PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column\Rule::AUTOFILTER_RULETYPE_DATEGROUP_MONTH])) &&
                                ($singleRuleDataSet[\PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column\Rule::AUTOFILTER_RULETYPE_DATEGROUP_MONTH] != '')) {
                                $date .= \sprintf('%02d', $singleRuleDataSet[\PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column\Rule::AUTOFILTER_RULETYPE_DATEGROUP_MONTH]);
                            }
                            if ((isset($singleRuleDataSet[\PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column\Rule::AUTOFILTER_RULETYPE_DATEGROUP_DAY])) &&
                                ($singleRuleDataSet[\PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column\Rule::AUTOFILTER_RULETYPE_DATEGROUP_DAY] !== '')) {
                                $date .= \sprintf('%02d', $singleRuleDataSet[\PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column\Rule::AUTOFILTER_RULETYPE_DATEGROUP_DAY]);
                            }
                            if ((isset($singleRuleDataSet[\PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column\Rule::AUTOFILTER_RULETYPE_DATEGROUP_HOUR])) &&
                                ($singleRuleDataSet[\PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column\Rule::AUTOFILTER_RULETYPE_DATEGROUP_HOUR] !== '')) {
                                $time .= \sprintf('%02d', $singleRuleDataSet[\PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column\Rule::AUTOFILTER_RULETYPE_DATEGROUP_HOUR]);
                            }
                            if ((isset($singleRuleDataSet[\PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column\Rule::AUTOFILTER_RULETYPE_DATEGROUP_MINUTE])) &&
                                ($singleRuleDataSet[\PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column\Rule::AUTOFILTER_RULETYPE_DATEGROUP_MINUTE] !== '')) {
                                $time .= \sprintf('%02d', $singleRuleDataSet[\PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column\Rule::AUTOFILTER_RULETYPE_DATEGROUP_MINUTE]);
                            }
                            if ((isset($singleRuleDataSet[\PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column\Rule::AUTOFILTER_RULETYPE_DATEGROUP_SECOND])) &&
                                ($singleRuleDataSet[\PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column\Rule::AUTOFILTER_RULETYPE_DATEGROUP_SECOND] !== '')) {
                                $time .= \sprintf('%02d', $singleRuleDataSet[\PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column\Rule::AUTOFILTER_RULETYPE_DATEGROUP_SECOND]);
                            }
                            $dateTime = $date . $time;
                            $arguments['date'][] = $date;
                            $arguments['time'][] = $time;
                            $arguments['dateTime'][] = $dateTime;
                        }
                        //    Remove empty elements
                        $arguments['date'] = \array_filter($arguments['date']);
                        $arguments['time'] = \array_filter($arguments['time']);
                        $arguments['dateTime'] = \array_filter($arguments['dateTime']);
                        $columnFilterTests[$columnID] = array(
                            'method' => 'filterTestInDateGroupSet',
                            'arguments' => array('filterValues' => $arguments, 'blanks' => $blanks)
                        );
                    }
                    break;
                case \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column::AUTOFILTER_FILTERTYPE_CUSTOMFILTER:
                    $customRuleForBlanks = \false;
                    $ruleValues = array();
                    //    Build a list of the filter value selections
                    foreach ($rules as $rule) {
                        $ruleType = $rule->getRuleType();
                        $ruleValue = $rule->getValue();
                        if (!\is_numeric($ruleValue)) {
                            //    Convert to a regexp allowing for regexp reserved characters, wildcards and escaped wildcards
                            $ruleValue = \preg_quote($ruleValue);
                            $ruleValue = \str_replace(self::$fromReplace, self::$toReplace, $ruleValue);
                            if (\trim($ruleValue) == '') {
                                $customRuleForBlanks = \true;
                                $ruleValue = \trim($ruleValue);
                            }
                        }
                        $ruleValues[] = array('operator' => $rule->getOperator(), 'value' => $ruleValue);
                    }
                    $join = $filterColumn->getJoin();
                    $columnFilterTests[$columnID] = array(
                        'method' => 'filterTestInCustomDataSet',
                        'arguments' => array('filterRules' => $ruleValues, 'join' => $join, 'customRuleForBlanks' => $customRuleForBlanks)
                    );
                    break;
                case \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column::AUTOFILTER_FILTERTYPE_DYNAMICFILTER:
                    $ruleValues = array();
                    foreach ($rules as $rule) {
                        //    We should only ever have one Dynamic Filter Rule anyway
                        $dynamicRuleType = $rule->getGrouping();
                        if (($dynamicRuleType == \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column\Rule::AUTOFILTER_RULETYPE_DYNAMIC_ABOVEAVERAGE) ||
                            ($dynamicRuleType == \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column\Rule::AUTOFILTER_RULETYPE_DYNAMIC_BELOWAVERAGE)) {
                            //    Number (Average) based
                            //    Calculate the average
                            $averageFormula = '=AVERAGE('.$columnID.($rangeStart[1]+1).':'.$columnID.$rangeEnd[1].')';
                            $average = \PhpOffice\PhpSpreadsheet\Calculation\Calculation::getInstance()->calculateFormula($averageFormula, \null, $this->phpExcelWorksheet->getCell('A1', true));
                            //    Set above/below rule based on greaterThan or LessTan
                            $operator = ($dynamicRuleType === \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column\Rule::AUTOFILTER_RULETYPE_DYNAMIC_ABOVEAVERAGE)
                                ? \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column\Rule::AUTOFILTER_COLUMN_RULE_GREATERTHAN
                                : \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column\Rule::AUTOFILTER_COLUMN_RULE_LESSTHAN;
                            $ruleValues[] = array('operator' => $operator,
                                                   'value' => $average
                                                 );
                            $columnFilterTests[$columnID] = array(
                                'method' => 'filterTestInCustomDataSet',
                                'arguments' => array('filterRules' => $ruleValues, 'join' => \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column::AUTOFILTER_COLUMN_JOIN_OR)
                            );
                        } elseif ($dynamicRuleType{0} == 'M' || $dynamicRuleType{0} == 'Q') {
                            //    Month or Quarter
                            \sscanf($dynamicRuleType, '%[A-Z]%d', $periodType, $period);
                            if ($periodType == 'M') {
                                $ruleValues = array($period);
                            } else {
                                --$period;
                                $periodEnd = (1+$period)*3;
                                $periodStart = 1+$period*3;
                                $ruleValues = \range($periodStart, $periodEnd);
                            }
                            $columnFilterTests[$columnID] = array(
                                'method' => 'filterTestInPeriodDateSet',
                                'arguments' => $ruleValues
                            );
                            $filterColumn->setAttributes(array());
                        } else {
                            //    Date Range
                            $columnFilterTests[$columnID] = $this->dynamicFilterDateRange($dynamicRuleType, $filterColumn);
                            break;
                        }
                    }
                    break;
                case \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column::AUTOFILTER_FILTERTYPE_TOPTENFILTER:
                    $ruleValues = array();
                    $dataRowCount = $rangeEnd[1] - $rangeStart[1];
                    foreach ($rules as $rule) {
                        //    We should only ever have one Dynamic Filter Rule anyway
                        $toptenRuleType = $rule->getGrouping();
                        $ruleValue = $rule->getValue();
                        $ruleOperator = $rule->getOperator();
                    }
                    if ($ruleOperator === \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column\Rule::AUTOFILTER_COLUMN_RULE_TOPTEN_PERCENT) {
                        $ruleValue = \floor($ruleValue * ($dataRowCount / 100));
                    }
                    if ($ruleValue < 1) {
                        $ruleValue = 1;
                    }
                    if ($ruleValue > 500) {
                        $ruleValue = 500;
                    }

                    $maxVal = $this->calculateTopTenValue($columnID, $rangeStart[1]+1, $rangeEnd[1], $toptenRuleType, $ruleValue);

                    $operator = ($toptenRuleType == \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column\Rule::AUTOFILTER_COLUMN_RULE_TOPTEN_TOP)
                        ? \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column\Rule::AUTOFILTER_COLUMN_RULE_GREATERTHANOREQUAL
                        : \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column\Rule::AUTOFILTER_COLUMN_RULE_LESSTHANOREQUAL;
                    $ruleValues[] = array('operator' => $operator, 'value' => $maxVal);
                    $columnFilterTests[$columnID] = array(
                        'method' => 'filterTestInCustomDataSet',
                        'arguments' => array('filterRules' => $ruleValues, 'join' => \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column::AUTOFILTER_COLUMN_JOIN_OR)
                    );
                    $filterColumn->setAttributes(array('maxVal' => $maxVal));
                    break;
            }
        }

//        echo 'Column Filter Test CRITERIA',PHP_EOL;
//        var_dump($columnFilterTests);
//
        //    Execute the column tests for each row in the autoFilter range to determine show/hide,
        for ($row = $rangeStart[1]+1; $row <= $rangeEnd[1]; ++$row) {
//            echo 'Testing Row = ', $row,PHP_EOL;
            $result = \true;
            foreach ($columnFilterTests as $columnID => $columnFilterTest) {
//                echo 'Testing cell ', $columnID.$row,PHP_EOL;
                $cellValue = $this->phpExcelWorksheet->getCell($columnID.$row, true)->getCalculatedValue();
//                echo 'Value is ', $cellValue,PHP_EOL;
                //    Execute the filter test
                $result = $result &&
                    \call_user_func_array(
                        array('PHPExcel_Worksheet_AutoFilter', $columnFilterTest['method']),
                        array($cellValue, $columnFilterTest['arguments'])
                    );
//                echo (($result) ? 'VALID' : 'INVALID'),PHP_EOL;
                //    If filter test has resulted in FALSE, exit the loop straightaway rather than running any more tests
                if (!$result) {
                    break;
                }
            }
            //    Set show/hide for the row based on the result of the autoFilter result
//            echo (($result) ? 'SHOW' : 'HIDE'),PHP_EOL;
            $this->phpExcelWorksheet->getRowDimension($row, true)->setVisible($result);
        }

        return $this;
    }


    /**
     * Implement PHP __clone to create a deep clone, not just a shallow copy.
     */
    public function __clone()
    {
        $vars = \get_object_vars($this);
        foreach ($vars as $key => $value) {
            if (\is_object($value)) {
                $this->{$key} = $key == 'workSheet' ? \null : clone $value;
            } elseif ((\is_array($value)) && ($key == 'columns')) {
                //    The columns array of PHPExcel_Worksheet_AutoFilter objects
                $this->{$key} = array();
                foreach ($value as $k => $v) {
                    $this->{$key}[$k] = clone $v;
                    // attach the new cloned Column to this new cloned Autofilter object
                    $this->{$key}[$k]->setParent($this);
                }
            } else {
                $this->{$key} = $value;
            }
        }
    }

    /**
     * toString method replicates previous behavior by returning the range if object is
     *    referenced as a property of its parent.
     */
    public function __toString()
    {
        return $this->range;
    }
}

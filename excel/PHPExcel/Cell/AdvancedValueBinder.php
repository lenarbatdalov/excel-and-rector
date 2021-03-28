<?php

/** PHPExcel root directory */
if (!defined('PHPEXCEL_ROOT')) {
    /**
     * @ignore
     */
    define('PHPEXCEL_ROOT', dirname(__FILE__) . '/../../');
    require(PHPEXCEL_ROOT . 'PHPExcel/Autoloader.php');
}

namespace PhpOffice\PhpSpreadsheet\Cell;

/**
 * PHPExcel_Cell_AdvancedValueBinder
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
 * @package    PHPExcel_Cell
 * @copyright  Copyright (c) 2006 - 2015 PHPExcel (http://www.codeplex.com/PHPExcel)
 * @license    http://www.gnu.org/licenses/old-licenses/lgpl-2.1.txt    LGPL
 * @version    ##VERSION##, ##DATE##
 */
class AdvancedValueBinder extends \PhpOffice\PhpSpreadsheet\Cell\DefaultValueBinder implements \PhpOffice\PhpSpreadsheet\Cell\IValueBinder
{
    /**
     * Bind value to a cell
     *
     * @param \PhpOffice\PhpSpreadsheet\Cell\Cell  $phpExcelCell  Cell to bind value to
     * @param  mixed $value          Value to bind in cell
     * @return boolean
     */
    public function bindValue(\PhpOffice\PhpSpreadsheet\Cell\Cell $phpExcelCell, $value = \null)
    {
        // sanitize UTF-8 strings
        if (\is_string($value)) {
            $value = \PhpOffice\PhpSpreadsheet\Shared\StringHelper::SanitizeUTF8($value);
        }

        // Find out data type
        $dataType = parent::dataTypeForValue($value);

        // Style logic - strings
        if ($dataType === \PhpOffice\PhpSpreadsheet\Cell\DataType::TYPE_STRING && !$value instanceof \PhpOffice\PhpSpreadsheet\RichText\RichText) {
            //    Test for booleans using locale-setting
            if ($value == \PhpOffice\PhpSpreadsheet\Calculation\Calculation::getTRUE()) {
                $phpExcelCell->setValueExplicit(\true, \PhpOffice\PhpSpreadsheet\Cell\DataType::TYPE_BOOL);
                return \true;
            } elseif ($value == \PhpOffice\PhpSpreadsheet\Calculation\Calculation::getFALSE()) {
                $phpExcelCell->setValueExplicit(\false, \PhpOffice\PhpSpreadsheet\Cell\DataType::TYPE_BOOL);
                return \true;
            }

            // Check for number in scientific format
            if (\preg_match('/^'.\PhpOffice\PhpSpreadsheet\Calculation\Calculation::CALCULATION_REGEXP_NUMBER.'$/', $value)) {
                $phpExcelCell->setValueExplicit((float) $value, \PhpOffice\PhpSpreadsheet\Cell\DataType::TYPE_NUMERIC);
                return \true;
            }

            // Check for fraction
            if (\preg_match('/^([+-]?)\s*(\d+)\s?\/\s*(\d+)$/', $value, $matches)) {
                // Convert value to number
                $value = $matches[2] / $matches[3];
                if ($matches[1] == '-') {
                    $value = 0 - $value;
                }
                $phpExcelCell->setValueExplicit((float) $value, \PhpOffice\PhpSpreadsheet\Cell\DataType::TYPE_NUMERIC);
                // Set style
                $phpExcelCell->getWorksheet()->getStyle($phpExcelCell->getCoordinate())
                    ->getNumberFormat()->setFormatCode('??/??');
                return \true;
            } elseif (\preg_match('/^([+-]?)(\d*) +(\d*)\s?\/\s*(\d*)$/', $value, $matches)) {
                // Convert value to number
                $value = $matches[2] + ($matches[3] / $matches[4]);
                if ($matches[1] == '-') {
                    $value = 0 - $value;
                }
                $phpExcelCell->setValueExplicit((float) $value, \PhpOffice\PhpSpreadsheet\Cell\DataType::TYPE_NUMERIC);
                // Set style
                $phpExcelCell->getWorksheet()->getStyle($phpExcelCell->getCoordinate())
                    ->getNumberFormat()->setFormatCode('# ??/??');
                return \true;
            }

            // Check for percentage
            if (\preg_match('/^\-?\d*\.?\d*\s?\%$/', $value)) {
                // Convert value to number
                $value = (float) \str_replace('%', '', $value) / 100;
                $phpExcelCell->setValueExplicit($value, \PhpOffice\PhpSpreadsheet\Cell\DataType::TYPE_NUMERIC);
                // Set style
                $phpExcelCell->getWorksheet()->getStyle($phpExcelCell->getCoordinate())
                    ->getNumberFormat()->setFormatCode(\PhpOffice\PhpSpreadsheet\Style\NumberFormat::FORMAT_PERCENTAGE_00);
                return \true;
            }

            // Check for currency
            $currencyCode = \PhpOffice\PhpSpreadsheet\Shared\StringHelper::getCurrencyCode();
            $decimalSeparator = \PhpOffice\PhpSpreadsheet\Shared\StringHelper::getDecimalSeparator();
            $thousandsSeparator = \PhpOffice\PhpSpreadsheet\Shared\StringHelper::getThousandsSeparator();
            if (\preg_match('/^'.\preg_quote($currencyCode, '/').' *(\d{1,3}('.\preg_quote($thousandsSeparator, '/').'\d{3})*|(\d+))('.\preg_quote($decimalSeparator, '/').'\d{2})?$/', $value)) {
                // Convert value to number
                $value = (float) \trim(\str_replace(array($currencyCode, $thousandsSeparator, $decimalSeparator), array('', '', '.'), $value));
                $phpExcelCell->setValueExplicit($value, \PhpOffice\PhpSpreadsheet\Cell\DataType::TYPE_NUMERIC);
                // Set style
                $phpExcelCell->getWorksheet()->getStyle($phpExcelCell->getCoordinate())
                    ->getNumberFormat()->setFormatCode(
                        \str_replace('$', $currencyCode, \PhpOffice\PhpSpreadsheet\Style\NumberFormat::FORMAT_CURRENCY_USD_SIMPLE)
                    );
                return \true;
            } elseif (\preg_match('/^\$ *(\d{1,3}(\,\d{3})*|(\d+))(\.\d{2})?$/', $value)) {
                // Convert value to number
                $value = (float) \trim(\str_replace(array('$',','), '', $value));
                $phpExcelCell->setValueExplicit($value, \PhpOffice\PhpSpreadsheet\Cell\DataType::TYPE_NUMERIC);
                // Set style
                $phpExcelCell->getWorksheet()->getStyle($phpExcelCell->getCoordinate())
                    ->getNumberFormat()->setFormatCode(\PhpOffice\PhpSpreadsheet\Style\NumberFormat::FORMAT_CURRENCY_USD_SIMPLE);
                return \true;
            }

            // Check for time without seconds e.g. '9:45', '09:45'
            if (\preg_match('/^(\d|[0-1]\d|2[0-3]):[0-5]\d$/', $value)) {
                // Convert value to number
                list($h, $m) = \explode(':', $value);
                $days = $h / 24 + $m / 1440;
                $phpExcelCell->setValueExplicit($days, \PhpOffice\PhpSpreadsheet\Cell\DataType::TYPE_NUMERIC);
                // Set style
                $phpExcelCell->getWorksheet()->getStyle($phpExcelCell->getCoordinate())
                    ->getNumberFormat()->setFormatCode(\PhpOffice\PhpSpreadsheet\Style\NumberFormat::FORMAT_DATE_TIME3);
                return \true;
            }

            // Check for time with seconds '9:45:59', '09:45:59'
            if (\preg_match('/^(\d|[0-1]\d|2[0-3]):[0-5]\d:[0-5]\d$/', $value)) {
                // Convert value to number
                list($h, $m, $s) = \explode(':', $value);
                $days = $h / 24 + $m / 1440 + $s / 86400;
                // Convert value to number
                $phpExcelCell->setValueExplicit($days, \PhpOffice\PhpSpreadsheet\Cell\DataType::TYPE_NUMERIC);
                // Set style
                $phpExcelCell->getWorksheet()->getStyle($phpExcelCell->getCoordinate())
                    ->getNumberFormat()->setFormatCode(\PhpOffice\PhpSpreadsheet\Style\NumberFormat::FORMAT_DATE_TIME4);
                return \true;
            }

            // Check for datetime, e.g. '2008-12-31', '2008-12-31 15:59', '2008-12-31 15:59:10'
            if (($d = \PhpOffice\PhpSpreadsheet\Shared\Date::stringToExcel($value)) !== \false) {
                // Convert value to number
                $phpExcelCell->setValueExplicit($d, \PhpOffice\PhpSpreadsheet\Cell\DataType::TYPE_NUMERIC);
                $formatCode = \strpos($value, ':') !== \false ? 'yyyy-mm-dd h:mm' : 'yyyy-mm-dd';
                $phpExcelCell->getWorksheet()->getStyle($phpExcelCell->getCoordinate())
                    ->getNumberFormat()->setFormatCode($formatCode);
                return \true;
            }

            // Check for newline character "\n"
            if (\strpos($value, "\n") !== \false) {
                $value = \PhpOffice\PhpSpreadsheet\Shared\StringHelper::SanitizeUTF8($value);
                $phpExcelCell->setValueExplicit($value, \PhpOffice\PhpSpreadsheet\Cell\DataType::TYPE_STRING);
                // Set style
                $phpExcelCell->getWorksheet()->getStyle($phpExcelCell->getCoordinate())
                    ->getAlignment()->setWrapText(\true);
                return \true;
            }
        }

        // Not bound yet? Use parent...
        return parent::bindValue($phpExcelCell, $value);
    }
}

<?php

/** PHPExcel root directory */
if (!defined('PHPEXCEL_ROOT')) {
    /**
     * @ignore
     */
    define('PHPEXCEL_ROOT', dirname(__FILE__) . '/../../');
    require(PHPEXCEL_ROOT . 'PHPExcel/Autoloader.php');
}

namespace PhpOffice\PhpSpreadsheet\Calculation;

/**
 * PHPExcel_Calculation_TextData
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
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the GNU
 * Lesser General Public License for more details.
 *
 * You should have received a copy of the GNU Lesser General Public
 * License along with this library; if not, write to the Free Software
 * Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston, MA 02110-1301 USA
 *
 * @category    PHPExcel
 * @package        PHPExcel_Calculation
 * @copyright    Copyright (c) 2006 - 2015 PHPExcel (http://www.codeplex.com/PHPExcel)
 * @license        http://www.gnu.org/licenses/old-licenses/lgpl-2.1.txt    LGPL
 * @version        ##VERSION##, ##DATE##
 */
class TextData
{
    private static $invalidChars;

    private static function unicodeToOrd($c)
    {
        if (\ord($c{0}) >=0 && \ord($c{0}) <= 127) {
            return \ord($c{0});
        } elseif (\ord($c{0}) >= 192 && \ord($c{0}) <= 223) {
            return (\ord($c{0})-192)*64 + (\ord($c{1})-128);
        } elseif (\ord($c{0}) >= 224 && \ord($c{0}) <= 239) {
            return (\ord($c{0})-224)*4096 + (\ord($c{1})-128)*64 + (\ord($c{2})-128);
        } elseif (\ord($c{0}) >= 240 && \ord($c{0}) <= 247) {
            return (\ord($c{0})-240)*262144 + (\ord($c{1})-128)*4096 + (\ord($c{2})-128)*64 + (\ord($c{3})-128);
        } elseif (\ord($c{0}) >= 248 && \ord($c{0}) <= 251) {
            return (\ord($c{0})-248)*16777216 + (\ord($c{1})-128)*262144 + (\ord($c{2})-128)*4096 + (\ord($c{3})-128)*64 + (\ord($c{4})-128);
        } elseif (\ord($c{0}) >= 252 && \ord($c{0}) <= 253) {
            return (\ord($c{0})-252)*1073741824 + (\ord($c{1})-128)*16777216 + (\ord($c{2})-128)*262144 + (\ord($c{3})-128)*4096 + (\ord($c{4})-128)*64 + (\ord($c{5})-128);
        } elseif (\ord($c{0}) >= 254 && \ord($c{0}) <= 255) {
            // error
            return \PhpOffice\PhpSpreadsheet\Calculation\Functions::VALUE();
        }
        return 0;
    }

    /**
     * CHARACTER
     *
     * @param    string    $character    Value
     * @return    int
     */
    public static function CHARACTER($character)
    {
        $character = \PhpOffice\PhpSpreadsheet\Calculation\Functions::flattenSingleValue($character);

        if ((!\is_numeric($character)) || ($character < 0)) {
            return \PhpOffice\PhpSpreadsheet\Calculation\Functions::VALUE();
        }

        if (\function_exists('mb_convert_encoding')) {
            return \mb_convert_encoding('&#'.(int) $character.';', 'UTF-8', 'HTML-ENTITIES');
        } else {
            return \chr((int) $character);
        }
    }


    /**
     * TRIMNONPRINTABLE
     *
     * @param    mixed    $stringValue    Value to check
     * @return    string
     */
    public static function TRIMNONPRINTABLE($stringValue = '')
    {
        $stringValue    = \PhpOffice\PhpSpreadsheet\Calculation\Functions::flattenSingleValue($stringValue);

        if (\is_bool($stringValue)) {
            return ($stringValue) ? \PhpOffice\PhpSpreadsheet\Calculation\Calculation::getTRUE() : \PhpOffice\PhpSpreadsheet\Calculation\Calculation::getFALSE();
        }

        if (self::$invalidChars == \null) {
            self::$invalidChars = \range(\chr(0), \chr(31));
        }

        if (\is_string($stringValue) || \is_numeric($stringValue)) {
            return \str_replace(self::$invalidChars, '', \trim($stringValue, "\x00..\x1F"));
        }
        return \null;
    }


    /**
     * TRIMSPACES
     *
     * @param    mixed    $stringValue    Value to check
     * @return    string
     */
    public static function TRIMSPACES($stringValue = '')
    {
        $stringValue = \PhpOffice\PhpSpreadsheet\Calculation\Functions::flattenSingleValue($stringValue);
        if (\is_bool($stringValue)) {
            return ($stringValue) ? \PhpOffice\PhpSpreadsheet\Calculation\Calculation::getTRUE() : \PhpOffice\PhpSpreadsheet\Calculation\Calculation::getFALSE();
        }

        if (\is_string($stringValue) || \is_numeric($stringValue)) {
            return \trim(\preg_replace('/ +/', ' ', \trim($stringValue, ' ')), ' ');
        }
        return \null;
    }


    /**
     * ASCIICODE
     *
     * @param    string    $characters        Value
     * @return    int
     */
    public static function ASCIICODE($characters)
    {
        if (($characters === \null) || ($characters === '')) {
            return \PhpOffice\PhpSpreadsheet\Calculation\Functions::VALUE();
        }
        $characters    = \PhpOffice\PhpSpreadsheet\Calculation\Functions::flattenSingleValue($characters);
        if (\is_bool($characters)) {
            if (\PhpOffice\PhpSpreadsheet\Calculation\Functions::getCompatibilityMode() == \PhpOffice\PhpSpreadsheet\Calculation\Functions::COMPATIBILITY_OPENOFFICE) {
                $characters = (int) $characters;
            } else {
                $characters = ($characters) ? \PhpOffice\PhpSpreadsheet\Calculation\Calculation::getTRUE() : \PhpOffice\PhpSpreadsheet\Calculation\Calculation::getFALSE();
            }
        }

        $character = $characters;
        if ((\function_exists('mb_strlen')) && (\function_exists('mb_substr'))) {
            if (\mb_strlen($characters, 'UTF-8') > 1) {
                $character = \mb_substr($characters, 0, 1, 'UTF-8');
            }
            return self::unicodeToOrd($character);
        } else {
            if (\strlen($characters) > 0) {
                $character = \substr($characters, 0, 1);
            }
            return \ord($character);
        }
    }


    /**
     * CONCATENATE
     *
     * @return    string
     */
    public static function CONCATENATE()
    {
        $returnValue = '';

        // Loop through arguments
        $aArgs = \PhpOffice\PhpSpreadsheet\Calculation\Functions::flattenArray(\func_get_args());
        foreach ($aArgs as $aArg) {
            if (\is_bool($aArg)) {
                if (\PhpOffice\PhpSpreadsheet\Calculation\Functions::getCompatibilityMode() == \PhpOffice\PhpSpreadsheet\Calculation\Functions::COMPATIBILITY_OPENOFFICE) {
                    $aArg = (int) $aArg;
                } else {
                    $aArg = ($aArg) ? \PhpOffice\PhpSpreadsheet\Calculation\Calculation::getTRUE() : \PhpOffice\PhpSpreadsheet\Calculation\Calculation::getFALSE();
                }
            }
            $returnValue .= $aArg;
        }

        return $returnValue;
    }


    /**
     * DOLLAR
     *
     * This function converts a number to text using currency format, with the decimals rounded to the specified place.
     * The format used is $#,##0.00_);($#,##0.00)..
     *
     * @param    float    $value            The value to format
     * @param    int        $decimals        The number of digits to display to the right of the decimal point.
     *                                    If decimals is negative, number is rounded to the left of the decimal point.
     *                                    If you omit decimals, it is assumed to be 2
     * @return    string
     */
    public static function DOLLAR($value = 0, $decimals = 2)
    {
        $value        = \PhpOffice\PhpSpreadsheet\Calculation\Functions::flattenSingleValue($value);
        $decimals    = \is_null($decimals) ? 0 : \PhpOffice\PhpSpreadsheet\Calculation\Functions::flattenSingleValue($decimals);

        // Validate parameters
        if (!\is_numeric($value) || !\is_numeric($decimals)) {
            return \PhpOffice\PhpSpreadsheet\Calculation\Functions::NaN();
        }
        $decimals = \floor($decimals);

        $mask = '$#,##0';
        if ($decimals > 0) {
            $mask .= '.' . \str_repeat('0', $decimals);
        } else {
            $round = \pow(10, \abs($decimals));
            if ($value < 0) {
                $round = 0-$round;
            }
            $value = \PhpOffice\PhpSpreadsheet\Calculation\MathTrig::MROUND($value, $round);
        }

        return \PhpOffice\PhpSpreadsheet\Style\NumberFormat::toFormattedString($value, $mask, null);

    }


    /**
     * SEARCHSENSITIVE
     *
     * @param    string    $needle        The string to look for
     * @param    string    $haystack    The string in which to look
     * @param    int        $offset        Offset within $haystack
     * @return    string
     */
    public static function SEARCHSENSITIVE($needle, $haystack, $offset = 1)
    {
        $needle   = \PhpOffice\PhpSpreadsheet\Calculation\Functions::flattenSingleValue($needle);
        $haystack = \PhpOffice\PhpSpreadsheet\Calculation\Functions::flattenSingleValue($haystack);
        $offset   = \PhpOffice\PhpSpreadsheet\Calculation\Functions::flattenSingleValue($offset);

        if (!\is_bool($needle)) {
            if (\is_bool($haystack)) {
                $haystack = ($haystack) ? \PhpOffice\PhpSpreadsheet\Calculation\Calculation::getTRUE() : \PhpOffice\PhpSpreadsheet\Calculation\Calculation::getFALSE();
            }

            if (($offset > 0) && (\PhpOffice\PhpSpreadsheet\Shared\StringHelper::CountCharacters($haystack) > $offset)) {
                if (\PhpOffice\PhpSpreadsheet\Shared\StringHelper::CountCharacters($needle) == 0) {
                    return $offset;
                }
                if (\function_exists('mb_strpos')) {
                    $pos = \mb_strpos($haystack, $needle, --$offset, 'UTF-8');
                } else {
                    $pos = \strpos($haystack, $needle, --$offset);
                }
                if ($pos !== \false) {
                    return ++$pos;
                }
            }
        }
        return \PhpOffice\PhpSpreadsheet\Calculation\Functions::VALUE();
    }


    /**
     * SEARCHINSENSITIVE
     *
     * @param    string    $needle        The string to look for
     * @param    string    $haystack    The string in which to look
     * @param    int        $offset        Offset within $haystack
     * @return    string
     */
    public static function SEARCHINSENSITIVE($needle, $haystack, $offset = 1)
    {
        $needle   = \PhpOffice\PhpSpreadsheet\Calculation\Functions::flattenSingleValue($needle);
        $haystack = \PhpOffice\PhpSpreadsheet\Calculation\Functions::flattenSingleValue($haystack);
        $offset   = \PhpOffice\PhpSpreadsheet\Calculation\Functions::flattenSingleValue($offset);

        if (!\is_bool($needle)) {
            if (\is_bool($haystack)) {
                $haystack = ($haystack) ? \PhpOffice\PhpSpreadsheet\Calculation\Calculation::getTRUE() : \PhpOffice\PhpSpreadsheet\Calculation\Calculation::getFALSE();
            }

            if (($offset > 0) && (\PhpOffice\PhpSpreadsheet\Shared\StringHelper::CountCharacters($haystack) > $offset)) {
                if (\PhpOffice\PhpSpreadsheet\Shared\StringHelper::CountCharacters($needle) == 0) {
                    return $offset;
                }
                if (\function_exists('mb_stripos')) {
                    $pos = \mb_stripos($haystack, $needle, --$offset, 'UTF-8');
                } else {
                    $pos = \stripos($haystack, $needle, --$offset);
                }
                if ($pos !== \false) {
                    return ++$pos;
                }
            }
        }
        return \PhpOffice\PhpSpreadsheet\Calculation\Functions::VALUE();
    }


    /**
     * FIXEDFORMAT
     *
     * @param    mixed        $value    Value to check
     * @param    integer        $decimals
     * @param    boolean        $no_commas
     * @return    boolean
     */
    public static function FIXEDFORMAT($value, $decimals = 2, $no_commas = \false)
    {
        $value     = \PhpOffice\PhpSpreadsheet\Calculation\Functions::flattenSingleValue($value);
        $decimals  = \PhpOffice\PhpSpreadsheet\Calculation\Functions::flattenSingleValue($decimals);
        $no_commas = \PhpOffice\PhpSpreadsheet\Calculation\Functions::flattenSingleValue($no_commas);

        // Validate parameters
        if (!\is_numeric($value) || !\is_numeric($decimals)) {
            return \PhpOffice\PhpSpreadsheet\Calculation\Functions::NaN();
        }
        $decimals = \floor($decimals);

        $valueResult = \round($value, $decimals);
        if ($decimals < 0) {
            $decimals = 0;
        }
        if (!$no_commas) {
            $valueResult = \number_format($valueResult, $decimals);
        }

        return (string) $valueResult;
    }


    /**
     * LEFT
     *
     * @param    string    $value    Value
     * @param    int        $chars    Number of characters
     * @return    string
     */
    public static function LEFT($value = '', $chars = 1)
    {
        $value = \PhpOffice\PhpSpreadsheet\Calculation\Functions::flattenSingleValue($value);
        $chars = \PhpOffice\PhpSpreadsheet\Calculation\Functions::flattenSingleValue($chars);

        if ($chars < 0) {
            return \PhpOffice\PhpSpreadsheet\Calculation\Functions::VALUE();
        }

        if (\is_bool($value)) {
            $value = ($value) ? \PhpOffice\PhpSpreadsheet\Calculation\Calculation::getTRUE() : \PhpOffice\PhpSpreadsheet\Calculation\Calculation::getFALSE();
        }

        if (\function_exists('mb_substr')) {
            return \mb_substr($value, 0, $chars, 'UTF-8');
        } else {
            return \substr($value, 0, $chars);
        }
    }


    /**
     * MID
     *
     * @param    string    $value    Value
     * @param    int        $start    Start character
     * @param    int        $chars    Number of characters
     * @return    string
     */
    public static function MID($value = '', $start = 1, $chars = \null)
    {
        $value = \PhpOffice\PhpSpreadsheet\Calculation\Functions::flattenSingleValue($value);
        $start = \PhpOffice\PhpSpreadsheet\Calculation\Functions::flattenSingleValue($start);
        $chars = \PhpOffice\PhpSpreadsheet\Calculation\Functions::flattenSingleValue($chars);

        if (($start < 1) || ($chars < 0)) {
            return \PhpOffice\PhpSpreadsheet\Calculation\Functions::VALUE();
        }

        if (\is_bool($value)) {
            $value = ($value) ? \PhpOffice\PhpSpreadsheet\Calculation\Calculation::getTRUE() : \PhpOffice\PhpSpreadsheet\Calculation\Calculation::getFALSE();
        }

        if (\function_exists('mb_substr')) {
            return \mb_substr($value, --$start, $chars, 'UTF-8');
        } else {
            return \substr($value, --$start, $chars);
        }
    }


    /**
     * RIGHT
     *
     * @param    string    $value    Value
     * @param    int        $chars    Number of characters
     * @return    string
     */
    public static function RIGHT($value = '', $chars = 1)
    {
        $value = \PhpOffice\PhpSpreadsheet\Calculation\Functions::flattenSingleValue($value);
        $chars = \PhpOffice\PhpSpreadsheet\Calculation\Functions::flattenSingleValue($chars);

        if ($chars < 0) {
            return \PhpOffice\PhpSpreadsheet\Calculation\Functions::VALUE();
        }

        if (\is_bool($value)) {
            $value = ($value) ? \PhpOffice\PhpSpreadsheet\Calculation\Calculation::getTRUE() : \PhpOffice\PhpSpreadsheet\Calculation\Calculation::getFALSE();
        }

        if ((\function_exists('mb_substr')) && (\function_exists('mb_strlen'))) {
            return \mb_substr($value, \mb_strlen($value, 'UTF-8') - $chars, $chars, 'UTF-8');
        } else {
            return \substr($value, \strlen($value) - $chars);
        }
    }


    /**
     * STRINGLENGTH
     *
     * @param    string    $value    Value
     * @return    string
     */
    public static function STRINGLENGTH($value = '')
    {
        $value = \PhpOffice\PhpSpreadsheet\Calculation\Functions::flattenSingleValue($value);

        if (\is_bool($value)) {
            $value = ($value) ? \PhpOffice\PhpSpreadsheet\Calculation\Calculation::getTRUE() : \PhpOffice\PhpSpreadsheet\Calculation\Calculation::getFALSE();
        }

        if (\function_exists('mb_strlen')) {
            return \mb_strlen($value, 'UTF-8');
        } else {
            return \strlen($value);
        }
    }


    /**
     * LOWERCASE
     *
     * Converts a string value to upper case.
     *
     * @param    string        $mixedCaseString
     * @return    string
     */
    public static function LOWERCASE($mixedCaseString)
    {
        $mixedCaseString = \PhpOffice\PhpSpreadsheet\Calculation\Functions::flattenSingleValue($mixedCaseString);

        if (\is_bool($mixedCaseString)) {
            $mixedCaseString = ($mixedCaseString) ? \PhpOffice\PhpSpreadsheet\Calculation\Calculation::getTRUE() : \PhpOffice\PhpSpreadsheet\Calculation\Calculation::getFALSE();
        }

        return \PhpOffice\PhpSpreadsheet\Shared\StringHelper::StrToLower($mixedCaseString);
    }


    /**
     * UPPERCASE
     *
     * Converts a string value to upper case.
     *
     * @param    string        $mixedCaseString
     * @return    string
     */
    public static function UPPERCASE($mixedCaseString)
    {
        $mixedCaseString = \PhpOffice\PhpSpreadsheet\Calculation\Functions::flattenSingleValue($mixedCaseString);

        if (\is_bool($mixedCaseString)) {
            $mixedCaseString = ($mixedCaseString) ? \PhpOffice\PhpSpreadsheet\Calculation\Calculation::getTRUE() : \PhpOffice\PhpSpreadsheet\Calculation\Calculation::getFALSE();
        }

        return \PhpOffice\PhpSpreadsheet\Shared\StringHelper::StrToUpper($mixedCaseString);
    }


    /**
     * PROPERCASE
     *
     * Converts a string value to upper case.
     *
     * @param    string        $mixedCaseString
     * @return    string
     */
    public static function PROPERCASE($mixedCaseString)
    {
        $mixedCaseString = \PhpOffice\PhpSpreadsheet\Calculation\Functions::flattenSingleValue($mixedCaseString);

        if (\is_bool($mixedCaseString)) {
            $mixedCaseString = ($mixedCaseString) ? \PhpOffice\PhpSpreadsheet\Calculation\Calculation::getTRUE() : \PhpOffice\PhpSpreadsheet\Calculation\Calculation::getFALSE();
        }

        return \PhpOffice\PhpSpreadsheet\Shared\StringHelper::StrToTitle($mixedCaseString);
    }


    /**
     * REPLACE
     *
     * @param    string    $oldText    String to modify
     * @param    int        $start        Start character
     * @param    int        $chars        Number of characters
     * @param    string    $newText    String to replace in defined position
     * @return    string
     */
    public static function REPLACE($oldText = '', $start = 1, $chars = \null, $newText)
    {
        $oldText = \PhpOffice\PhpSpreadsheet\Calculation\Functions::flattenSingleValue($oldText);
        $start   = \PhpOffice\PhpSpreadsheet\Calculation\Functions::flattenSingleValue($start);
        $chars   = \PhpOffice\PhpSpreadsheet\Calculation\Functions::flattenSingleValue($chars);
        $newText = \PhpOffice\PhpSpreadsheet\Calculation\Functions::flattenSingleValue($newText);

        $left = self::LEFT($oldText, $start-1);
        $right = self::RIGHT($oldText, self::STRINGLENGTH($oldText)-($start+$chars)+1);

        return $left.$newText.$right;
    }


    /**
     * SUBSTITUTE
     *
     * @param    string    $text        Value
     * @param    string    $fromText    From Value
     * @param    string    $toText        To Value
     * @param    integer    $instance    Instance Number
     * @return    string
     */
    public static function SUBSTITUTE($text = '', $fromText = '', $toText = '', $instance = 0)
    {
        $text     = \PhpOffice\PhpSpreadsheet\Calculation\Functions::flattenSingleValue($text);
        $fromText = \PhpOffice\PhpSpreadsheet\Calculation\Functions::flattenSingleValue($fromText);
        $toText   = \PhpOffice\PhpSpreadsheet\Calculation\Functions::flattenSingleValue($toText);
        $instance = \floor(\PhpOffice\PhpSpreadsheet\Calculation\Functions::flattenSingleValue($instance));

        if ($instance == 0) {
            if (\function_exists('mb_str_replace')) {
                return \mb_str_replace($fromText, $toText, $text);
            } else {
                return \str_replace($fromText, $toText, $text);
            }
        } else {
            $pos = -1;
            while ($instance > 0) {
                if (\function_exists('mb_strpos')) {
                    $pos = \mb_strpos($text, $fromText, $pos+1, 'UTF-8');
                } else {
                    $pos = \strpos($text, $fromText, $pos+1);
                }
                if ($pos === \false) {
                    break;
                }
                --$instance;
            }
            if ($pos !== \false) {
                if (\function_exists('mb_strlen')) {
                    return self::REPLACE($text, ++$pos, \mb_strlen($fromText, 'UTF-8'), $toText);
                } else {
                    return self::REPLACE($text, ++$pos, \strlen($fromText), $toText);
                }
            }
        }

        return $text;
    }


    /**
     * RETURNSTRING
     *
     * @param    mixed    $testValue    Value to check
     * @return    boolean
     */
    public static function RETURNSTRING($testValue = '')
    {
        $testValue = \PhpOffice\PhpSpreadsheet\Calculation\Functions::flattenSingleValue($testValue);

        if (\is_string($testValue)) {
            return $testValue;
        }
        return \null;
    }


    /**
     * TEXTFORMAT
     *
     * @param    mixed    $value    Value to check
     * @param    string    $format    Format mask to use
     * @return    boolean
     */
    public static function TEXTFORMAT($value, $format)
    {
        $value  = \PhpOffice\PhpSpreadsheet\Calculation\Functions::flattenSingleValue($value);
        $format = \PhpOffice\PhpSpreadsheet\Calculation\Functions::flattenSingleValue($format);

        if ((\is_string($value)) && (!\is_numeric($value)) && \PhpOffice\PhpSpreadsheet\Shared\Date::isDateTimeFormatCode($format)) {
            $value = \PhpOffice\PhpSpreadsheet\Calculation\DateTime::DATEVALUE($value);
        }

        return \PhpOffice\PhpSpreadsheet\Style\NumberFormat::toFormattedString($value, $format, null);
    }

    /**
     * VALUE
     *
     * @param    mixed    $value    Value to check
     * @return    boolean
     */
    public static function VALUE($value = '')
    {
        $value = \PhpOffice\PhpSpreadsheet\Calculation\Functions::flattenSingleValue($value);

        if (!\is_numeric($value)) {
            $numberValue = \str_replace(
                \PhpOffice\PhpSpreadsheet\Shared\StringHelper::getThousandsSeparator(),
                '',
                \trim($value, " \t\n\r\0\x0B" . \PhpOffice\PhpSpreadsheet\Shared\StringHelper::getCurrencyCode())
            );
            if (\is_numeric($numberValue)) {
                return (float) $numberValue;
            }

            $returnDateType = \PhpOffice\PhpSpreadsheet\Calculation\Functions::getReturnDateType();
            \PhpOffice\PhpSpreadsheet\Calculation\Functions::setReturnDateType(\PhpOffice\PhpSpreadsheet\Calculation\Functions::RETURNDATE_EXCEL);

            if (\strpos($value, ':') !== \false) {
                $timeValue = \PhpOffice\PhpSpreadsheet\Calculation\DateTime::TIMEVALUE($value);
                if ($timeValue !== \PhpOffice\PhpSpreadsheet\Calculation\Functions::VALUE()) {
                    \PhpOffice\PhpSpreadsheet\Calculation\Functions::setReturnDateType($returnDateType);
                    return $timeValue;
                }
            }
            $dateValue = \PhpOffice\PhpSpreadsheet\Calculation\DateTime::DATEVALUE($value);
            if ($dateValue !== \PhpOffice\PhpSpreadsheet\Calculation\Functions::VALUE()) {
                \PhpOffice\PhpSpreadsheet\Calculation\Functions::setReturnDateType($returnDateType);
                return $dateValue;
            }
            \PhpOffice\PhpSpreadsheet\Calculation\Functions::setReturnDateType($returnDateType);

            return \PhpOffice\PhpSpreadsheet\Calculation\Functions::VALUE();
        }
        return (float) $value;
    }
}

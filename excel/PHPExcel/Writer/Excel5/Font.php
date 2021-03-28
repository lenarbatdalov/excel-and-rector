<?php

namespace PhpOffice\PhpSpreadsheet\Writer\Xls;

/**
 * PHPExcel_Writer_Excel5_Font
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
 * @package    PHPExcel_Writer_Excel5
 * @copyright  Copyright (c) 2006 - 2015 PHPExcel (http://www.codeplex.com/PHPExcel)
 * @license    http://www.gnu.org/licenses/old-licenses/lgpl-2.1.txt    LGPL
 * @version    ##VERSION##, ##DATE##
 */
class Font
{
    /**
     * Color index
     *
     * @var int
     */
    private $colorIndex;

    /**
     * Font
     *
     * @var \PhpOffice\PhpSpreadsheet\Style\Font
     */
    private $phpExcelStyleFont;

    /**
     * Constructor
     */
    public function __construct(\PhpOffice\PhpSpreadsheet\Style\Font $phpExcelStyleFont = \null)
    {
        $this->colorIndex = 0x7FFF;
        $this->phpExcelStyleFont = $phpExcelStyleFont;
    }

    /**
     * Set the color index
     *
     * @param int $colorIndex
     */
    public function setColorIndex($colorIndex)
    {
        $this->colorIndex = $colorIndex;
    }

    /**
     * Get font record data
     *
     * @return string
     */
    public function writeFont()
    {
        $font_outline = 0;
        $font_shadow = 0;

        $icv = $this->colorIndex; // Index to color palette
        if ($this->phpExcelStyleFont->getSuperScript()) {
            $sss = 1;
        } elseif ($this->phpExcelStyleFont->getSubScript()) {
            $sss = 2;
        } else {
            $sss = 0;
        }
        $bFamily = 0; // Font family
        $bCharSet = \PhpOffice\PhpSpreadsheet\Shared\Font::getCharsetFromFontName($this->phpExcelStyleFont->getName()); // Character set

        $record = 0x31;        // Record identifier
        $reserved = 0x00;    // Reserved
        $grbit = 0x00;        // Font attributes
        if ($this->phpExcelStyleFont->getItalic()) {
            $grbit |= 0x02;
        }
        if ($this->phpExcelStyleFont->getStrikethrough()) {
            $grbit |= 0x08;
        }
        if ($font_outline !== 0) {
            $grbit |= 0x10;
        }
        if ($font_shadow !== 0) {
            $grbit |= 0x20;
        }

        $data = \pack(
            "vvvvvCCCC",
            // Fontsize (in twips)
            $this->phpExcelStyleFont->getSize() * 20,
            $grbit,
            // Colour
            $icv,
            // Font weight
            self::mapBold($this->phpExcelStyleFont->getBold()),
            // Superscript/Subscript
            $sss,
            self::mapUnderline($this->phpExcelStyleFont->getUnderline()),
            $bFamily,
            $bCharSet,
            $reserved
        );
        $data .= \PhpOffice\PhpSpreadsheet\Shared\StringHelper::UTF8toBIFF8UnicodeShort($this->phpExcelStyleFont->getName());

        $length = \strlen($data);
        $header = \pack("vv", $record, $length);

        return($header . $data);
    }

    /**
     * Map to BIFF5-BIFF8 codes for bold
     *
     * @param boolean $bold
     * @return int
     */
    private static function mapBold($bold)
    {
        if ($bold) {
            return 0x2BC;  //  700 = Bold font weight
        }
        return 0x190;      //  400 = Normal font weight
    }

    /**
     * Map of BIFF2-BIFF8 codes for underline styles
     * @static    array of int
     *
     */
    private static $mapUnderline = array(
        \PhpOffice\PhpSpreadsheet\Style\Font::UNDERLINE_NONE              => 0x00,
        \PhpOffice\PhpSpreadsheet\Style\Font::UNDERLINE_SINGLE            => 0x01,
        \PhpOffice\PhpSpreadsheet\Style\Font::UNDERLINE_DOUBLE            => 0x02,
        \PhpOffice\PhpSpreadsheet\Style\Font::UNDERLINE_SINGLEACCOUNTING  => 0x21,
        \PhpOffice\PhpSpreadsheet\Style\Font::UNDERLINE_DOUBLEACCOUNTING  => 0x22,
    );

    /**
     * Map underline
     *
     * @param string
     * @return int
     */
    private static function mapUnderline($underline)
    {
        if (isset(self::$mapUnderline[$underline])) {
            return self::$mapUnderline[$underline];
        }
        return 0x00;
    }
}

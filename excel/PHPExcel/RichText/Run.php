<?php

namespace PhpOffice\PhpSpreadsheet\RichText;

/**
 * PHPExcel_RichText_Run
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
 * @package    PHPExcel_RichText
 * @copyright  Copyright (c) 2006 - 2015 PHPExcel (http://www.codeplex.com/PHPExcel)
 * @license    http://www.gnu.org/licenses/old-licenses/lgpl-2.1.txt    LGPL
 * @version    ##VERSION##, ##DATE##
 */
class Run extends \PhpOffice\PhpSpreadsheet\RichText\TextElement implements \PhpOffice\PhpSpreadsheet\RichText\ITextElement
{
    /**
     * Font
     *
     * @var \PhpOffice\PhpSpreadsheet\Style\Font
     */
    private $phpExcelStyleFont;

    /**
     * Create a new PHPExcel_RichText_Run instance
     *
     * @param     string        $pText        Text
     */
    public function __construct($pText = '')
    {
        // Initialise variables
        $this->setText($pText);
        $this->phpExcelStyleFont = new \PhpOffice\PhpSpreadsheet\Style\Font();
    }

    /**
     * Get font
     *
     * @return \PhpOffice\PhpSpreadsheet\Style\Font
     */
    public function getFont()
    {
        return $this->phpExcelStyleFont;
    }

    /**
     * Set font
     *
     * @param \PhpOffice\PhpSpreadsheet\Style\Font        $phpExcelStyleFont        Font
     * @throws     \PhpOffice\PhpSpreadsheet\Exception
     * @return \PhpOffice\PhpSpreadsheet\RichText\ITextElement
     */
    public function setFont(\PhpOffice\PhpSpreadsheet\Style\Font $phpExcelStyleFont = \null)
    {
        $this->phpExcelStyleFont = $phpExcelStyleFont;
        return $this;
    }

    /**
     * Get hash code
     *
     * @return string    Hash code
     */
    public function getHashCode()
    {
        return \md5(
            $this->getText() .
            $this->phpExcelStyleFont->getHashCode() .
            __CLASS__
        );
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

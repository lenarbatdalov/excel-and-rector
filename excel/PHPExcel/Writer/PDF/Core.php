<?php

namespace PhpOffice\PhpSpreadsheet\Writer;

/**
 *  PHPExcel_Writer_PDF_Core
 *
 *  Copyright (c) 2006 - 2015 PHPExcel
 *
 *  This library is free software; you can redistribute it and/or
 *  modify it under the terms of the GNU Lesser General Public
 *  License as published by the Free Software Foundation; either
 *  version 2.1 of the License, or (at your option) any later version.
 *
 *  This library is distributed in the hope that it will be useful,
 *  but WITHOUT ANY WARRANTY; without even the implied warranty of
 *  MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
 *  Lesser General Public License for more details.
 *
 *  You should have received a copy of the GNU Lesser General Public
 *  License along with this library; if not, write to the Free Software
 *  Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston, MA  02110-1301  USA
 *
 *  @category    PHPExcel
 *  @package     PHPExcel_Writer_PDF
 *  @copyright   Copyright (c) 2006 - 2015 PHPExcel (http://www.codeplex.com/PHPExcel)
 *  @license     http://www.gnu.org/licenses/old-licenses/lgpl-2.1.txt    LGPL
 *  @version     ##VERSION##, ##DATE##
 */
abstract class Pdf extends \PhpOffice\PhpSpreadsheet\Writer\Html
{
    /**
     * Temporary storage directory
     *
     * @var string
     */
    protected $tempDir = '';

    /**
     * Font
     *
     * @var string
     */
    protected $font = 'freesans';

    /**
     * Orientation (Over-ride)
     *
     * @var string
     */
    protected $orientation;

    /**
     * Paper size (Over-ride)
     *
     * @var int
     */
    protected $paperSize;


    /**
     * Temporary storage for Save Array Return type
     *
     * @var string
     */
    private $saveArrayReturnType;

    /**
     * Paper Sizes xRef List
     *
     * @var array
     */
    protected static $paperSizes = array(
        \PhpOffice\PhpSpreadsheet\Worksheet\PageSetup::PAPERSIZE_LETTER
            => 'LETTER',                 //    (8.5 in. by 11 in.)
        \PhpOffice\PhpSpreadsheet\Worksheet\PageSetup::PAPERSIZE_LETTER_SMALL
            => 'LETTER',                 //    (8.5 in. by 11 in.)
        \PhpOffice\PhpSpreadsheet\Worksheet\PageSetup::PAPERSIZE_TABLOID
            => array(792.00, 1224.00),   //    (11 in. by 17 in.)
        \PhpOffice\PhpSpreadsheet\Worksheet\PageSetup::PAPERSIZE_LEDGER
            => array(1224.00, 792.00),   //    (17 in. by 11 in.)
        \PhpOffice\PhpSpreadsheet\Worksheet\PageSetup::PAPERSIZE_LEGAL
            => 'LEGAL',                  //    (8.5 in. by 14 in.)
        \PhpOffice\PhpSpreadsheet\Worksheet\PageSetup::PAPERSIZE_STATEMENT
            => array(396.00, 612.00),    //    (5.5 in. by 8.5 in.)
        \PhpOffice\PhpSpreadsheet\Worksheet\PageSetup::PAPERSIZE_EXECUTIVE
            => 'EXECUTIVE',              //    (7.25 in. by 10.5 in.)
        \PhpOffice\PhpSpreadsheet\Worksheet\PageSetup::PAPERSIZE_A3
            => 'A3',                     //    (297 mm by 420 mm)
        \PhpOffice\PhpSpreadsheet\Worksheet\PageSetup::PAPERSIZE_A4
            => 'A4',                     //    (210 mm by 297 mm)
        \PhpOffice\PhpSpreadsheet\Worksheet\PageSetup::PAPERSIZE_A4_SMALL
            => 'A4',                     //    (210 mm by 297 mm)
        \PhpOffice\PhpSpreadsheet\Worksheet\PageSetup::PAPERSIZE_A5
            => 'A5',                     //    (148 mm by 210 mm)
        \PhpOffice\PhpSpreadsheet\Worksheet\PageSetup::PAPERSIZE_B4
            => 'B4',                     //    (250 mm by 353 mm)
        \PhpOffice\PhpSpreadsheet\Worksheet\PageSetup::PAPERSIZE_B5
            => 'B5',                     //    (176 mm by 250 mm)
        \PhpOffice\PhpSpreadsheet\Worksheet\PageSetup::PAPERSIZE_FOLIO
            => 'FOLIO',                  //    (8.5 in. by 13 in.)
        \PhpOffice\PhpSpreadsheet\Worksheet\PageSetup::PAPERSIZE_QUARTO
            => array(609.45, 779.53),    //    (215 mm by 275 mm)
        \PhpOffice\PhpSpreadsheet\Worksheet\PageSetup::PAPERSIZE_STANDARD_1
            => array(720.00, 1008.00),   //    (10 in. by 14 in.)
        \PhpOffice\PhpSpreadsheet\Worksheet\PageSetup::PAPERSIZE_STANDARD_2
            => array(792.00, 1224.00),   //    (11 in. by 17 in.)
        \PhpOffice\PhpSpreadsheet\Worksheet\PageSetup::PAPERSIZE_NOTE
            => 'LETTER',                 //    (8.5 in. by 11 in.)
        \PhpOffice\PhpSpreadsheet\Worksheet\PageSetup::PAPERSIZE_NO9_ENVELOPE
            => array(279.00, 639.00),    //    (3.875 in. by 8.875 in.)
        \PhpOffice\PhpSpreadsheet\Worksheet\PageSetup::PAPERSIZE_NO10_ENVELOPE
            => array(297.00, 684.00),    //    (4.125 in. by 9.5 in.)
        \PhpOffice\PhpSpreadsheet\Worksheet\PageSetup::PAPERSIZE_NO11_ENVELOPE
            => array(324.00, 747.00),    //    (4.5 in. by 10.375 in.)
        \PhpOffice\PhpSpreadsheet\Worksheet\PageSetup::PAPERSIZE_NO12_ENVELOPE
            => array(342.00, 792.00),    //    (4.75 in. by 11 in.)
        \PhpOffice\PhpSpreadsheet\Worksheet\PageSetup::PAPERSIZE_NO14_ENVELOPE
            => array(360.00, 828.00),    //    (5 in. by 11.5 in.)
        \PhpOffice\PhpSpreadsheet\Worksheet\PageSetup::PAPERSIZE_C
            => array(1224.00, 1584.00),  //    (17 in. by 22 in.)
        \PhpOffice\PhpSpreadsheet\Worksheet\PageSetup::PAPERSIZE_D
            => array(1584.00, 2448.00),  //    (22 in. by 34 in.)
        \PhpOffice\PhpSpreadsheet\Worksheet\PageSetup::PAPERSIZE_E
            => array(2448.00, 3168.00),  //    (34 in. by 44 in.)
        \PhpOffice\PhpSpreadsheet\Worksheet\PageSetup::PAPERSIZE_DL_ENVELOPE
            => array(311.81, 623.62),    //    (110 mm by 220 mm)
        \PhpOffice\PhpSpreadsheet\Worksheet\PageSetup::PAPERSIZE_C5_ENVELOPE
            => 'C5',                     //    (162 mm by 229 mm)
        \PhpOffice\PhpSpreadsheet\Worksheet\PageSetup::PAPERSIZE_C3_ENVELOPE
            => 'C3',                     //    (324 mm by 458 mm)
        \PhpOffice\PhpSpreadsheet\Worksheet\PageSetup::PAPERSIZE_C4_ENVELOPE
            => 'C4',                     //    (229 mm by 324 mm)
        \PhpOffice\PhpSpreadsheet\Worksheet\PageSetup::PAPERSIZE_C6_ENVELOPE
            => 'C6',                     //    (114 mm by 162 mm)
        \PhpOffice\PhpSpreadsheet\Worksheet\PageSetup::PAPERSIZE_C65_ENVELOPE
            => array(323.15, 649.13),    //    (114 mm by 229 mm)
        \PhpOffice\PhpSpreadsheet\Worksheet\PageSetup::PAPERSIZE_B4_ENVELOPE
            => 'B4',                     //    (250 mm by 353 mm)
        \PhpOffice\PhpSpreadsheet\Worksheet\PageSetup::PAPERSIZE_B5_ENVELOPE
            => 'B5',                     //    (176 mm by 250 mm)
        \PhpOffice\PhpSpreadsheet\Worksheet\PageSetup::PAPERSIZE_B6_ENVELOPE
            => array(498.90, 354.33),    //    (176 mm by 125 mm)
        \PhpOffice\PhpSpreadsheet\Worksheet\PageSetup::PAPERSIZE_ITALY_ENVELOPE
            => array(311.81, 651.97),    //    (110 mm by 230 mm)
        \PhpOffice\PhpSpreadsheet\Worksheet\PageSetup::PAPERSIZE_MONARCH_ENVELOPE
            => array(279.00, 540.00),    //    (3.875 in. by 7.5 in.)
        \PhpOffice\PhpSpreadsheet\Worksheet\PageSetup::PAPERSIZE_6_3_4_ENVELOPE
            => array(261.00, 468.00),    //    (3.625 in. by 6.5 in.)
        \PhpOffice\PhpSpreadsheet\Worksheet\PageSetup::PAPERSIZE_US_STANDARD_FANFOLD
            => array(1071.00, 792.00),   //    (14.875 in. by 11 in.)
        \PhpOffice\PhpSpreadsheet\Worksheet\PageSetup::PAPERSIZE_GERMAN_STANDARD_FANFOLD
            => array(612.00, 864.00),    //    (8.5 in. by 12 in.)
        \PhpOffice\PhpSpreadsheet\Worksheet\PageSetup::PAPERSIZE_GERMAN_LEGAL_FANFOLD
            => 'FOLIO',                  //    (8.5 in. by 13 in.)
        \PhpOffice\PhpSpreadsheet\Worksheet\PageSetup::PAPERSIZE_ISO_B4
            => 'B4',                     //    (250 mm by 353 mm)
        \PhpOffice\PhpSpreadsheet\Worksheet\PageSetup::PAPERSIZE_JAPANESE_DOUBLE_POSTCARD
            => array(566.93, 419.53),    //    (200 mm by 148 mm)
        \PhpOffice\PhpSpreadsheet\Worksheet\PageSetup::PAPERSIZE_STANDARD_PAPER_1
            => array(648.00, 792.00),    //    (9 in. by 11 in.)
        \PhpOffice\PhpSpreadsheet\Worksheet\PageSetup::PAPERSIZE_STANDARD_PAPER_2
            => array(720.00, 792.00),    //    (10 in. by 11 in.)
        \PhpOffice\PhpSpreadsheet\Worksheet\PageSetup::PAPERSIZE_STANDARD_PAPER_3
            => array(1080.00, 792.00),   //    (15 in. by 11 in.)
        \PhpOffice\PhpSpreadsheet\Worksheet\PageSetup::PAPERSIZE_INVITE_ENVELOPE
            => array(623.62, 623.62),    //    (220 mm by 220 mm)
        \PhpOffice\PhpSpreadsheet\Worksheet\PageSetup::PAPERSIZE_LETTER_EXTRA_PAPER
            => array(667.80, 864.00),    //    (9.275 in. by 12 in.)
        \PhpOffice\PhpSpreadsheet\Worksheet\PageSetup::PAPERSIZE_LEGAL_EXTRA_PAPER
            => array(667.80, 1080.00),   //    (9.275 in. by 15 in.)
        \PhpOffice\PhpSpreadsheet\Worksheet\PageSetup::PAPERSIZE_TABLOID_EXTRA_PAPER
            => array(841.68, 1296.00),   //    (11.69 in. by 18 in.)
        \PhpOffice\PhpSpreadsheet\Worksheet\PageSetup::PAPERSIZE_A4_EXTRA_PAPER
            => array(668.98, 912.76),    //    (236 mm by 322 mm)
        \PhpOffice\PhpSpreadsheet\Worksheet\PageSetup::PAPERSIZE_LETTER_TRANSVERSE_PAPER
            => array(595.80, 792.00),    //    (8.275 in. by 11 in.)
        \PhpOffice\PhpSpreadsheet\Worksheet\PageSetup::PAPERSIZE_A4_TRANSVERSE_PAPER
            => 'A4',                     //    (210 mm by 297 mm)
        \PhpOffice\PhpSpreadsheet\Worksheet\PageSetup::PAPERSIZE_LETTER_EXTRA_TRANSVERSE_PAPER
            => array(667.80, 864.00),    //    (9.275 in. by 12 in.)
        \PhpOffice\PhpSpreadsheet\Worksheet\PageSetup::PAPERSIZE_SUPERA_SUPERA_A4_PAPER
            => array(643.46, 1009.13),   //    (227 mm by 356 mm)
        \PhpOffice\PhpSpreadsheet\Worksheet\PageSetup::PAPERSIZE_SUPERB_SUPERB_A3_PAPER
            => array(864.57, 1380.47),   //    (305 mm by 487 mm)
        \PhpOffice\PhpSpreadsheet\Worksheet\PageSetup::PAPERSIZE_LETTER_PLUS_PAPER
            => array(612.00, 913.68),    //    (8.5 in. by 12.69 in.)
        \PhpOffice\PhpSpreadsheet\Worksheet\PageSetup::PAPERSIZE_A4_PLUS_PAPER
            => array(595.28, 935.43),    //    (210 mm by 330 mm)
        \PhpOffice\PhpSpreadsheet\Worksheet\PageSetup::PAPERSIZE_A5_TRANSVERSE_PAPER
            => 'A5',                     //    (148 mm by 210 mm)
        \PhpOffice\PhpSpreadsheet\Worksheet\PageSetup::PAPERSIZE_JIS_B5_TRANSVERSE_PAPER
            => array(515.91, 728.50),    //    (182 mm by 257 mm)
        \PhpOffice\PhpSpreadsheet\Worksheet\PageSetup::PAPERSIZE_A3_EXTRA_PAPER
            => array(912.76, 1261.42),   //    (322 mm by 445 mm)
        \PhpOffice\PhpSpreadsheet\Worksheet\PageSetup::PAPERSIZE_A5_EXTRA_PAPER
            => array(493.23, 666.14),    //    (174 mm by 235 mm)
        \PhpOffice\PhpSpreadsheet\Worksheet\PageSetup::PAPERSIZE_ISO_B5_EXTRA_PAPER
            => array(569.76, 782.36),    //    (201 mm by 276 mm)
        \PhpOffice\PhpSpreadsheet\Worksheet\PageSetup::PAPERSIZE_A2_PAPER
            => 'A2',                     //    (420 mm by 594 mm)
        \PhpOffice\PhpSpreadsheet\Worksheet\PageSetup::PAPERSIZE_A3_TRANSVERSE_PAPER
            => 'A3',                     //    (297 mm by 420 mm)
        \PhpOffice\PhpSpreadsheet\Worksheet\PageSetup::PAPERSIZE_A3_EXTRA_TRANSVERSE_PAPER
            => array(912.76, 1261.42)    //    (322 mm by 445 mm)
    );

    /**
     *  Create a new PHPExcel_Writer_PDF
     *
     *  @param     \PhpOffice\PhpSpreadsheet\Spreadsheet    $phpExcel    PHPExcel object
     */
    public function __construct(\PhpOffice\PhpSpreadsheet\Spreadsheet $phpExcel)
    {
        parent::__construct($phpExcel);
        $this->setUseInlineCss(\true);
        $this->tempDir = \PhpOffice\PhpSpreadsheet\Shared\File::sys_get_temp_dir();
    }

    /**
     *  Get Font
     *
     *  @return string
     */
    public function getFont()
    {
        return $this->font;
    }

    /**
     *  Set font. Examples:
     *      'arialunicid0-chinese-simplified'
     *      'arialunicid0-chinese-traditional'
     *      'arialunicid0-korean'
     *      'arialunicid0-japanese'
     *
     *  @param    string    $fontName
     */
    public function setFont($fontName)
    {
        $this->font = $fontName;
        return $this;
    }

    /**
     *  Get Paper Size
     *
     *  @return int
     */
    public function getPaperSize()
    {
        return $this->paperSize;
    }

    /**
     *  Set Paper Size
     *
     *  @param  string  $pValue Paper size
     *  @return \PhpOffice\PhpSpreadsheet\Writer\Pdf
     */
    public function setPaperSize($pValue = \PhpOffice\PhpSpreadsheet\Worksheet\PageSetup::PAPERSIZE_LETTER)
    {
        $this->paperSize = $pValue;
        return $this;
    }

    /**
     *  Get Orientation
     *
     *  @return string
     */
    public function getOrientation()
    {
        return $this->orientation;
    }

    /**
     *  Set Orientation
     *
     *  @param string $pValue  Page orientation
     *  @return \PhpOffice\PhpSpreadsheet\Writer\Pdf
     */
    public function setOrientation($pValue = \PhpOffice\PhpSpreadsheet\Worksheet\PageSetup::ORIENTATION_DEFAULT)
    {
        $this->orientation = $pValue;
        return $this;
    }

    /**
     *  Get temporary storage directory
     *
     *  @return string
     */
    public function getTempDir()
    {
        return $this->tempDir;
    }

    /**
     *  Set temporary storage directory
     *
     *  @param     string        $pValue        Temporary storage directory
     *  @throws    \PhpOffice\PhpSpreadsheet\Writer\Exception    when directory does not exist
     *  @return    \PhpOffice\PhpSpreadsheet\Writer\Pdf
     */
    public function setTempDir($pValue = '')
    {
        if (\is_dir($pValue)) {
            $this->tempDir = $pValue;
        } else {
            throw new \PhpOffice\PhpSpreadsheet\Writer\Exception("Directory does not exist: $pValue");
        }
        return $this;
    }

    /**
     *  Save PHPExcel to PDF file, pre-save
     *
     *  @param     string     $pFilename   Name of the file to save as
     *  @throws    \PhpOffice\PhpSpreadsheet\Writer\Exception
     */
    protected function prepareForSave($pFilename = \null)
    {
        //  garbage collect
        $this->phpExcel->garbageCollect();

        $this->saveArrayReturnType = \PhpOffice\PhpSpreadsheet\Calculation\Calculation::getArrayReturnType();
        \PhpOffice\PhpSpreadsheet\Calculation\Calculation::setArrayReturnType(\PhpOffice\PhpSpreadsheet\Calculation\Calculation::RETURN_ARRAY_AS_VALUE);

        //  Open file
        $fileHandle = \fopen($pFilename, 'w');
        if ($fileHandle === \false) {
            throw new \PhpOffice\PhpSpreadsheet\Writer\Exception("Could not open file $pFilename for writing.");
        }

        //  Set PDF
        $this->isPdf = \true;
        //  Build CSS
        $this->buildCSS(\true);

        return $fileHandle;
    }

    /**
     *  Save PHPExcel to PDF file, post-save
     *
     *  @param     resource      $fileHandle
     *  @throws    \PhpOffice\PhpSpreadsheet\Writer\Exception
     */
    protected function restoreStateAfterSave($fileHandle)
    {
        //  Close file
        \fclose($fileHandle);

        \PhpOffice\PhpSpreadsheet\Calculation\Calculation::setArrayReturnType($this->saveArrayReturnType);
    }
}

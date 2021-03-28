<?php

/**  Require tcPDF library */
$pdfRendererClassFile = \PhpOffice\PhpSpreadsheet\Settings::getPdfRendererPath() . '/tcpdf.php';
if (file_exists($pdfRendererClassFile)) {
    $k_path_url = \PhpOffice\PhpSpreadsheet\Settings::getPdfRendererPath();
    require_once $pdfRendererClassFile;
} else {
    throw new \PhpOffice\PhpSpreadsheet\Writer\Exception('Unable to load PDF Rendering library');
}

namespace PhpOffice\PhpSpreadsheet\Writer\Pdf;

/**
 *  PHPExcel_Writer_PDF_tcPDF
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
class Tcpdf extends \PhpOffice\PhpSpreadsheet\Writer\Pdf implements \PhpOffice\PhpSpreadsheet\Writer\IWriter
{
    /**
     *  Save PHPExcel to file
     *
     *  @param     string     $pFilename   Name of the file to save as
     *  @throws    \PhpOffice\PhpSpreadsheet\Writer\Exception
     */
    public function save($pFilename = \null)
    {
        $fileHandle = parent::prepareForSave($pFilename);

        //  Default PDF paper size
        $paperSize = 'LETTER';    //    Letter    (8.5 in. by 11 in.)

        //  Check for paper size and page orientation
        if (\is_null($this->getSheetIndex())) {
            $orientation = ($this->phpExcel->getSheet(0)->getPageSetup()->getOrientation()
                == \PhpOffice\PhpSpreadsheet\Worksheet\PageSetup::ORIENTATION_LANDSCAPE) ? 'L' : 'P';
            $printPaperSize = $this->phpExcel->getSheet(0)->getPageSetup()->getPaperSize();
            $printMargins = $this->phpExcel->getSheet(0)->getPageMargins();
        } else {
            $orientation = ($this->phpExcel->getSheet($this->getSheetIndex())->getPageSetup()->getOrientation()
                == \PhpOffice\PhpSpreadsheet\Worksheet\PageSetup::ORIENTATION_LANDSCAPE) ? 'L' : 'P';
            $printPaperSize = $this->phpExcel->getSheet($this->getSheetIndex())->getPageSetup()->getPaperSize();
            $printMargins = $this->phpExcel->getSheet($this->getSheetIndex())->getPageMargins();
        }

        //  Override Page Orientation
        if (!\is_null($this->getOrientation())) {
            $orientation = ($this->getOrientation() == \PhpOffice\PhpSpreadsheet\Worksheet\PageSetup::ORIENTATION_LANDSCAPE)
                ? 'L'
                : 'P';
        }
        //  Override Paper Size
        if (!\is_null($this->getPaperSize())) {
            $printPaperSize = $this->getPaperSize();
        }

        if (isset(self::$paperSizes[$printPaperSize])) {
            $paperSize = self::$paperSizes[$printPaperSize];
        }


        //  Create PDF
        $tcpdf = new \TCPDF($orientation, 'pt', $paperSize);
        $tcpdf->setFontSubsetting(\false);
        //    Set margins, converting inches to points (using 72 dpi)
        $tcpdf->SetMargins($printMargins->getLeft() * 72, $printMargins->getTop() * 72, $printMargins->getRight() * 72);
        $tcpdf->SetAutoPageBreak(\true, $printMargins->getBottom() * 72);

        $tcpdf->setPrintHeader(\false);
        $tcpdf->setPrintFooter(\false);

        $tcpdf->AddPage();

        //  Set the appropriate font
        $tcpdf->SetFont($this->getFont());
        $tcpdf->writeHTML(
            $this->generateHTMLHeader() .
            $this->generateSheetData() .
            $this->generateHTMLFooter()
        );

        //  Document info
        $tcpdf->SetTitle($this->phpExcel->getProperties()->getTitle());
        $tcpdf->SetAuthor($this->phpExcel->getProperties()->getCreator());
        $tcpdf->SetSubject($this->phpExcel->getProperties()->getSubject());
        $tcpdf->SetKeywords($this->phpExcel->getProperties()->getKeywords());
        $tcpdf->SetCreator($this->phpExcel->getProperties()->getCreator());

        //  Write to file
        \fwrite($fileHandle, $tcpdf->output($pFilename, 'S'));

        parent::restoreStateAfterSave($fileHandle);
    }
}

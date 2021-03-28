<?php

namespace PhpOffice\PhpSpreadsheet\Writer\Xlsx;

/**
 * PHPExcel_Writer_Excel2007_Drawing
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
 * @package    PHPExcel_Writer_Excel2007
 * @copyright  Copyright (c) 2006 - 2015 PHPExcel (http://www.codeplex.com/PHPExcel)
 * @license    http://www.gnu.org/licenses/old-licenses/lgpl-2.1.txt    LGPL
 * @version    ##VERSION##, ##DATE##
 */
class Drawing extends \PhpOffice\PhpSpreadsheet\Writer\Xlsx\WriterPart
{
    /**
     * Write drawings to XML format
     *
     * @param     int                    &$chartRef        Chart ID
     * @param    boolean                $includeCharts    Flag indicating if we should include drawing details for charts
     * @return     string                 XML Output
     * @throws     \PhpOffice\PhpSpreadsheet\Writer\Exception
     */
    public function writeDrawings(\PhpOffice\PhpSpreadsheet\Worksheet\Worksheet $phpExcelWorksheet = \null, &$chartRef, $includeCharts = \false)
    {
        // Create XML writer
        $objWriter = \null;
        if ($this->getParentWriter()->getUseDiskCaching()) {
            $objWriter = new \PhpOffice\PhpSpreadsheet\Shared\XMLWriter(\PhpOffice\PhpSpreadsheet\Shared\XMLWriter::STORAGE_DISK, $this->getParentWriter()->getDiskCachingDirectory());
        } else {
            $objWriter = new \PhpOffice\PhpSpreadsheet\Shared\XMLWriter(\PhpOffice\PhpSpreadsheet\Shared\XMLWriter::STORAGE_MEMORY);
        }

        // XML header
        $objWriter->startDocument('1.0', 'UTF-8', 'yes');

        // xdr:wsDr
        $objWriter->startElement('xdr:wsDr');
        $objWriter->writeAttribute('xmlns:xdr', 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing');
        $objWriter->writeAttribute('xmlns:a', 'http://schemas.openxmlformats.org/drawingml/2006/main');

        // Loop through images and write drawings
        $i = 1;
        $iterator = $phpExcelWorksheet->getDrawingCollection()->getIterator();
        while ($iterator->valid()) {
            $this->writeDrawing($objWriter, $iterator->current(), $i);

            $iterator->next();
            ++$i;
        }

        if ($includeCharts) {
            $chartCount = $phpExcelWorksheet->getChartCount();
            // Loop through charts and write the chart position
            if ($chartCount > 0) {
                for ($c = 0; $c < $chartCount; ++$c) {
                    $this->writeChart($objWriter, $phpExcelWorksheet->getChartByIndex($c), $c+$i);
                }
            }
        }

        $objWriter->endElement();

        // Return
        return $objWriter->getData();
    }

    /**
     * Write drawings to XML format
     *
     * @param \PhpOffice\PhpSpreadsheet\Shared\XMLWriter    $phpExcelSharedXMLWriter         XML Writer
     * @param     int                            $pRelationId
     * @throws     \PhpOffice\PhpSpreadsheet\Writer\Exception
     */
    public function writeChart(\PhpOffice\PhpSpreadsheet\Shared\XMLWriter $phpExcelSharedXMLWriter = \null, \PhpOffice\PhpSpreadsheet\Chart\Chart $phpExcelChart = \null, $pRelationId = -1)
    {
        $tl = $phpExcelChart->getTopLeftPosition();
        $tl['colRow'] = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::coordinateFromString($tl['cell']);
        $br = $phpExcelChart->getBottomRightPosition();
        $br['colRow'] = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::coordinateFromString($br['cell']);

        $phpExcelSharedXMLWriter->startElement('xdr:twoCellAnchor');

            $phpExcelSharedXMLWriter->startElement('xdr:from');
                $phpExcelSharedXMLWriter->writeElement('xdr:col', \PhpOffice\PhpSpreadsheet\Cell\Coordinate::columnIndexFromString($tl['colRow'][0]) - 1);
                $phpExcelSharedXMLWriter->writeElement('xdr:colOff', \PhpOffice\PhpSpreadsheet\Shared\Drawing::pixelsToEMU($tl['xOffset']));
                $phpExcelSharedXMLWriter->writeElement('xdr:row', $tl['colRow'][1] - 1);
                $phpExcelSharedXMLWriter->writeElement('xdr:rowOff', \PhpOffice\PhpSpreadsheet\Shared\Drawing::pixelsToEMU($tl['yOffset']));
            $phpExcelSharedXMLWriter->endElement();
            $phpExcelSharedXMLWriter->startElement('xdr:to');
                $phpExcelSharedXMLWriter->writeElement('xdr:col', \PhpOffice\PhpSpreadsheet\Cell\Coordinate::columnIndexFromString($br['colRow'][0]) - 1);
                $phpExcelSharedXMLWriter->writeElement('xdr:colOff', \PhpOffice\PhpSpreadsheet\Shared\Drawing::pixelsToEMU($br['xOffset']));
                $phpExcelSharedXMLWriter->writeElement('xdr:row', $br['colRow'][1] - 1);
                $phpExcelSharedXMLWriter->writeElement('xdr:rowOff', \PhpOffice\PhpSpreadsheet\Shared\Drawing::pixelsToEMU($br['yOffset']));
            $phpExcelSharedXMLWriter->endElement();

            $phpExcelSharedXMLWriter->startElement('xdr:graphicFrame');
                $phpExcelSharedXMLWriter->writeAttribute('macro', '');
                $phpExcelSharedXMLWriter->startElement('xdr:nvGraphicFramePr');
                    $phpExcelSharedXMLWriter->startElement('xdr:cNvPr');
                        $phpExcelSharedXMLWriter->writeAttribute('name', 'Chart '.$pRelationId);
                        $phpExcelSharedXMLWriter->writeAttribute('id', 1025 * $pRelationId);
                    $phpExcelSharedXMLWriter->endElement();
                    $phpExcelSharedXMLWriter->startElement('xdr:cNvGraphicFramePr');
                        $phpExcelSharedXMLWriter->startElement('a:graphicFrameLocks');
                        $phpExcelSharedXMLWriter->endElement();
                    $phpExcelSharedXMLWriter->endElement();
                $phpExcelSharedXMLWriter->endElement();

                $phpExcelSharedXMLWriter->startElement('xdr:xfrm');
                    $phpExcelSharedXMLWriter->startElement('a:off');
                        $phpExcelSharedXMLWriter->writeAttribute('x', '0');
                        $phpExcelSharedXMLWriter->writeAttribute('y', '0');
                    $phpExcelSharedXMLWriter->endElement();
                    $phpExcelSharedXMLWriter->startElement('a:ext');
                        $phpExcelSharedXMLWriter->writeAttribute('cx', '0');
                        $phpExcelSharedXMLWriter->writeAttribute('cy', '0');
                    $phpExcelSharedXMLWriter->endElement();
                $phpExcelSharedXMLWriter->endElement();

                $phpExcelSharedXMLWriter->startElement('a:graphic');
                    $phpExcelSharedXMLWriter->startElement('a:graphicData');
                        $phpExcelSharedXMLWriter->writeAttribute('uri', 'http://schemas.openxmlformats.org/drawingml/2006/chart');
                        $phpExcelSharedXMLWriter->startElement('c:chart');
                            $phpExcelSharedXMLWriter->writeAttribute('xmlns:c', 'http://schemas.openxmlformats.org/drawingml/2006/chart');
                            $phpExcelSharedXMLWriter->writeAttribute('xmlns:r', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships');
                            $phpExcelSharedXMLWriter->writeAttribute('r:id', 'rId'.$pRelationId);
                        $phpExcelSharedXMLWriter->endElement();
                    $phpExcelSharedXMLWriter->endElement();
                $phpExcelSharedXMLWriter->endElement();
            $phpExcelSharedXMLWriter->endElement();

            $phpExcelSharedXMLWriter->startElement('xdr:clientData');
            $phpExcelSharedXMLWriter->endElement();

        $phpExcelSharedXMLWriter->endElement();
    }

    /**
     * Write drawings to XML format
     *
     * @param \PhpOffice\PhpSpreadsheet\Shared\XMLWriter            $phpExcelSharedXMLWriter         XML Writer
     * @param     int                                    $pRelationId
     * @throws     \PhpOffice\PhpSpreadsheet\Writer\Exception
     */
    public function writeDrawing(\PhpOffice\PhpSpreadsheet\Shared\XMLWriter $phpExcelSharedXMLWriter = \null, \PhpOffice\PhpSpreadsheet\Worksheet\BaseDrawing $phpExcelWorksheetBaseDrawing = \null, $pRelationId = -1)
    {
        if ($pRelationId >= 0) {
            // xdr:oneCellAnchor
            $phpExcelSharedXMLWriter->startElement('xdr:oneCellAnchor');
            // Image location
            $aCoordinates         = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::coordinateFromString($phpExcelWorksheetBaseDrawing->getCoordinates());
            $aCoordinates[0]     = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::columnIndexFromString($aCoordinates[0]);

            // xdr:from
            $phpExcelSharedXMLWriter->startElement('xdr:from');
            $phpExcelSharedXMLWriter->writeElement('xdr:col', $aCoordinates[0] - 1);
            $phpExcelSharedXMLWriter->writeElement('xdr:colOff', \PhpOffice\PhpSpreadsheet\Shared\Drawing::pixelsToEMU($phpExcelWorksheetBaseDrawing->getOffsetX()));
            $phpExcelSharedXMLWriter->writeElement('xdr:row', $aCoordinates[1] - 1);
            $phpExcelSharedXMLWriter->writeElement('xdr:rowOff', \PhpOffice\PhpSpreadsheet\Shared\Drawing::pixelsToEMU($phpExcelWorksheetBaseDrawing->getOffsetY()));
            $phpExcelSharedXMLWriter->endElement();

            // xdr:ext
            $phpExcelSharedXMLWriter->startElement('xdr:ext');
            $phpExcelSharedXMLWriter->writeAttribute('cx', \PhpOffice\PhpSpreadsheet\Shared\Drawing::pixelsToEMU($phpExcelWorksheetBaseDrawing->getWidth()));
            $phpExcelSharedXMLWriter->writeAttribute('cy', \PhpOffice\PhpSpreadsheet\Shared\Drawing::pixelsToEMU($phpExcelWorksheetBaseDrawing->getHeight()));
            $phpExcelSharedXMLWriter->endElement();

            // xdr:pic
            $phpExcelSharedXMLWriter->startElement('xdr:pic');

            // xdr:nvPicPr
            $phpExcelSharedXMLWriter->startElement('xdr:nvPicPr');

            // xdr:cNvPr
            $phpExcelSharedXMLWriter->startElement('xdr:cNvPr');
            $phpExcelSharedXMLWriter->writeAttribute('id', $pRelationId);
            $phpExcelSharedXMLWriter->writeAttribute('name', $phpExcelWorksheetBaseDrawing->getName());
            $phpExcelSharedXMLWriter->writeAttribute('descr', $phpExcelWorksheetBaseDrawing->getDescription());
            $phpExcelSharedXMLWriter->endElement();

            // xdr:cNvPicPr
            $phpExcelSharedXMLWriter->startElement('xdr:cNvPicPr');

            // a:picLocks
            $phpExcelSharedXMLWriter->startElement('a:picLocks');
            $phpExcelSharedXMLWriter->writeAttribute('noChangeAspect', '1');
            $phpExcelSharedXMLWriter->endElement();

            $phpExcelSharedXMLWriter->endElement();

            $phpExcelSharedXMLWriter->endElement();

            // xdr:blipFill
            $phpExcelSharedXMLWriter->startElement('xdr:blipFill');

            // a:blip
            $phpExcelSharedXMLWriter->startElement('a:blip');
            $phpExcelSharedXMLWriter->writeAttribute('xmlns:r', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships');
            $phpExcelSharedXMLWriter->writeAttribute('r:embed', 'rId' . $pRelationId);
            $phpExcelSharedXMLWriter->endElement();

            // a:stretch
            $phpExcelSharedXMLWriter->startElement('a:stretch');
                $phpExcelSharedXMLWriter->writeElement('a:fillRect', \null);
            $phpExcelSharedXMLWriter->endElement();

            $phpExcelSharedXMLWriter->endElement();

            // xdr:spPr
            $phpExcelSharedXMLWriter->startElement('xdr:spPr');

            // a:xfrm
            $phpExcelSharedXMLWriter->startElement('a:xfrm');
            $phpExcelSharedXMLWriter->writeAttribute('rot', \PhpOffice\PhpSpreadsheet\Shared\Drawing::degreesToAngle($phpExcelWorksheetBaseDrawing->getRotation()));
            $phpExcelSharedXMLWriter->endElement();

            // a:prstGeom
            $phpExcelSharedXMLWriter->startElement('a:prstGeom');
            $phpExcelSharedXMLWriter->writeAttribute('prst', 'rect');

            // a:avLst
            $phpExcelSharedXMLWriter->writeElement('a:avLst', \null);

            $phpExcelSharedXMLWriter->endElement();

//                        // a:solidFill
//                        $objWriter->startElement('a:solidFill');

//                            // a:srgbClr
//                            $objWriter->startElement('a:srgbClr');
//                            $objWriter->writeAttribute('val', 'FFFFFF');

///* SHADE
//                                // a:shade
//                                $objWriter->startElement('a:shade');
//                                $objWriter->writeAttribute('val', '85000');
//                                $objWriter->endElement();
//*/

//                            $objWriter->endElement();

//                        $objWriter->endElement();
/*
            // a:ln
            $objWriter->startElement('a:ln');
            $objWriter->writeAttribute('w', '88900');
            $objWriter->writeAttribute('cap', 'sq');

                // a:solidFill
                $objWriter->startElement('a:solidFill');

                    // a:srgbClr
                    $objWriter->startElement('a:srgbClr');
                    $objWriter->writeAttribute('val', 'FFFFFF');
                    $objWriter->endElement();

                $objWriter->endElement();

                // a:miter
                $objWriter->startElement('a:miter');
                $objWriter->writeAttribute('lim', '800000');
                $objWriter->endElement();

            $objWriter->endElement();
*/

            if ($phpExcelWorksheetBaseDrawing->getShadow()->getVisible()) {
                // a:effectLst
                $phpExcelSharedXMLWriter->startElement('a:effectLst');

                // a:outerShdw
                $phpExcelSharedXMLWriter->startElement('a:outerShdw');
                $phpExcelSharedXMLWriter->writeAttribute('blurRad', \PhpOffice\PhpSpreadsheet\Shared\Drawing::pixelsToEMU($phpExcelWorksheetBaseDrawing->getShadow()->getBlurRadius()));
                $phpExcelSharedXMLWriter->writeAttribute('dist', \PhpOffice\PhpSpreadsheet\Shared\Drawing::pixelsToEMU($phpExcelWorksheetBaseDrawing->getShadow()->getDistance()));
                $phpExcelSharedXMLWriter->writeAttribute('dir', \PhpOffice\PhpSpreadsheet\Shared\Drawing::degreesToAngle($phpExcelWorksheetBaseDrawing->getShadow()->getDirection()));
                $phpExcelSharedXMLWriter->writeAttribute('algn', $phpExcelWorksheetBaseDrawing->getShadow()->getAlignment());
                $phpExcelSharedXMLWriter->writeAttribute('rotWithShape', '0');

                // a:srgbClr
                $phpExcelSharedXMLWriter->startElement('a:srgbClr');
                $phpExcelSharedXMLWriter->writeAttribute('val', $phpExcelWorksheetBaseDrawing->getShadow()->getColor()->getRGB());

                // a:alpha
                $phpExcelSharedXMLWriter->startElement('a:alpha');
                $phpExcelSharedXMLWriter->writeAttribute('val', $phpExcelWorksheetBaseDrawing->getShadow()->getAlpha() * 1000);
                $phpExcelSharedXMLWriter->endElement();

                $phpExcelSharedXMLWriter->endElement();

                $phpExcelSharedXMLWriter->endElement();

                $phpExcelSharedXMLWriter->endElement();
            }
/*

                // a:scene3d
                $objWriter->startElement('a:scene3d');

                    // a:camera
                    $objWriter->startElement('a:camera');
                    $objWriter->writeAttribute('prst', 'orthographicFront');
                    $objWriter->endElement();

                    // a:lightRig
                    $objWriter->startElement('a:lightRig');
                    $objWriter->writeAttribute('rig', 'twoPt');
                    $objWriter->writeAttribute('dir', 't');

                        // a:rot
                        $objWriter->startElement('a:rot');
                        $objWriter->writeAttribute('lat', '0');
                        $objWriter->writeAttribute('lon', '0');
                        $objWriter->writeAttribute('rev', '0');
                        $objWriter->endElement();

                    $objWriter->endElement();

                $objWriter->endElement();
*/
/*
                // a:sp3d
                $objWriter->startElement('a:sp3d');

                    // a:bevelT
                    $objWriter->startElement('a:bevelT');
                    $objWriter->writeAttribute('w', '25400');
                    $objWriter->writeAttribute('h', '19050');
                    $objWriter->endElement();

                    // a:contourClr
                    $objWriter->startElement('a:contourClr');

                        // a:srgbClr
                        $objWriter->startElement('a:srgbClr');
                        $objWriter->writeAttribute('val', 'FFFFFF');
                        $objWriter->endElement();

                    $objWriter->endElement();

                $objWriter->endElement();
*/
            $phpExcelSharedXMLWriter->endElement();

            $phpExcelSharedXMLWriter->endElement();

            // xdr:clientData
            $phpExcelSharedXMLWriter->writeElement('xdr:clientData', \null);

            $phpExcelSharedXMLWriter->endElement();
        } else {
            throw new \PhpOffice\PhpSpreadsheet\Writer\Exception("Invalid parameters passed.");
        }
    }

    /**
     * Write VML header/footer images to XML format
     *
     * @return     string                                 XML Output
     * @throws     \PhpOffice\PhpSpreadsheet\Writer\Exception
     */
    public function writeVMLHeaderFooterImages(\PhpOffice\PhpSpreadsheet\Worksheet\Worksheet $phpExcelWorksheet = \null)
    {
        // Create XML writer
        $objWriter = \null;
        if ($this->getParentWriter()->getUseDiskCaching()) {
            $objWriter = new \PhpOffice\PhpSpreadsheet\Shared\XMLWriter(\PhpOffice\PhpSpreadsheet\Shared\XMLWriter::STORAGE_DISK, $this->getParentWriter()->getDiskCachingDirectory());
        } else {
            $objWriter = new \PhpOffice\PhpSpreadsheet\Shared\XMLWriter(\PhpOffice\PhpSpreadsheet\Shared\XMLWriter::STORAGE_MEMORY);
        }

        // XML header
        $objWriter->startDocument('1.0', 'UTF-8', 'yes');

        // Header/footer images
        $images = $phpExcelWorksheet->getHeaderFooter()->getImages();

        // xml
        $objWriter->startElement('xml');
        $objWriter->writeAttribute('xmlns:v', 'urn:schemas-microsoft-com:vml');
        $objWriter->writeAttribute('xmlns:o', 'urn:schemas-microsoft-com:office:office');
        $objWriter->writeAttribute('xmlns:x', 'urn:schemas-microsoft-com:office:excel');

        // o:shapelayout
        $objWriter->startElement('o:shapelayout');
        $objWriter->writeAttribute('v:ext', 'edit');

        // o:idmap
        $objWriter->startElement('o:idmap');
        $objWriter->writeAttribute('v:ext', 'edit');
        $objWriter->writeAttribute('data', '1');
        $objWriter->endElement();

        $objWriter->endElement();

        // v:shapetype
        $objWriter->startElement('v:shapetype');
        $objWriter->writeAttribute('id', '_x0000_t75');
        $objWriter->writeAttribute('coordsize', '21600,21600');
        $objWriter->writeAttribute('o:spt', '75');
        $objWriter->writeAttribute('o:preferrelative', 't');
        $objWriter->writeAttribute('path', 'm@4@5l@4@11@9@11@9@5xe');
        $objWriter->writeAttribute('filled', 'f');
        $objWriter->writeAttribute('stroked', 'f');

        // v:stroke
        $objWriter->startElement('v:stroke');
        $objWriter->writeAttribute('joinstyle', 'miter');
        $objWriter->endElement();

        // v:formulas
        $objWriter->startElement('v:formulas');

        // v:f
        $objWriter->startElement('v:f');
        $objWriter->writeAttribute('eqn', 'if lineDrawn pixelLineWidth 0');
        $objWriter->endElement();

        // v:f
        $objWriter->startElement('v:f');
        $objWriter->writeAttribute('eqn', 'sum @0 1 0');
        $objWriter->endElement();

        // v:f
        $objWriter->startElement('v:f');
        $objWriter->writeAttribute('eqn', 'sum 0 0 @1');
        $objWriter->endElement();

        // v:f
        $objWriter->startElement('v:f');
        $objWriter->writeAttribute('eqn', 'prod @2 1 2');
        $objWriter->endElement();

        // v:f
        $objWriter->startElement('v:f');
        $objWriter->writeAttribute('eqn', 'prod @3 21600 pixelWidth');
        $objWriter->endElement();

        // v:f
        $objWriter->startElement('v:f');
        $objWriter->writeAttribute('eqn', 'prod @3 21600 pixelHeight');
        $objWriter->endElement();

        // v:f
        $objWriter->startElement('v:f');
        $objWriter->writeAttribute('eqn', 'sum @0 0 1');
        $objWriter->endElement();

        // v:f
        $objWriter->startElement('v:f');
        $objWriter->writeAttribute('eqn', 'prod @6 1 2');
        $objWriter->endElement();

        // v:f
        $objWriter->startElement('v:f');
        $objWriter->writeAttribute('eqn', 'prod @7 21600 pixelWidth');
        $objWriter->endElement();

        // v:f
        $objWriter->startElement('v:f');
        $objWriter->writeAttribute('eqn', 'sum @8 21600 0');
        $objWriter->endElement();

        // v:f
        $objWriter->startElement('v:f');
        $objWriter->writeAttribute('eqn', 'prod @7 21600 pixelHeight');
        $objWriter->endElement();

        // v:f
        $objWriter->startElement('v:f');
        $objWriter->writeAttribute('eqn', 'sum @10 21600 0');
        $objWriter->endElement();

        $objWriter->endElement();

        // v:path
        $objWriter->startElement('v:path');
        $objWriter->writeAttribute('o:extrusionok', 'f');
        $objWriter->writeAttribute('gradientshapeok', 't');
        $objWriter->writeAttribute('o:connecttype', 'rect');
        $objWriter->endElement();

        // o:lock
        $objWriter->startElement('o:lock');
        $objWriter->writeAttribute('v:ext', 'edit');
        $objWriter->writeAttribute('aspectratio', 't');
        $objWriter->endElement();

        $objWriter->endElement();

        // Loop through images
        foreach ($images as $key => $value) {
            $this->writeVMLHeaderFooterImage($objWriter, $key, $value);
        }

        $objWriter->endElement();

        // Return
        return $objWriter->getData();
    }

    /**
     * Write VML comment to XML format
     *
     * @param \PhpOffice\PhpSpreadsheet\Shared\XMLWriter        $phpExcelSharedXMLWriter             XML Writer
     * @param    string                            $pReference            Reference
     * @param \PhpOffice\PhpSpreadsheet\Worksheet\HeaderFooterDrawing    $phpExcelWorksheetHeaderFooterDrawing        Image
     * @throws     \PhpOffice\PhpSpreadsheet\Writer\Exception
     */
    private function writeVMLHeaderFooterImage(\PhpOffice\PhpSpreadsheet\Shared\XMLWriter $phpExcelSharedXMLWriter = \null, $pReference = '', \PhpOffice\PhpSpreadsheet\Worksheet\HeaderFooterDrawing $phpExcelWorksheetHeaderFooterDrawing = \null)
    {
        // Calculate object id
        \preg_match('{(\d+)}', \md5($pReference), $m);
        $id = 1500 + (\substr($m[1], 0, 2) * 1);

        // Calculate offset
        $width = $phpExcelWorksheetHeaderFooterDrawing->getWidth();
        $height = $phpExcelWorksheetHeaderFooterDrawing->getHeight();
        $marginLeft = $phpExcelWorksheetHeaderFooterDrawing->getOffsetX();
        $marginTop = $phpExcelWorksheetHeaderFooterDrawing->getOffsetY();

        // v:shape
        $phpExcelSharedXMLWriter->startElement('v:shape');
        $phpExcelSharedXMLWriter->writeAttribute('id', $pReference);
        $phpExcelSharedXMLWriter->writeAttribute('o:spid', '_x0000_s' . $id);
        $phpExcelSharedXMLWriter->writeAttribute('type', '#_x0000_t75');
        $phpExcelSharedXMLWriter->writeAttribute('style', "position:absolute;margin-left:{$marginLeft}px;margin-top:{$marginTop}px;width:{$width}px;height:{$height}px;z-index:1");

        // v:imagedata
        $phpExcelSharedXMLWriter->startElement('v:imagedata');
        $phpExcelSharedXMLWriter->writeAttribute('o:relid', 'rId' . $pReference);
        $phpExcelSharedXMLWriter->writeAttribute('o:title', $phpExcelWorksheetHeaderFooterDrawing->getName());
        $phpExcelSharedXMLWriter->endElement();

        // o:lock
        $phpExcelSharedXMLWriter->startElement('o:lock');
        $phpExcelSharedXMLWriter->writeAttribute('v:ext', 'edit');
        $phpExcelSharedXMLWriter->writeAttribute('rotation', 't');
        $phpExcelSharedXMLWriter->endElement();

        $phpExcelSharedXMLWriter->endElement();
    }


    /**
     * Get an array of all drawings
     *
     * @return     \PhpOffice\PhpSpreadsheet\Worksheet\Drawing[]        All drawings in PHPExcel
     * @throws     \PhpOffice\PhpSpreadsheet\Writer\Exception
     */
    public function allDrawings(\PhpOffice\PhpSpreadsheet\Spreadsheet $phpExcel = \null)
    {
        // Get an array of all drawings
        $aDrawings    = array();

        // Loop through PHPExcel
        $sheetCount = $phpExcel->getSheetCount();
        for ($i = 0; $i < $sheetCount; ++$i) {
            // Loop through images and add to array
            $iterator = $phpExcel->getSheet($i)->getDrawingCollection()->getIterator();
            while ($iterator->valid()) {
                $aDrawings[] = $iterator->current();

                  $iterator->next();
            }
        }

        return $aDrawings;
    }
}

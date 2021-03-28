<?php

class PHPExcel_Reader_Excel5_Style_FillPattern
{
    protected static $map = array(
        0x00 => \PhpOffice\PhpSpreadsheet\Style\Fill::FILL_NONE,
        0x01 => \PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID,
        0x02 => \PhpOffice\PhpSpreadsheet\Style\Fill::FILL_PATTERN_MEDIUMGRAY,
        0x03 => \PhpOffice\PhpSpreadsheet\Style\Fill::FILL_PATTERN_DARKGRAY,
        0x04 => \PhpOffice\PhpSpreadsheet\Style\Fill::FILL_PATTERN_LIGHTGRAY,
        0x05 => \PhpOffice\PhpSpreadsheet\Style\Fill::FILL_PATTERN_DARKHORIZONTAL,
        0x06 => \PhpOffice\PhpSpreadsheet\Style\Fill::FILL_PATTERN_DARKVERTICAL,
        0x07 => \PhpOffice\PhpSpreadsheet\Style\Fill::FILL_PATTERN_DARKDOWN,
        0x08 => \PhpOffice\PhpSpreadsheet\Style\Fill::FILL_PATTERN_DARKUP,
        0x09 => \PhpOffice\PhpSpreadsheet\Style\Fill::FILL_PATTERN_DARKGRID,
        0x0A => \PhpOffice\PhpSpreadsheet\Style\Fill::FILL_PATTERN_DARKTRELLIS,
        0x0B => \PhpOffice\PhpSpreadsheet\Style\Fill::FILL_PATTERN_LIGHTHORIZONTAL,
        0x0C => \PhpOffice\PhpSpreadsheet\Style\Fill::FILL_PATTERN_LIGHTVERTICAL,
        0x0D => \PhpOffice\PhpSpreadsheet\Style\Fill::FILL_PATTERN_LIGHTDOWN,
        0x0E => \PhpOffice\PhpSpreadsheet\Style\Fill::FILL_PATTERN_LIGHTUP,
        0x0F => \PhpOffice\PhpSpreadsheet\Style\Fill::FILL_PATTERN_LIGHTGRID,
        0x10 => \PhpOffice\PhpSpreadsheet\Style\Fill::FILL_PATTERN_LIGHTTRELLIS,
        0x11 => \PhpOffice\PhpSpreadsheet\Style\Fill::FILL_PATTERN_GRAY125,
        0x12 => \PhpOffice\PhpSpreadsheet\Style\Fill::FILL_PATTERN_GRAY0625,
    );

    /**
     * Get fill pattern from index
     * OpenOffice documentation: 2.5.12
     *
     * @param int $index
     * @return string
     */
    public static function lookup($index)
    {
        if (isset(self::$map[$index])) {
            return self::$map[$index];
        }
        return \PhpOffice\PhpSpreadsheet\Style\Fill::FILL_NONE;
    }
}
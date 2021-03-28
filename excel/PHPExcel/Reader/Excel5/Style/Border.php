<?php

class PHPExcel_Reader_Excel5_Style_Border
{
    protected static $map = array(
        0x00 => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_NONE,
        0x01 => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN,
        0x02 => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_MEDIUM,
        0x03 => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_DASHED,
        0x04 => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_DOTTED,
        0x05 => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THICK,
        0x06 => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_DOUBLE,
        0x07 => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_HAIR,
        0x08 => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_MEDIUMDASHED,
        0x09 => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_DASHDOT,
        0x0A => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_MEDIUMDASHDOT,
        0x0B => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_DASHDOTDOT,
        0x0C => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_MEDIUMDASHDOTDOT,
        0x0D => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_SLANTDASHDOT,
    );

    /**
     * Map border style
     * OpenOffice documentation: 2.5.11
     *
     * @param int $index
     * @return string
     */
    public static function lookup($index)
    {
        if (isset(self::$map[$index])) {
            return self::$map[$index];
        }
        return \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_NONE;
    }
}
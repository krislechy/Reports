<?php

namespace Yolva\Reports;

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Style\Alignment;
use PhpOffice\PhpSpreadsheet\Style\Style;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

class Excel
{
    public string $pathTemplate;
    public string $tempTemplate;
    public Spreadsheet $spreadsheet;
    public function __construct($pathTemplate, $tempTemplate)
    {
        $this->pathTemplate = $pathTemplate;
        $this->tempTemplate = $tempTemplate;
        $this->init();
    }
    private function init()
    {
        $this->spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load($this->pathTemplate);
    }
    public function save()
    {
        $writer = new Xlsx($this->spreadsheet);
        $writer->save($this->tempTemplate);
    }
    /**
     * @return Worksheet
     */
    public function getSheet()
    {
        return $this->spreadsheet->getActiveSheet();
    }
    // fill:CCCCFF
    // color:0000FF
    public function setStyle(string $rangeOrCell, string $fill = "none", string $color = "000000", bool $bold = false, $aligment = Alignment::HORIZONTAL_LEFT)
    {
        $style = $this->getStyle($rangeOrCell);
        if ($fill != "none") {
            $style
                ->getFill()
                ->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID)
                ->getStartColor()
                ->setARGB($fill);
        } else {
            $style
                ->getFill()
                ->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_NONE);
        }
        $style
            ->getFont()
            ->getColor()
            ->setARGB($color);
        $style
            ->getFont()
            ->setBold($bold);
        $style
            ->getAlignment()
            ->setHorizontal($aligment);
    }
    public function getStyle(string $rangeOrCell): Style
    {
        return $this->spreadsheet
            ->getActiveSheet()
            ->getStyle($rangeOrCell);
    }
    public function setLocale(string $locale = 'ru'): bool
    {
        return \PhpOffice\PhpSpreadsheet\Settings::setLocale($locale);
    }
    public function translateFormula(string $formula): string
    {
        return \PhpOffice\PhpSpreadsheet\Calculation\Calculation::getInstance()->_translateFormulaToEnglish($formula);
    }
}

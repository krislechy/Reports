<?php

namespace Yolva\Reports;

class PDFConverter
{
    public function __construct()
    {
        $domPdfPath = $_SERVER['DOCUMENT_ROOT'] . '/local/php_interface/vendor/mpdf/mpdf';
        \PhpOffice\PhpWord\Settings::setPdfRendererPath($domPdfPath);
        \PhpOffice\PhpWord\Settings::setPdfRendererName(\PhpOffice\PhpWord\Settings::PDF_RENDERER_MPDF);
    }
    private function getContent(string $docxPath)
    {
        return \PhpOffice\PhpWord\IOFactory::load($docxPath);
    }
    public function convert(string $docxPath, string $outputPath)
    {
        $content = $this->getContent($docxPath);
        $PDFWriter = \PhpOffice\PhpWord\IOFactory::createWriter($content, 'PDF');
        $PDFWriter->save($outputPath);
    }
}

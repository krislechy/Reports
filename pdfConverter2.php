<?php

namespace Yolva\Reports;

use Mnvx\Lowrapper\Converter;
use Mnvx\Lowrapper\Format;
use Mnvx\Lowrapper\LowrapperParameters;

class PDFConverter2
{
    private $converter;
    private $parameters;
    public function __construct()
    {
        $this->converter = new Converter();
    }
    private function setParameters(string $docxPath, string $outputPath)
    {
        $this->parameters = (new LowrapperParameters())
            ->setInputFile($docxPath)
            ->setOutputFormat(Format::TEXT_PDF)
            ->setOutputFile($outputPath);
    }
    public function convert(string $docxPath, string $outputPath)
    {
        $this->setParameters($docxPath, $outputPath);
        $this->converter
            ->convert($this->parameters);
    }
}

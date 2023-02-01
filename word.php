<?php

namespace Yolva\Reports;

use RecursiveDirectoryIterator;
use RecursiveIteratorIterator;
use ZipArchive;

class Word
{
    private string $pathTemplate;
    public string $content;
    public function __construct($pathTemplate)
    {
        $this->pathTemplate = $pathTemplate;
        $this->content = $this->getContent();
    }
    public function getContent($pathTemplate = null): string
    {
        return file_get_contents($pathTemplate == null ? $this->pathTemplate : $pathTemplate);
    }
    public function setContent(array $param, &$content = null)
    {
        if ($content == null)
            foreach ($param as $key => $value) {
                $this->content = str_replace("{{$key}}", $value, $this->content);
            }
        else
            foreach ($param as $key => $value) {
                $content = str_replace("{{$key}}", $value, $content);
            }
    }
    public function addContent($pathTemplate, $array, &$resultContent)
    {
        $content = $this->getContent($pathTemplate);
        $this->setContent($array, $content);
        $resultContent .= $content;
    }
    private function zip($pathTemplate, $pathResultDoc)
    {
        // $rootPath = realpath('folder-to-zip');
        $rootPath = realpath($pathTemplate);

        $zip = new ZipArchive();
        $zip->open($pathResultDoc, ZipArchive::CREATE | ZipArchive::OVERWRITE);

        /** @var SplFileInfo[] $files */
        $files = new RecursiveIteratorIterator(
            new RecursiveDirectoryIterator($rootPath),
            RecursiveIteratorIterator::LEAVES_ONLY
        );

        foreach ($files as $name => $file) {
            if (!$file->isDir()) {
                $filePath = $file->getRealPath();
                $relativePath = substr($filePath, strlen($rootPath) + 1);
                $zip->addFile($filePath, $relativePath);
            }
        }
        $zip->close();
    }
    public function save($documentXml, $pathTemplate, $pathResultDoc)
    {
        file_put_contents($documentXml, $this->content);
        $this->zip($pathTemplate, $pathResultDoc);
    }
}

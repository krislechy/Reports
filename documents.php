<?php

namespace Yolva\Reports;

use Bitrix\Main\Diag\Debug;
use CCrmOwnerType;
use Error;
use PhpOffice\PhpSpreadsheet\Style\Alignment;
use PhpOffice\PhpSpreadsheet\Style\Border;
use Yolva\Helper\FileEntity;

class Documents
{
    private $entityId;
    private $entityTypeId;
    public function __construct($rawEntityId)
    {
        if (empty($rawEntityId))
            throw new Error('Input data of entity is empty');
        if (!strpos($rawEntityId, '_'))
            throw new Error('Data of entity is incorrect format, example: T80_5, current: ' . $rawEntityId);
        $explodedRaw = explode('_', $rawEntityId);
        $this->entityTypeId = hexdec($explodedRaw[0]);
        $this->entityId = $explodedRaw[1];

        $this->baseUrl = $_SERVER['DOCUMENT_ROOT'] . "/local/yolva/classes/reports/";
        $this->tkpDocxUrl = $this->baseUrl . "templates/docx/tkp/";
    }

    public function Create()
    {
        $insideQuoteName = "Смета_внутренняя";
        $customerQuoteName = "Смета_для_заказчика";
        $partnerQuoteName = "Смета_для_партнера";
        $tkpName = "ТКП";

        $dataInstance = new DataTable($this->entityTypeId, $this->entityId);
        $data = $dataInstance->getData();
        $quote = $dataInstance->getQuote();
        $companyType = $dataInstance->getQuote()['COMPANY']['COMPANY_TYPE']; //'Партнер' => CUSTOMER - перепутано в bitrix
        $this->formInsideQuote(
            $this->baseUrl . "templates/" . $insideQuoteName . ".xlsx",
            $this->baseUrl . "tmp/" . $insideQuoteName . ".xlsx",
            $data
        );

        $this->formCustomerQuote(
            $this->baseUrl . "templates/" . $customerQuoteName . ".xlsx",
            $this->baseUrl . "tmp/" . $customerQuoteName . ".xlsx",
            $data
        );
        if ($companyType == "CUSTOMER")
            $this->formPartnerQuote(
                $this->baseUrl . "templates/" . $partnerQuoteName . ".xlsx",
                $this->baseUrl . "tmp/" . $partnerQuoteName . ".xlsx",
                $data
            );
        $this->formTKP(
            $this->baseUrl . "tmp/" . $tkpName . ".docx",
            $data,
            $quote
        );

        $sp = new FileEntity($this->entityTypeId, $this->entityId);
        $filesSmeta = [
            $this->baseUrl . "tmp/" . $insideQuoteName . ".xlsx",
            $this->baseUrl . "tmp/" . $customerQuoteName . ".xlsx",
        ];
        if ($companyType == "CUSTOMER")
            $filesSmeta[] = $this->baseUrl . "tmp/" . $partnerQuoteName . ".xlsx";

        $pdfConverter = new PDFConverter2();
        $pdfConverter->convert($this->baseUrl . "tmp/" . $tkpName . ".docx", $this->baseUrl . "tmp/" . $tkpName . ".pdf");

        $sp->saveFilesToFields(
            [
                'UF_CRM_2_1667808065' => $filesSmeta, //файлы сметы
                'UF_CRM_2_1667814073' => [ //файлы ткп
                    $this->baseUrl . "tmp/" . $tkpName . ".docx",
                    $this->baseUrl . "tmp/" . $tkpName . ".pdf"
                ]
            ],
            "reports"
        );
        $dealId = $sp->entityItem->getData()['PARENT_ID_' . CCrmOwnerType::Deal];
        $deal = new FileEntity(CCrmOwnerType::Deal, $dealId);
        $deal->saveFilesToFields(
            [
                'UF_CRM_1661434966094' => $filesSmeta, //файлы сметы
                'UF_CRM_1661434930696' => [ //файлы ткп
                    $this->baseUrl . "tmp/" . $tkpName . ".docx",
                    $this->baseUrl . "tmp/" . $tkpName . ".pdf"
                ],
            ],
            "reports"
        );

        $this->deleteFileFromDir($this->baseUrl . 'tmp/*');
    }
    private function floatToCurrencyStr($currency)
    {
        if (!isset($currency)) return "";
        $exploded = explode('.', $currency);
        return "{$exploded[0]} руб. {$exploded[1]} коп.";
    }
    private function deleteFileFromDir($pattern)
    {
        $files = glob($pattern);
        foreach ($files as $file) {
            if (is_file($file)) {
                unlink($file);
            }
        }
    }
    private function formTKP(string $outputPath, $productQuotes, $quote)
    {
        $helper = new Word($this->tkpDocxUrl . "tkp_docs_template.xml");

        $helper->setContent([
            "COMPANY_NAME" => $quote['COMPANY']['TITLE'],
            "DATE" => date('d.m.Y')
        ]);

        $oneTimeTaxRate = 0;
        $periodTaxRate = 0;
        $oneTimeTotalSum = 0;
        $periodTotalSum = 0;
        foreach ($productQuotes as $productQuote) {
            $isHeaderOneTime = true;
            $isHeaderPeriodic = true;
            foreach ($productQuote['GOODS'] as $good) {
                if ($good['PERIODICHNOST_PLATEZHA_HBCJZ9'] == "Разовая") {
                    if ($isHeaderOneTime)
                        $helper->addContent(
                            $this->tkpDocxUrl . "onetime_payment_row.xml",
                            [
                                "PRODUCT_NAME" => $productQuote['TITLE'],
                                "MEASURE_NAME" => "",
                                "PRICE" => "",
                                "QUANTITY" => "",
                                "ROW_SUM" => "",
                            ],
                            $oneTimeRows
                        );
                    unset($isHeaderOneTime);

                    $oneTimeTaxRate = $good['TAX_RATE'];
                    $oneTimeTotalSum += $good['PRICE'] * $good['QUANTITY'];
                    $helper->addContent(
                        $this->tkpDocxUrl . "onetime_payment_row.xml",
                        [
                            "PRODUCT_NAME" => $this->getProductName($good),
                            "MEASURE_NAME" => $good['MEASURE_NAME'],
                            "PRICE" => $good['PRICE'],
                            "QUANTITY" => $good['QUANTITY'] + 0,
                            "ROW_SUM" => $good['PRICE'] * $good['QUANTITY'],
                        ],
                        $oneTimeRows
                    );
                } else {
                    if ($isHeaderPeriodic)
                        $helper->addContent(
                            $this->tkpDocxUrl . "period_payment_row.xml",
                            [
                                "PRODUCT_NAME" => $productQuote['TITLE'],
                                "MEASURE_NAME" => "",
                                "PRICE" => "",
                                "QUANTITY" => "",
                                "ROW_SUM" => "",
                            ],
                            $periodRows
                        );
                    unset($isHeaderPeriodic);

                    $periodTaxRate = $good['TAX_RATE'];
                    $periodTotalSum += $good['PRICE'] * $good['QUANTITY'];
                    $helper->addContent(
                        $this->tkpDocxUrl . "period_payment_row.xml",
                        [
                            "PRODUCT_NAME" => $this->getProductName($good),
                            "MEASURE_NAME" => $good['MEASURE_NAME'],
                            "PRICE" => $good['PRICE'],
                            "QUANTITY" => $good['QUANTITY'] + 0,
                            "ROW_SUM" => $good['PRICE'] * $good['QUANTITY'],
                        ],
                        $periodRows
                    );
                }
            }
        }
        if (isset($oneTimeRows))
            $helper->addContent(
                $this->tkpDocxUrl . "onetime_payment_total.xml",
                [
                    "TOTAL_SUM" => $this->floatToCurrencyStr($oneTimeTotalSum),
                    "TAX_SUM" => $this->floatToCurrencyStr($oneTimeTotalSum * $oneTimeTaxRate / 100),
                ],
                $oneTimeRows
            );
        if (isset($periodRows))
            $helper->addContent(
                $this->tkpDocxUrl . "period_payment_total.xml",
                [
                    "TOTAL_SUM" => $this->floatToCurrencyStr($periodTotalSum),
                    "TAX_SUM" => $this->floatToCurrencyStr($periodTotalSum * $periodTaxRate / 100),
                ],
                $periodRows
            );

        if (!isset($oneTimeRows) || !isset($periodRows))
            $isSingle = true;

        if (isset($oneTimeRows))
            $helper->addContent(
                $this->tkpDocxUrl . "onetime_payment_table.xml",
                [
                    "ONETIME_ROWS" => $oneTimeRows,
                ],
                $oneTimeTable,
            );
        if (isset($periodRows))
            $helper->addContent(
                $this->tkpDocxUrl . "period_payment_table.xml",
                [
                    "PERIOD_ROWS" => $periodRows,
                    "NUM_ID_REF" => $isSingle ? "2" : "6",
                ],
                $periodTable,
            );

        if (isset($oneTimeTable)) {
            $helper->addContent(
                $this->tkpDocxUrl . "onetime_payment_header.xml",
                [
                    "REF_NUMBER" => "2.1.",
                ],
                $oneTimeHeader,
            );
        }
        if (isset($periodTable)) {
            $helper->addContent(
                $this->tkpDocxUrl . "period_payment_header.xml",
                [
                    "REF_NUMBER" => $isSingle ? "2.1." : "2.2.",
                ],
                $periodHeader,
            );
        }
        $helper->setContent(
            [
                "ONETIME_TABLE" => $oneTimeTable,
                "PERIOD_TABLE" => $periodTable,

                "ONETIME_HEADER" => $oneTimeHeader,
                "PERIOD_HEADER" => $periodHeader,
            ]
        );

        $helper->save(
            $this->tkpDocxUrl . 'template/word/document.xml',
            $this->tkpDocxUrl . 'template',
            $outputPath
        );
    }
    private function formPartnerQuote(string $pathTemplate, string $outputPath, $data)
    {
        $helper = new Excel($pathTemplate, $outputPath);
        $helper->setLocale();
        $sheet = $helper->getSheet($pathTemplate);
        $firstRow = 11;
        $currentRow = $firstRow;
        foreach ($data as $d) {
            $helper->getStyle("A{$currentRow}:M{$currentRow}")
                ->getBorders()
                ->getAllBorders()
                ->setBorderStyle(Border::BORDER_THIN)
                ->getColor()
                ->setARGB('FFFFFF');

            $helper->getStyle("B{$currentRow}:L{$currentRow}")
                ->getBorders()
                ->getBottom()
                ->setBorderStyle(Border::BORDER_THIN)
                ->getColor()
                ->setARGB('B7B7B7');

            $sheet->setCellValue("A{$currentRow}", $currentRow - 10);
            $helper->setStyle("A{$currentRow}", 'none', 'B7B7B7', false, Alignment::HORIZONTAL_RIGHT);

            $sheet->setCellValue("B{$currentRow}", $d['TITLE']);
            $helper->setStyle("B{$currentRow}:H{$currentRow}", 'none', '000000', true, Alignment::HORIZONTAL_LEFT);

            foreach ($d['GOODS'] as $good) {
                $currentRow++;
                $helper->getStyle("A{$currentRow}:M{$currentRow}")
                    ->getBorders()
                    ->getAllBorders()
                    ->setBorderStyle(Border::BORDER_THIN)
                    ->getColor()
                    ->setARGB('FFFFFF');

                $helper->getStyle("B{$currentRow}:L{$currentRow}")
                    ->getBorders()
                    ->getBottom()
                    ->setBorderStyle(Border::BORDER_THIN)
                    ->getColor()
                    ->setARGB('B7B7B7');

                $sheet->setCellValue("A{$currentRow}", $currentRow - 10);
                $helper->setStyle("A{$currentRow}", 'none', 'B7B7B7', false, Alignment::HORIZONTAL_RIGHT);

                $sheet->setCellValue("B{$currentRow}", $this->getProductName($good));
                $helper->setStyle("B{$currentRow}");

                $sheet->setCellValue("C{$currentRow}", $good['MEASURE_NAME']);
                $helper->setStyle("C{$currentRow}", 'none', '000000', false, Alignment::HORIZONTAL_CENTER);

                $sheet->setCellValue("D{$currentRow}", $good['QUANTITY']);
                $helper->setStyle("D{$currentRow}", 'none', '000000', true, Alignment::HORIZONTAL_CENTER);

                $freqPay = $good['PERIODICHNOST_PLATEZHA_HBCJZ9'];

                if ($freqPay == "Разовая") {
                    $sheet->setCellValue("G{$currentRow}", 0);
                    $sheet->setCellValue("E{$currentRow}", $good['PRICE']);
                } else {
                    $sheet->setCellValue("G{$currentRow}", $good['PRICE']);
                    $sheet->setCellValue("E{$currentRow}", 0);
                }

                if (isset($good['VARIATIONS'])) {
                    foreach ($good['VARIATIONS'] as $variation) {
                        if ($variation['PRICE_TYPE']  == 'Партнер') {
                            $taxPrice = $variation['PRICE']['PRICE'] + ($variation['PRICE']['PRICE'] * ($good['TAX_RATE'] / 100));
                            if ($freqPay == "Разовая") {
                                $sheet->setCellValue("K{$currentRow}", 0);
                                $sheet->setCellValue("I{$currentRow}", $taxPrice);
                            } else {
                                $sheet->setCellValue("K{$currentRow}", $taxPrice);
                                $sheet->setCellValue("I{$currentRow}", 0);
                            }
                            break;
                        }
                    }
                }

                $sheet->setCellValue("F{$currentRow}", $helper->translateFormula('=ЕСЛИ(ДВССЫЛ(СЦЕПИТЬ("E";СТРОКА())) <> 0;(ДВССЫЛ(СЦЕПИТЬ("E";СТРОКА())) *ДВССЫЛ(СЦЕПИТЬ("D";СТРОКА())));"")'));
                $sheet->setCellValue("H{$currentRow}", $helper->translateFormula('=ЕСЛИ(ДВССЫЛ(СЦЕПИТЬ("G";СТРОКА())) <> 0;(ДВССЫЛ(СЦЕПИТЬ("G";СТРОКА())) *ДВССЫЛ(СЦЕПИТЬ("D";СТРОКА())));"")'));
                $sheet->setCellValue("J{$currentRow}", $helper->translateFormula('=ЕСЛИ(ДВССЫЛ(СЦЕПИТЬ("I";СТРОКА())) <> 0;(ДВССЫЛ(СЦЕПИТЬ("I";СТРОКА())) *ДВССЫЛ(СЦЕПИТЬ("D";СТРОКА())));"")'));
                $sheet->setCellValue("L{$currentRow}", $helper->translateFormula('=ЕСЛИ(ДВССЫЛ(СЦЕПИТЬ("K";СТРОКА())) <> 0;(ДВССЫЛ(СЦЕПИТЬ("K";СТРОКА())) *ДВССЫЛ(СЦЕПИТЬ("D";СТРОКА())));"")'));

                $helper->setStyle("E{$currentRow}:L{$currentRow}", 'none', '000000', false, Alignment::HORIZONTAL_RIGHT);
            }
            $currentRow++;
        }

        $sheet
            ->getColumnDimension('B')
            ->setAutoSize(true);
        $sheet->calculateColumnWidths();

        $totalCell = $currentRow - 1;
        $totalRow = [
            "F{$currentRow}" => $helper->translateFormula("=СУММ(F{$firstRow}:F{$totalCell})"),
            "H{$currentRow}" => $helper->translateFormula("=СУММ(H{$firstRow}:H{$totalCell})"),
            "J{$currentRow}" => $helper->translateFormula("=СУММ(J{$firstRow}:J{$totalCell})"),
            "L{$currentRow}" => $helper->translateFormula("=СУММ(L{$firstRow}:L{$totalCell})"),
        ];

        foreach ($totalRow as $coordinate => $value) {
            $sheet->setCellValue($coordinate, $value);
            $helper->setStyle($coordinate, 'none', '000000', true, Alignment::HORIZONTAL_RIGHT);
        }
        $currentRow++;
        $helper->getStyle("B{$currentRow}:L{$currentRow}")
            ->getBorders()
            ->getBottom()
            ->setBorderStyle(Border::BORDER_THIN);
        $helper->save();
    }
    private function formCustomerQuote(string $pathTemplate, string $outputPath, $data)
    {
        $helper = new Excel($pathTemplate, $outputPath);
        $helper->setLocale();
        $sheet = $helper->getSheet($pathTemplate);
        $firstRow = 11;
        $currentRow = $firstRow;
        foreach ($data as $d) {
            $sheet->setCellValue("B{$currentRow}", $d['TITLE']);
            $helper->setStyle("B{$currentRow}:H{$currentRow}", 'CCCCFF', '000000', true, Alignment::HORIZONTAL_LEFT);

            $sheet->setCellValue("A{$currentRow}", $currentRow - 10);
            $helper->setStyle("A{$currentRow}", 'none', '000000', true, Alignment::HORIZONTAL_CENTER);

            foreach ($d['GOODS'] as $good) {
                $currentRow++;

                $sheet->setCellValue("A{$currentRow}", $currentRow - 10);
                $helper->setStyle("A{$currentRow}", 'none', '000000', true, Alignment::HORIZONTAL_CENTER);

                $sheet->setCellValue("B{$currentRow}", $this->getProductName($good));
                $helper->setStyle("B{$currentRow}");

                $sheet->setCellValue("C{$currentRow}", $good['MEASURE_NAME']);
                $helper->setStyle("C{$currentRow}", 'none', '000000', false, Alignment::HORIZONTAL_CENTER);

                $sheet->setCellValue("D{$currentRow}", $good['QUANTITY']);
                $helper->setStyle("D{$currentRow}", 'none', '000000', true, Alignment::HORIZONTAL_CENTER);

                $freqPay = $good['PERIODICHNOST_PLATEZHA_HBCJZ9'];

                if ($freqPay == "Разовая") {
                    $sheet->setCellValue("G{$currentRow}", 0);
                    $sheet->setCellValue("E{$currentRow}", $good['PRICE']);
                } else {
                    $sheet->setCellValue("G{$currentRow}", $good['PRICE']);
                    $sheet->setCellValue("E{$currentRow}", 0);
                }

                $sheet->setCellValue("F{$currentRow}", $helper->translateFormula('=ЕСЛИ(ДВССЫЛ(СЦЕПИТЬ("E";СТРОКА())) <> 0;(ДВССЫЛ(СЦЕПИТЬ("E";СТРОКА())) *ДВССЫЛ(СЦЕПИТЬ("D";СТРОКА())));"")'));
                $sheet->setCellValue("H{$currentRow}", $helper->translateFormula('=ЕСЛИ(ДВССЫЛ(СЦЕПИТЬ("G";СТРОКА())) <> 0;(ДВССЫЛ(СЦЕПИТЬ("G";СТРОКА())) *ДВССЫЛ(СЦЕПИТЬ("D";СТРОКА())));"")'));

                $helper->setStyle("E{$currentRow}:H{$currentRow}", 'none', '0000FF', false, Alignment::HORIZONTAL_RIGHT);

                //style border
                $AH = $helper->getStyle("A{$currentRow}:H{$currentRow}");
                $AH->getBorders()->getAllBorders()->setBorderStyle(Border::BORDER_THIN);
            }
            $AH = $helper->getStyle("A{$currentRow}:H{$currentRow}");
            $AH->getBorders()->getBottom()->setBorderStyle(Border::BORDER_MEDIUM);
            $currentRow++;
        }

        $totalCell = $currentRow - 1;
        $totalRow = [
            "E{$currentRow}" => "ИТОГО:",
            "F{$currentRow}" => $helper->translateFormula("=СУММ(F{$firstRow}:F{$totalCell})"),
            "H{$currentRow}" => $helper->translateFormula("=СУММ(H{$firstRow}:H{$totalCell})"),
        ];

        $sheet->getDefaultRowDimension()->setRowHeight(-1);
        foreach ($sheet->getRowDimensions() as $row) {
            if ($row->getRowIndex() >= $firstRow)
                $row->setRowHeight(-1);
        }
        $sheet->getStyle('A')->getAlignment()->setWrapText(true);

        foreach ($totalRow as $coordinate => $value) {
            $sheet->setCellValue($coordinate, $value);

            //style border
            $helper->setStyle($coordinate, 'none', '000000', true, Alignment::HORIZONTAL_CENTER);
            $AL = $helper->getStyle($coordinate);
            $AL->getBorders()->getAllBorders()->setBorderStyle(Border::BORDER_MEDIUM);
            if ($coordinate != "E{$currentRow}")
                $helper->setStyle($coordinate, 'none', '0000FF', true, Alignment::HORIZONTAL_RIGHT);
        }

        $helper->save();
    }
    private function formInsideQuote(string $pathTemplate, string $outputPath, $data)
    {
        $helper = new Excel($pathTemplate, $outputPath);
        $helper->setLocale();
        $sheet = $helper->getSheet($pathTemplate);
        $firstRow = 11;
        $currentRow = $firstRow;
        foreach ($data as $d) {
            $sheet->setCellValue("A{$currentRow}", $d['TITLE']);
            $helper->setStyle("A{$currentRow}:L{$currentRow}", 'CCCCFF', '000000', true, Alignment::HORIZONTAL_LEFT);

            foreach ($d['GOODS'] as $good) {
                $currentRow++;
                $sheet->setCellValue("A{$currentRow}", $this->getProductName($good));
                $helper->setStyle("A{$currentRow}");

                $sheet->setCellValue("B{$currentRow}", $good['MEASURE_NAME']);
                $helper->setStyle("B{$currentRow}", 'none', '000000', false, Alignment::HORIZONTAL_CENTER);

                $sheet->setCellValue("C{$currentRow}", $good['QUANTITY']);
                $helper->setStyle("C{$currentRow}", 'none', '000000', true, Alignment::HORIZONTAL_CENTER);

                $freqPay = $good['PERIODICHNOST_PLATEZHA_HBCJZ9'];
                $period = $good['MINIMALNYY_PERIOD_TARIFIKATSII_EN33BQ'];
                $sheet->setCellValue("D{$currentRow}", $period);
                $helper->setStyle("D{$currentRow}", 'none', '0000FF', true, Alignment::HORIZONTAL_CENTER);

                if ($freqPay == "Разовая") {
                    $sheet->setCellValue("E{$currentRow}", 0);
                    $sheet->setCellValue("F{$currentRow}", $good['PRICE']);
                } else {
                    $sheet->setCellValue("E{$currentRow}", $good['PRICE']);
                    $sheet->setCellValue("F{$currentRow}", 0);
                }

                $sheet->setCellValue("G{$currentRow}", $helper->translateFormula('=ЕСЛИ(ДВССЫЛ(СЦЕПИТЬ("E";СТРОКА())) <> 0;(ДВССЫЛ(СЦЕПИТЬ("E";СТРОКА())) * ДВССЫЛ(СЦЕПИТЬ("C";СТРОКА())));"")'));
                $sheet->setCellValue("H{$currentRow}", $helper->translateFormula('=ЕСЛИ(ДВССЫЛ(СЦЕПИТЬ("F";СТРОКА())) <> 0;(ДВССЫЛ(СЦЕПИТЬ("F";СТРОКА())) *ДВССЫЛ(СЦЕПИТЬ("C";СТРОКА())));"")'));
                $sheet->setCellValue("I{$currentRow}", $good['PARAMETRS']['PURCHASING_PRICE']);
                $sheet->setCellValue("J{$currentRow}", $helper->translateFormula('=ЕСЛИ(ДВССЫЛ(СЦЕПИТЬ("I";СТРОКА())) <> 0;(ДВССЫЛ(СЦЕПИТЬ("I";СТРОКА())) *ДВССЫЛ(СЦЕПИТЬ("C";СТРОКА())));"")'));
                $sheet->setCellValue("K{$currentRow}", $helper->translateFormula('=ЕСЛИ(ДВССЫЛ(СЦЕПИТЬ("G";СТРОКА())) <> "";ОКРУГЛ((ОКРУГЛ((ДВССЫЛ(СЦЕПИТЬ("G";СТРОКА())) - (ДВССЫЛ(СЦЕПИТЬ("G";СТРОКА())) * 0.2 / 1.2) );2)-ДВССЫЛ(СЦЕПИТЬ("J";СТРОКА())));4);ОКРУГЛ((ОКРУГЛ((ДВССЫЛ(СЦЕПИТЬ("H";СТРОКА()))-(ДВССЫЛ(СЦЕПИТЬ("H";СТРОКА())) * 0.2 / 1.2) );2)-ДВССЫЛ(СЦЕПИТЬ("J";СТРОКА())));4))'));
                // $sheet->setCellValue("L{$currentRow}", $helper->translateFormula('=ЕСЛИ(ДВССЫЛ(СЦЕПИТЬ("G";СТРОКА()))<>"";ОКРУГЛ(  (ДВССЫЛ(СЦЕПИТЬ("K";СТРОКА()))  /  ОКРУГЛ((ДВССЫЛ(СЦЕПИТЬ("G";СТРОКА()))  - (ДВССЫЛ(СЦЕПИТЬ("G";СТРОКА())) * 0.2 / 1.2) );2)  *  100);2 );ОКРУГЛ((ДВССЫЛ(СЦЕПИТЬ("K";СТРОКА()))  /  ОКРУГЛ((ДВССЫЛ(СЦЕПИТЬ("H";СТРОКА()))  - (ДВССЫЛ(СЦЕПИТЬ("H";СТРОКА())) * 0.2 / 1.2) );2)  *  100);2))'));
                $sheet->setCellValue("L{$currentRow}", $helper->translateFormula('=ЕСЛИ(K' . $currentRow . '<>0;ЕСЛИ(ДВССЫЛ(СЦЕПИТЬ("G";СТРОКА()))<>"";ОКРУГЛ((ДВССЫЛ(СЦЕПИТЬ("K";СТРОКА()))/ОКРУГЛ((ДВССЫЛ(СЦЕПИТЬ("G";СТРОКА()))-(ДВССЫЛ(СЦЕПИТЬ("G";СТРОКА()))*0.2/1.2));2)*100);2);ОКРУГЛ((ДВССЫЛ(СЦЕПИТЬ("K";СТРОКА()))/ОКРУГЛ((ДВССЫЛ(СЦЕПИТЬ("H";СТРОКА()))-(ДВССЫЛ(СЦЕПИТЬ("H";СТРОКА()))*0.2/1.2));2)*100);2));0)'));

                $helper->setStyle("E{$currentRow}:J{$currentRow}", 'none', '0000FF', false, Alignment::HORIZONTAL_RIGHT);
                $helper->setStyle("K{$currentRow}:L{$currentRow}", 'none', '000000', false, Alignment::HORIZONTAL_RIGHT);

                //style border
                $BL = $helper->getStyle("A{$currentRow}:L{$currentRow}");
                $BL->getBorders()->getAllBorders()->setBorderStyle(Border::BORDER_THIN);

                $AL = $helper->getStyle("A{$currentRow}:L{$currentRow}");
                $AL->getBorders()->getBottom()->setBorderStyle(Border::BORDER_MEDIUM);
            }
            $currentRow++;
        }

        $totalCell = $currentRow - 1;
        $totalRow = [
            "F{$currentRow}" => "ИТОГО:",
            "G{$currentRow}" => $helper->translateFormula("=СУММ(G{$firstRow}:G{$totalCell})"),
            "H{$currentRow}" => $helper->translateFormula("=СУММ(H{$firstRow}:H{$totalCell})"),
            "J{$currentRow}" => $helper->translateFormula("=СУММ(J{$firstRow}:J{$totalCell})"),
            "K{$currentRow}" => $helper->translateFormula("=ОКРУГЛ(СУММ(K{$firstRow}:K{$totalCell});2)"),
            "L{$currentRow}" => $helper->translateFormula("=ОКРУГЛ((ОКРУГЛ(СУММ(K{$firstRow}:K{$totalCell});2) / (((СУММ(G{$firstRow}:G{$totalCell}) + СУММ(H{$firstRow}:H{$totalCell})) - ОКРУГЛ(( ( СУММ(G{$firstRow}:G{$totalCell}) + СУММ(H{$firstRow}:H{$totalCell}) ) * 0.2 / 1.2);2))) * 100 ); 2)")
        ];
        // foreach ($sheet->getColumnDimensions() as $column)
        //     $column->setAutoSize(true);
        $sheet->getDefaultRowDimension()->setRowHeight(-1);
        foreach ($sheet->getRowDimensions() as $row) {
            if ($row->getRowIndex() >= $firstRow)
                $row->setRowHeight(-1);
        }
        $sheet->getStyle('A')->getAlignment()->setWrapText(true);

        foreach ($totalRow as $coordinate => $value) {
            $sheet->setCellValue($coordinate, $value);

            //style border
            $helper->setStyle($coordinate, 'none', '000000', true, Alignment::HORIZONTAL_CENTER);
            $AL = $helper->getStyle($coordinate);
            $AL->getBorders()->getAllBorders()->setBorderStyle(Border::BORDER_MEDIUM);
            if ($coordinate != "F{$currentRow}")
                $helper->setStyle($coordinate, 'none', '0000FF', true, Alignment::HORIZONTAL_RIGHT);
        }

        $helper->save();
    }

    private function getProductName($good)
    {
        $productName = $good['PRODUCT_NAME'];
        if (empty($productName))
            $productName = $good['ORIGINAL_PRODUCT_NAME'];
        if (empty($productName))
            $productName = $good['POLNOE_NAZVANIE_FVV6MZ'];
        return $productName;
    }
}

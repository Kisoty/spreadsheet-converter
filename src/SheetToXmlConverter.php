<?php

declare(strict_types=1);

namespace DC\V3\Parser;

use Google\Service\Sheets\CellData;
use Google\Service\Sheets\CellFormat;
use Google\Service\Sheets\ExtendedValue;
use Google\Service\Sheets\GridData;
use Google\Service\Sheets\RowData;
use Google\Service\Sheets\Sheet;
use Google_Client;
use Google_Service_Sheets;
use SimpleXMLElement;

final class SheetToXmlConverter
{
    private Google_Service_Sheets $sheetsService;

    private SimpleXMLElement $xml;

    private array $columnIndexes;

    public function __construct(string $developerKey)
    {
        $client = new Google_Client();
        $client->setDeveloperKey($developerKey);
        $this->sheetsService = new Google_Service_Sheets($client);
        $this->prepareSheetColumnIndexes();
    }

    // По идее это надо не считать каждый раз, а впихнуть в массив и забыть,
    // но мне чувство прекрасного не позволяет такую огромную дуру накатать
    private function prepareSheetColumnIndexes(): void
    {
        $this->columnIndexes = [];
        $letter = 'A';

        while ($letter !== 'AAA') {
            $this->columnIndexes[] = $letter++;
        }
    }

    public function convertToXml(string $filename, string $spreadSheetId, array $ranges = null)
    {
        $spreadSheet = $this->sheetsService->spreadsheets->get($spreadSheetId, [
            'includeGridData' => true,
            'ranges' => $ranges
        ]);

        $this->setXmlBase(
            $spreadSheet->getSpreadsheetId(),
            $spreadSheet->getProperties()->getTitle(),
            $spreadSheet->getProperties()->getAutoRecalc(),
            $spreadSheet->getProperties()->getLocale(),
            $spreadSheet->getProperties()->getTimeZone()
        );

        foreach ($spreadSheet->getSheets() as $sheet) {
            $this->addSheet($sheet);
        }

        $this->saveToFile($filename);
    }

    private function setXmlBase(string $spreadSheetId, string $spreadSheetTitle,
                                string $autoRecalc, string $locale, string $timezone): void
    {
        $this->xml = new SimpleXMLElement(<<<XML
<?xml version="1.0" encoding="UTF-8"?>
<spreadsheet id="$spreadSheetId" title="$spreadSheetTitle" autoRecalc="$autoRecalc" locale="$locale" timezone="$timezone">
</spreadsheet>
XML);
    }

    private function addSheet(Sheet $googleSheet): void
    {
        if ($googleSheet->getProperties()->getSheetType() !== 'GRID') {
            throw new \Exception('Only grid sheets supported.');
        }

        if (count($googleSheet->getData()) > 1) {
            throw new \Exception('Only 1 range per sheet supported.');
        }

        $sheetNode = $this->xml->addChild('sheet');
        $sheetNode->addAttribute('id', (string)$googleSheet->getProperties()->getSheetId());
        $sheetNode->addAttribute('title', $googleSheet->getProperties()->getTitle());

        $rowNum = $this->getStartRowNum(current($googleSheet->getData()));
        $startColumnNum = $this->getStartColumnNum(current($googleSheet->getData()));

        foreach (current($googleSheet->getData())->getRowData() as $rowData) {
            $this->addRowToSheet($sheetNode, $rowData, $rowNum, $startColumnNum);
            $rowNum++;
        }
    }

    private function getStartRowNum(GridData $gridData): int
    {
        if ($gridData->getStartRow() === null) {
            return 1;
        } else {
            return $gridData->getStartRow() + 1;
        }
    }

    private function getStartColumnNum(GridData $gridData): int
    {
        return $gridData->getStartColumn() ?? 0;
    }

    private function addRowToSheet(SimpleXMLElement $sheetNode, RowData $rowData, int $rowNum, int $startColumnNum): void
    {
        $rowNode = $sheetNode->addChild('row');
        $rowNode->addAttribute('id', (string)$rowNum);

        $columnNum = $startColumnNum;

        foreach ($rowData->getValues() as $cellData) {
            $this->addColumnToRow($rowNode, $cellData, $columnNum);
            $columnNum++;
        }
    }

    private function addColumnToRow(SimpleXMLElement $rowNode, CellData $cellData, int $columnNum): void
    {
        $columnNode = $rowNode->addChild('column');
        $columnNode->addAttribute('id', $this->columnIndexes[$columnNum]);

        if ($cellData->getUserEnteredValue() === null) {
            $columnNode->addChild('formula', '');
        } else {
            $columnNode->addChild('formula', $cellData->getUserEnteredValue()->getFormulaValue());
        }

        $columnNode->addChild('hyperlink', $cellData->getHyperlink() ?? '');

        $this->setColumnValue($columnNode, $cellData->getEffectiveValue());
        $this->setColumnStyle($columnNode, $cellData->getEffectiveFormat());
    }

    private function setColumnValue(SimpleXMLElement $columnNode, ?ExtendedValue $effectiveValue): void
    {
        if ($effectiveValue === null) {
            $columnNode->addChild('value', '');
        } elseif ($effectiveValue->getBoolValue() !== null) {
            $columnNode->addChild('value', (string)$effectiveValue->getBoolValue());
        } elseif ($effectiveValue->getNumberValue() !== null) {
            $columnNode->addChild('value', (string)$effectiveValue->getNumberValue());
        } elseif ($effectiveValue->getStringValue() !== null) {
            $columnNode->addChild('value', (string)$effectiveValue->getStringValue());
        } elseif ($effectiveValue->getErrorValue() !== null) {
            $columnNode->addChild('value', $effectiveValue->getErrorValue()->getMessage());
        } else {
            $columnNode->addChild('value', '');
        }
    }

    private function setColumnStyle(SimpleXMLElement $columnNode, CellFormat $cellFormat): void
    {
        $styleNode = $columnNode->addChild('style');
        $styleNode->addChild('bgc', );

    }

    private function saveToFile(string $fileName): void
    {
        $dom = new \DOMDocument();
        $dom->preserveWhiteSpace = false;
        $dom->formatOutput = true;
        $dom->loadXML($this->xml->asXML());
        $dom->saveXML();
        $dom->save($fileName);
    }
}

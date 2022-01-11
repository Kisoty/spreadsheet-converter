<?php

declare(strict_types=1);

namespace DC\V3\SheetConverter;

use DC\V3\SheetConverter\Exceptions\SheetParseException;
use Google\Client;
use Google\Service\Sheets;
use Google\Service\Sheets\CellData;
use Google\Service\Sheets\CellFormat;
use Google\Service\Sheets\Color;
use Google\Service\Sheets\ExtendedValue;
use Google\Service\Sheets\GridData;
use Google\Service\Sheets\Padding;
use Google\Service\Sheets\RowData;
use Google\Service\Sheets\Sheet;
use Google\Service\Sheets\TextFormat;
use SimpleXMLElement;

final class SheetToXmlConverter
{
    private Sheets $sheetsService;

    /**
     * @var array Соответствие int индексов буквенным в google sheets
     */
    private array $columnIndexes;

    public function __construct(Client $client)
    {
        $this->sheetsService = new Sheets($client);
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

    /**
     * @throws SheetParseException
     */
    public function convert(string $filename, string $spreadSheetId, array $ranges = null): void
    {
        $spreadSheet = $this->sheetsService->spreadsheets->get($spreadSheetId, [
            'includeGridData' => true,
            'ranges' => $ranges
        ]);

        $xml = $this->createBaseXml(
            $spreadSheet->getSpreadsheetId(),
            $spreadSheet->getProperties()->getTitle(),
            $spreadSheet->getProperties()->getAutoRecalc(),
            $spreadSheet->getProperties()->getLocale(),
            $spreadSheet->getProperties()->getTimeZone()
        );

        foreach ($spreadSheet->getSheets() as $sheet) {
            $this->addSheet($xml, $sheet);
        }

        $this->saveToFile($xml, $filename);
    }

    private function createBaseXml(string $spreadSheetId, string $spreadSheetTitle,
                                   string $autoRecalc, string $locale, string $timezone): SimpleXMLElement
    {
        return new SimpleXMLElement(<<<XML
<?xml version="1.0" encoding="UTF-8"?>
<spreadsheet id="$spreadSheetId" title="$spreadSheetTitle" autoRecalc="$autoRecalc" locale="$locale" timezone="$timezone">
</spreadsheet>
XML);
    }

    /**
     * @throws SheetParseException
     */
    private function addSheet(SimpleXMLElement $xml, Sheet $googleSheet): void
    {
        if ($googleSheet->getProperties()->getSheetType() !== 'GRID') {
            throw new SheetParseException('Only grid sheets supported.');
        }

        if (count($googleSheet->getData()) > 1) {
            throw new SheetParseException('Only 1 range per sheet supported.');
        }

        $sheetNode = $xml->addChild('sheet');
        $sheetNode->addAttribute('id', (string)$googleSheet->getProperties()->getSheetId());
        $sheetNode->addAttribute('title', $googleSheet->getProperties()->getTitle());
        $this->addMergesToSheet($sheetNode, $googleSheet->getMerges());

        $rowNum = $this->getStartRowNum(current($googleSheet->getData()));
        $startColumnNum = $this->getStartColumnNum(current($googleSheet->getData()));

        foreach (current($googleSheet->getData())->getRowData() as $rowData) {
            $this->addRowToSheet($sheetNode, $rowData, $rowNum, $startColumnNum);
            $rowNum++;
        }
    }

    /**
     * @param Sheets\GridRange[] $merges
     */
    private function addMergesToSheet(SimpleXMLElement $sheetNode, array $merges): void
    {
        foreach ($merges as $merge) {
            $mergeNode = $sheetNode->addChild('merge');

            $mergeNode->addChild('startRowIndex', (string)$merge->getStartRowIndex());
            $mergeNode->addChild('endRowIndex', (string)$merge->getEndRowIndex());
            $mergeNode->addChild('startColumnIndex', (string)$merge->getStartColumnIndex());
            $mergeNode->addChild('endColumnIndex', (string)$merge->getEndColumnIndex());
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
            $columnNode->addChild('value', '')->addAttribute('type', 'NULL');
        } elseif ($effectiveValue->getBoolValue() !== null) {
            $columnNode->addChild('value', (string)$effectiveValue->getBoolValue())
                ->addAttribute('type', 'BOOLEAN');
        } elseif ($effectiveValue->getNumberValue() !== null) {
            $columnNode->addChild('value', (string)$effectiveValue->getNumberValue())
                ->addAttribute('type', 'NUMBER');
        } elseif ($effectiveValue->getStringValue() !== null) {
            $columnNode->addChild('value', (string)$effectiveValue->getStringValue())
                ->addAttribute('type', 'STRING');
        } elseif ($effectiveValue->getErrorValue() !== null) {
            $columnNode->addChild('value', $effectiveValue->getErrorValue()->getMessage())
                ->addAttribute('type', 'STRING');
        } else {
            $columnNode->addChild('value', '')->addAttribute('type', 'NULL');
        }
    }

    private function setColumnStyle(SimpleXMLElement $columnNode, ?CellFormat $cellFormat): void
    {
        if ($cellFormat === null) {
            return;
        }

        $styleNode = $columnNode->addChild('style');

        if ($cellFormat->getBackgroundColor() !== null) {
            $bgcNode = $styleNode->addChild('bgc');

            $this->fillColorNode($bgcNode, $cellFormat->getBackgroundColor());
        }

        if ($cellFormat->getTextFormat() !== null) {
            $textFormatNode = $styleNode->addChild('textFormat');

            $this->fillTextFormatNode($textFormatNode, $cellFormat->getTextFormat());
        }

        if ($cellFormat->getPadding() !== null) {
            $paddingNode = $styleNode->addChild('padding');

            $this->fillPaddingNode($paddingNode, $cellFormat->getPadding());
        }

        $alignmentNode = $styleNode->addChild('alignment');
        $alignmentNode->addChild('vertical', $cellFormat->getVerticalAlignment() ?? '');
        $alignmentNode->addChild('horizontal', $cellFormat->getHorizontalAlignment() ?? '');

        $styleNode->addChild('wrapStrategy', $cellFormat->getWrapStrategy() ?? '');
    }

    private function fillColorNode(SimpleXMLElement $node, Color $apiColor): void
    {
        if ($apiColor->getAlpha() !== null) {
            $node->addChild('alpha', (string)$apiColor->getAlpha());
        }
        if ($apiColor->getBlue() !== null) {
            $node->addChild('blue', (string)$apiColor->getBlue());
        }
        if ($apiColor->getGreen() !== null) {
            $node->addChild('green', (string)$apiColor->getGreen());
        }
        if ($apiColor->getRed() !== null) {
            $node->addChild('red', (string)$apiColor->getRed());
        }
    }

    private function fillTextFormatNode(SimpleXMLElement $node, TextFormat $apiFormat): void
    {
        if ($apiFormat->getFontFamily() !== null) {
            $node->addChild('fontFamily', $apiFormat->getFontFamily());
        }
        if ($apiFormat->getFontSize() !== null) {
            $node->addChild('fontSize', (string)$apiFormat->getFontSize());
        }

        if ($apiFormat->getForegroundColor() !== null) {
            $fontColorNode = $node->addChild('fontColor');

            $this->fillColorNode($fontColorNode, $apiFormat->getForegroundColor());
        }

        if ($apiFormat->getItalic()) {
            $node->addChild('italic');
        }
        if ($apiFormat->getBold()) {
            $node->addChild('bold');
        }
        if ($apiFormat->getUnderline()) {
            $node->addChild('underline');
        }
        if ($apiFormat->getStrikethrough()) {
            $node->addChild('strikethrough');
        }
    }

    private function fillPaddingNode(SimpleXMLElement $node, Padding $apiPadding): void
    {
        if ($apiPadding->getBottom() !== null) {
            $node->addChild('bottom', (string)$apiPadding->getBottom());
        }
        if ($apiPadding->getTop() !== null) {
            $node->addChild('top', (string)$apiPadding->getTop());
        }
        if ($apiPadding->getLeft() !== null) {
            $node->addChild('left', (string)$apiPadding->getLeft());
        }
        if ($apiPadding->getRight() !== null) {
            $node->addChild('right', (string)$apiPadding->getRight());
        }
    }

    private function saveToFile(SimpleXMLElement $xml, string $fileName): void
    {
        $dom = new \DOMDocument();
        $dom->preserveWhiteSpace = false;
        $dom->formatOutput = true;
        $dom->loadXML($xml->asXML());
        $dom->saveXML();
        $dom->save($fileName);
    }
}

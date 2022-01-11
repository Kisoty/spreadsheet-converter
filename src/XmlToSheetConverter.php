<?php

declare(strict_types=1);

namespace DC\V3\SheetConverter;

use Google\Client;
use Google\Service\Drive;
use Google\Service\Sheets;

final class XmlToSheetConverter
{
    private Client $apiClient;

    /**
     * @var array Соответствие int индексов буквенным в google sheets
     */
    private array $columnIndexes;

    public function __construct(Client $client)
    {
        $this->apiClient = $client;

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

    // TODO добавить валидацию переданного xml на основе xsd схемы до создания пустого файла на гугл диске
    //todo spreadsheet attributes
    public function convert(string $xmlFilename, string $spreadSheetName, string $folderId = null)
    {
        $xml = new \SimpleXMLElement($xmlFilename);

        $spreadSheetId = $this->createEmptySheet($spreadSheetName, $folderId);

        $sheetService = new Sheets($this->apiClient);

        $sheetRequests = [];

        $sheetRequests = $this->appendUpdateSheetPropertiesRequest($sheetRequests, $xml);

        foreach ($xml->sheet as $sheet) {
            $sheetRequests = $this->appendAddSheetRequest($sheetRequests, $sheet);
            $sheetId = (int)$sheet->attributes()['id'];

            $sheetRequests = $this->appendUpdateCellsRequest($sheetRequests, $sheet->row, $sheetId);
            $sheetRequests = $this->appendMergeCellsRequests($sheetRequests, $sheet->merge, $sheetId);
        }

        $batchUpdateRequest = new Sheets\BatchUpdateSpreadsheetRequest([
            'requests' => $sheetRequests
        ]);

        $sheetService->spreadsheets->batchUpdate($spreadSheetId, $batchUpdateRequest);
    }

    private function appendUpdateSheetPropertiesRequest(array $requestArray, \SimpleXMLElement $spreadSheetNode): array
    {
        $requestArray[] = [
            'updateSpreadsheetProperties' => [
                'properties' => [
                    'title' => (string)$spreadSheetNode->attributes()['title'],
                    'locale' => (string)$spreadSheetNode->attributes()['locale'],
                    'timeZone' => (string)$spreadSheetNode->attributes()['timezone'],
                    'autoRecalc' => (string)$spreadSheetNode->attributes()['autoRecalc']
                ],
                'fields' => '*'
            ],
        ];

        return $requestArray;
    }

    private function appendAddSheetRequest(array $requestArray, ?\SimpleXMLElement $sheetNode): array
    {
        if ($sheetNode !== null) {
            $properties = [
                'title' => (string)$sheetNode->attributes()['title'],
                'sheetId' => (int)$sheetNode->attributes()['id']
            ];

            if ((int)$sheetNode->attributes()['id'] === 0) {
                $requestArray[] = [
                    'updateSheetProperties' => [
                        'properties' => $properties,
                        'fields' => 'title'
                    ]
                ];
            } else {
                $requestArray[] = [
                    'addSheet' => [
                        'properties' => $properties
                    ]
                ];
            }
        }

        return $requestArray;
    }

    private function appendUpdateCellsRequest(array $requestArray, ?\SimpleXMLElement $rowNodes, int $sheetId): array
    {
        $updateCellsRequest = [
            'updateCells' => [
                'rows' => [],
                'fields' => '*'
            ]
        ];

        if ($rowNodes !== null) {
            $startRowIndex = (int)current($rowNodes)['id'] - 1;
            $startColumnIndex = array_search(current($rowNodes[0]->column)['id'], $this->columnIndexes, true);

            foreach ($rowNodes as $rowNode) {
                $rowData = [];

                foreach ($rowNode->column as $cellNode) {
                    $cellData = [];

                    $cellValue = $this->parseCellValue($cellNode);
                    if ($cellValue !== null) {
                        $cellData['userEnteredValue'] = $cellValue;
                    }

                    $cellFormat = $this->parseCellFormat($cellNode);
                    if ($cellFormat !== null) {
                        $cellData['userEnteredFormat'] = $cellFormat;
                    }

                    $rowData['values'][] = $cellData;
                }

                $updateCellsRequest['updateCells']['rows'][] = $rowData;
            }
        }

        if (isset($startRowIndex, $startColumnIndex)) {
            $updateCellsRequest['updateCells']['range'] = [
                'sheetId' => $sheetId,
                'startRowIndex' => $startRowIndex,
                'startColumnIndex' => $startColumnIndex
            ];

            $requestArray[] = $updateCellsRequest;
        }

        return $requestArray;
    }

    private function parseCellValue(\SimpleXMLElement $cellNode): ?array
    {
        if ($cellNode->value[0] === null) {
            return null;
        }

        $valueType = current($cellNode->value)['type'];

        if (!empty((string)$cellNode->formula[0])) {
            return ['formulaValue' => (string)$cellNode->formula[0]];
        }

        switch ($valueType) {
            case 'BOOLEAN':
                return ['boolValue' => (bool)(int)$cellNode->value[0]];
            case 'NUMBER':
                return ['numberValue' => (float)$cellNode->value[0]];
            case 'STRING':
                return ['stringValue' => (string)$cellNode->value[0]];
            default:
                return null;
        }
    }

    private function parseCellFormat(\SimpleXMLElement $cellNode): ?array
    {
        $cellFormat = [];

        $styleNode = $cellNode->style[0];

        if ($styleNode === null) {
            return null;
        }

        $bgcNode = $styleNode->bgc[0];
        $cellFormat['backgroundColor'] = [
            'red' => (float)$bgcNode->red,
            'green' => (float)$bgcNode->green,
            'blue' => (float)$bgcNode->blue,
        ];

        $cellFormat['wrapStrategy'] = (string)$styleNode->wrapStrategy[0];

        $textFormatNode = $styleNode->textFormat[0];
        $cellFormat['textFormat'] = [
            'fontFamily' => (string)$textFormatNode->fontFamily,
            'fontSize' => (string)$textFormatNode->fontSize,
            'foregroundColor' => [
                'red' => (float)$textFormatNode->fontColor[0]->red,
                'green' => (float)$textFormatNode->fontColor[0]->green,
                'blue' => (float)$textFormatNode->fontColor[0]->blue,
            ]
        ];

        $alignmentNode = $styleNode->alignment[0];

        if (!empty((string)$alignmentNode->horizontal)) {
            $cellFormat['horizontalAlignment'] = (string)$alignmentNode->horizontal;
        }

        if (!empty((string)$alignmentNode->vertical)) {
            $cellFormat['verticalAlignment'] = (string)$alignmentNode->vertical;
        }

        $paddingNode = $styleNode->padding[0];
        $cellFormat['padding'] = [
            'top' => (float)$paddingNode->top,
            'left' => (float)$paddingNode->left,
            'bottom' => (float)$paddingNode->bottom,
            'right' => (float)$paddingNode->right
        ];

        return $cellFormat;
    }

    private function appendMergeCellsRequests(array $sheetRequests, ?\SimpleXMLElement $merges, int $sheetId): array
    {
        foreach ($merges as $merge) {
            $sheetRequests[] = [
                'mergeCells' => [
                    'range' => [
                        'sheetId' => $sheetId,
                        'startRowIndex' => (int)$merge->startRowIndex,
                        'endRowIndex' => (int)$merge->endRowIndex,
                        'startColumnIndex' => (int)$merge->startColumnIndex,
                        'endColumnIndex' => (int)$merge->endColumnIndex,
                    ],
                    'mergeType' => 'MERGE_ALL'
                ]
            ];
        }

        return $sheetRequests;
    }


    /**
     * Creates new spreadsheet and returns its ID
     * @param string $spreadSheetName
     * @param ?string $folderId
     * @return string New Spreadsheet ID
     */
    private function createEmptySheet(string $spreadSheetName, string $folderId = null): string
    {
        $drive = new Drive($this->apiClient);
        $newDriveFile = new Drive\DriveFile();

        if ($folderId !== null) {
            $newDriveFile->setParents([$folderId]);
        }

        $newDriveFile->setMimeType('application/vnd.google-apps.spreadsheet');
        $newDriveFile->setName($spreadSheetName);

        $res = $drive->files->create($newDriveFile);

        return $res->getId();
    }
}

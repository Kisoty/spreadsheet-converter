<?php

declare(strict_types=1);

namespace DC\V3\SheetConverter;

use Google\Client;
use Google\Service\Drive;
use Google\Service\Sheets;

final class XmlToSheetConverter
{
    private Client $apiClient;

    public function __construct(Client $client)
    {
        $this->apiClient = $client;

//        $this->prepareSheetColumnIndexes();
    }


    // TODO добавить валидацию переданного xml на основе xsd схемы до создания пустого файла на гугл диске
    //todo spreadsheet attributes
    public function convert(string $xml, string $spreadSheetName, string $folderId = null)
    {
        $xml = new \SimpleXMLElement($xml);

        $sheetId = $this->createEmptySheet($spreadSheetName, $folderId);

        $sheetService = new Sheets($this->apiClient);

//        $spreadSheet = $sheetService->spreadsheets->get($sheetId, [
//            'includeGridData' => true
//        ]);

        $sheetRequests = [];

        $sheetRequests = $this->appendUpdateSheetPropertiesRequest($sheetRequests, $xml);

        foreach ($xml->sheet as $sheet) {
            $sheetRequests = $this->appendAddSheetRequest($sheetRequests, $sheet);
        }

        $sheetRequests = $this->appendDeleteDefaultSheetRequest($sheetRequests);

        $batchUpdateRequest = new Sheets\BatchUpdateSpreadsheetRequest([
            'requests' => $sheetRequests
        ]);

        $sheetService->spreadsheets->batchUpdate($sheetId, $batchUpdateRequest);
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
            $requestArray[] = [
                'addSheet' => [
                    'properties' => [
                        'title' => (string)$sheetNode->attributes()['title'],
                        'sheetId' => (int)$sheetNode->attributes()['id']
                    ]
                ]
            ];
        }

        return $requestArray;
    }

    private function appendDeleteDefaultSheetRequest(array $requestArray): array
    {
        $requestArray[] = [
            'deleteSheet' => [
                'sheetId' => 0
            ]
        ];

        return $requestArray;
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

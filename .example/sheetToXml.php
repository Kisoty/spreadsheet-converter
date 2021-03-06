<?php

declare(strict_types=1);

use DC\V3\SheetConverter\SheetToXmlConverter;
use Google\Client;
use Google\Service\Drive;
use Google\Service\Sheets;

require __DIR__ . './../vendor/autoload.php';

// Use this function with passing google service account credentials or pass authorized google client object manually
// Might be bad decision cuz of putting env var
function getClient(string $accCredentials): Client {
    $client = new Client();

    putenv(
        'GOOGLE_APPLICATION_CREDENTIALS=' . $accCredentials
    );

    $client->setScopes([Sheets::SPREADSHEETS, Drive::DRIVE]);

    $client->useApplicationDefaultCredentials();

    return $client;
}


$googleCredentials = 'path-to-json-with-credentials';
$client = getClient($googleCredentials);

// Example spreadsheet ID (ren scenario)
$spreadsheetId = '1NzZp5r6O2TBJwg5cIA2yLxcGoViduwVF1yrOw9b457Y';

// Ranges of spreadsheet
// Can be left null or not passed to function, but in this case ALL rows will be retrieved
$ranges = [
    'Notes!A1:A23',
    'Env!A1:E13',
    'TaskCommon!A1:H3',
    'StatusType!A1:H10',
    'StatusGeneral!A1:F157',
    'Scenario!A3:C3',
    'Swimline!A1:E3',
    'Task!A1:I3',
    'StatusHook!A1:L54',
    'StepHook!A1:L98'
];

$sheetToXmlConverter = new SheetToXmlConverter($client);
$xmlString = $sheetToXmlConverter->convert($spreadsheetId, $ranges);

// Filename with path to save parsed sheet
$filename = '../smth.xml';

$sheetToXmlConverter::saveToFile($xmlString, $filename);

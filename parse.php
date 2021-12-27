<?php

declare(strict_types=1);

use DC\V3\SheetConverter\SheetToXmlConverter;
use DC\V3\SheetConverter\XmlToSheetConverter;
use Google\Client;
use Google\Service\Drive;
use Google\Service\Sheets;

require __DIR__ . '/vendor/autoload.php';


function getClient(string $accCredentials): Client {
    $client = new Client();

    putenv(
        'GOOGLE_APPLICATION_CREDENTIALS=' . $accCredentials
    );

    $client->setScopes([Sheets::SPREADSHEETS, Drive::DRIVE]);

    $client->useApplicationDefaultCredentials();

    return $client;
}

$client = getClient(__DIR__ . '/googleTestAuth.json');

//todo composer.json lib config, add exceptions, add example, files and folders structure

//$sheetToXmlConverter = new SheetToXmlConverter($client);

// my one
// 12IbKSuInKseCaLQybjYDnnfBvEL5QVZ8Au0NJ0a9T7A

// ren scenario
// 1NzZp5r6O2TBJwg5cIA2yLxcGoViduwVF1yrOw9b457Y


//$sheetToXmlConverter->convert('smth.xml', '1NzZp5r6O2TBJwg5cIA2yLxcGoViduwVF1yrOw9b457Y', [
////    'Notes!A1:A20',
//    'TaskCommon!A1:H3',
//    'StatusType!A1:H10',
////    'Scenario!B3:C3',
////    'List1!A1:D4',
//]);

// ----------------------------------

$xmlToSheetConverter = new XmlToSheetConverter($client);
$xmlToSheetConverter->convert(file_get_contents('smth.xml'),
    'Create Test ' . (new DateTime())->format('H:i:s d.m.Y'),
    '1GDaXbZ5d6qIhfzvB7yn74QETJ-A_suPN');

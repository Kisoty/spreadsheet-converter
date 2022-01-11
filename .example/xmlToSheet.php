<?php

declare(strict_types=1);

use DC\V3\SheetConverter\XmlToSheetConverter;
use Google\Client;
use Google\Service\Drive;
use Google\Service\Sheets;

require __DIR__ . '../vendor/autoload.php';

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

$xmlString = file_get_contents('smth.xml');

// Default name of created sheet before renaming it to title from xml
// (sheet will be left with this name in case of mistake during parsing)
$defaultSpreadsheetName = 'Create Test ' . (new DateTime())->format('H:i:s d.m.Y');

// ID of folder to save sheet to
$googleDriveFolderId = '1GDaXbZ5d6qIhfzvB7yn74QETJ-A_suPN';

$xmlToSheetConverter = new XmlToSheetConverter($client);
$xmlToSheetConverter->convert($xmlString, $defaultSpreadsheetName, $googleDriveFolderId);

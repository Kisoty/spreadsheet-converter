<?php

declare(strict_types=1);

namespace DC\V3\SheetConverter;

use Google_Client;
use Google_Service_Sheets;

final class XmlToSheetConverter
{
    private Google_Service_Sheets $sheetsService;

    public function __construct(string $developerKey)
    {
        $client = new Google_Client();
        $client->setDeveloperKey($developerKey);
        $this->sheetsService = new Google_Service_Sheets($client);
//        $this->prepareSheetColumnIndexes();
    }

    public function convert(string $filename)
    {

    }
}

<?php

declare(strict_types=1);

namespace DC\V3\Parser;

class SheetToXml
{
    private \SimpleXMLElement $xml;

    public function __construct()
    {

    }

    public function convert()
    {
        $xmlStr = <<<XML
<?xml version="1.0" encoding="UTF-8"?>
<root>
</root>
XML;
        $this->xml = new \SimpleXMLElement($xmlStr);


    }

    private function saveToFile(string $fileName)
    {
        $this->xml->saveXML($fileName);
    }
}

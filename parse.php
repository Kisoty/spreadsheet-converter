<?php

declare(strict_types=1);

use DC\V3\SheetConverter\SheetToXmlConverter;
use DC\V3\SheetConverter\XmlToSheetConverter;

require __DIR__ . '/vendor/autoload.php';

$devKey = 'AIzaSyCdA5GDWP_QGBIocOWbYw-Ppwhkq8sKt2I';

$sheetToXmlConverter = new SheetToXmlConverter($devKey);

// my one
// 12IbKSuInKseCaLQybjYDnnfBvEL5QVZ8Au0NJ0a9T7A

// ren scenario
// 1NzZp5r6O2TBJwg5cIA2yLxcGoViduwVF1yrOw9b457Y


$sheetToXmlConverter->convert('smth.xml', '1NzZp5r6O2TBJwg5cIA2yLxcGoViduwVF1yrOw9b457Y', [
//    'Notes!A1:A20',
    'TaskCommon!A1:H3',
    'StatusType!A1:H10',
//    'Scenario!B3:C3',
//    'List1!A1:D4',
]);

// ----------------------------------

$xmlToSheetConverter = new XmlToSheetConverter($devKey);

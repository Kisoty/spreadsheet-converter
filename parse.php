<?php

declare(strict_types=1);

use DC\V3\Parser\SheetToXmlConverter;

require __DIR__ . '/vendor/autoload.php';

if (php_sapi_name() != 'cli') {
    throw new Exception('This application must be run on the command line.');
}

$converter = new SheetToXmlConverter('AIzaSyCdA5GDWP_QGBIocOWbYw-Ppwhkq8sKt2I');

// my one
// 12IbKSuInKseCaLQybjYDnnfBvEL5QVZ8Au0NJ0a9T7A

// ren scenario
// 1NzZp5r6O2TBJwg5cIA2yLxcGoViduwVF1yrOw9b457Y


$converter->convertToXml('smth.xml', '12IbKSuInKseCaLQybjYDnnfBvEL5QVZ8Au0NJ0a9T7A', [
//    'Notes!A1:A20',
//    'TaskCommon!A1:H3',
//    'Scenario!A1:C3',
    'List1!A1:D4',
]);

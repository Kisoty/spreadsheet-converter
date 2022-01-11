<?php

declare(strict_types=1);

namespace DC\V3\SheetConverter\Exceptions;

class XmlParseException extends \Exception
{
    protected $message = 'Xml cannot be parsed';
}

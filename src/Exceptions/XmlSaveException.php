<?php

declare(strict_types=1);

namespace DC\V3\SheetConverter\Exceptions;

class XmlSaveException extends \Exception
{
    protected $message = 'Error occurred while saving xml to file';
}

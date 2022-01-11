<?php

declare(strict_types=1);

namespace DC\V3\SheetConverter\Exceptions;

class SheetParseException extends \Exception
{
    protected $message = 'Sheet cannot be parsed';
}

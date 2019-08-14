<?php

declare(strict_types=1);

namespace AskNicely\Spreadsheet;

require_once __DIR__ . '/../vendor/autoload.php';

use Symfony\Component\Console\Application;

const APP_NAME = 'ppt-generation';

$commands = [
    new PowerPointMergeCommand()
];

$application = new Application(APP_NAME);
$application->addCommands($commands);
$application->run();

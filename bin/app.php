<?php

declare(strict_types=1);

use AskNicely\Spreadsheet;

require_once __DIR__ . '/../vendor/autoload.php';

use Symfony\Component\Console\Application;

const APP_NAME = 'app';

$commands = [
    new Spreadsheet\PowerPointMergeCommand()
];

$application = new Application(APP_NAME);
$application->addCommands($commands);
$application->run();

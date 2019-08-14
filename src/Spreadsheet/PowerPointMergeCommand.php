<?php

declare(strict_types=1);

namespace AskNicely\Spreadsheet;

require_once __DIR__ . '/../../vendor/autoload.php';

use Symfony\Component\Console\Command\Command;
use Symfony\Component\Console\Input\InputArgument;
use Symfony\Component\Console\Input\InputInterface;
use Symfony\Component\Console\Output\OutputInterface;
use Symfony\Component\Console\Style\SymfonyStyle;

final class PowerPointMergeCommand extends Command
{
    static protected $defaultName = 'spreadsheet:powerpoint:generate';

    public function configure(): void
    {
        $this->setDescription('Create Powerpoint file from template with field values');
        $this->addArgument('input-file', InputArgument::REQUIRED, 'Powerpoint template file');
        $this->addArgument('output-file', InputArgument::REQUIRED, 'Merged Powerpoint file');
        $this->addArgument('data-file', InputArgument::REQUIRED, 'Fields values');
    }

    public function execute(InputInterface $input, OutputInterface $output)
    {
        $io = new SymfonyStyle($input, $output);

        $inputFile = $input->getArgument('input-file');
        $outputFile = $input->getArgument('output-file');
        $dataFile = $input->getArgument('data-file');

        $errors = [];
        if (!file_exists($inputFile)) {
            $errors[] = 'File does not exists: ' . $inputFile;
        }
        if (!file_exists($dataFile)) {
            $errors[] = 'File does not exists: ' . $dataFile;
        }
        if (!empty($errors)) {
            foreach ($errors as $error) {
                $io->error($error);
            }

            return;
        }

        (new PowerPointMerge($inputFile, $outputFile, json_decode(file_get_contents($dataFile) ?: '', true)))->merge();

        $io->success('Done! Successfully generated ' . basename($outputFile));
    }
}

<?php

declare(strict_types=1);

namespace AskNicely\Spreadsheet;

include_once 'vendor/tinybutstrong/tinybutstrong/tbs_class.php';
include_once 'vendor/tinybutstrong/opentbs/tbs_plugin_opentbs.php';

final class PowerPointMerge implements SpreadsheetMergeInterface
{
    /** @var string $templateFile */
    private $templateFile;

    /** @var ?string $outputFile */
    private $outputFile;

    /** @var ?array $data */
    private $data;

    public function __construct(string $templateFile, ?string $outputFile = null, ?array $data = null)
    {
        $this->templateFile = $templateFile;
        $this->outputFile = $outputFile;
        $this->data = $data;

        $this->sanitiseData();
    }

    public function merge(): void
    {
        $tbs = new \clsTinyButStrong;
        $tbs->VarRef = $this->data;

        $tbs->Plugin(TBS_INSTALL, OPENTBS_PLUGIN);
        $tbs->LoadTemplate($this->templateFile);
        $tbs->Plugin(OPENTBS_SELECT_SLIDE, 2);
        $tbs->PlugIn(OPENTBS_CHANGE_PICTURE, $this->data['npsTrend'], ['adjust' => 'sameheight']);
        $tbs->Plugin(OPENTBS_SELECT_SLIDE, 1);
        $tbs->Show(OPENTBS_FILE, $this->outputFile);
    }

    private function sanitiseData(): void
    {
        foreach ($this->data as $field => $value)
        {
            $this->data[$field] = preg_match('/[[:alnum:]]/', $value) ? $value : '';
        }
    }
}

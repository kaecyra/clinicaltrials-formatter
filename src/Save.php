<?php

/**
 * General purpose XML parser and formatter for studies published on clinicaltrials.gov
 *
 * @license MIT
 * @copyright 2021 Tim Gunter
 * @author Tim Gunter <gunter.tim@gmail.com>
 */

namespace Kaecyra\ClinicalTrials;

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

class Save {

    protected $workbook;

    protected $sheets;

    protected $summary;
    protected $rendergroups;
    protected $outcomes;

    protected $styles;

    public function __construct() {
        $this->styles = Styles::getStyles();
    }

    public function build($summary, $rendergroups, $outcomes) {

        $this->workbook = new Spreadsheet();

        $this->summary = $summary;
        $this->rendergroups = $rendergroups;
        $this->outcomes = $outcomes;

        // Write summary sheet

        $titleSheet = new Worksheet($this->workbook, "Summary");
        $this->workbook->addSheet($titleSheet);

        $titleSheet->setCellValue("B2", $this->summary['title']);
        $titleSheet->getStyle('B2:Z2')->applyFromArray($this->styles['title']);
        $titleSheet->setCellValue("B3", $this->summary['url']);
        $titleSheet->getCell('B3')->getHyperlink()->setUrl($this->summary['url']);
        $titleSheet->getStyle('B3:Z3')->applyFromArray($this->styles['subtitle']);

        $titleSheet->setCellValue("B5", "NCT ID");
        $titleSheet->setCellValue("C5", $this->summary['nct']);
        $titleSheet->setCellValue("B6", "Started");
        $titleSheet->setCellValueExplicit("C6", $this->summary['started'], \PhpOffice\PhpSpreadsheet\Cell\DataType::TYPE_STRING);
        $titleSheet->setCellValue("B7", "Completed");
        $titleSheet->setCellValueExplicit("C7", $this->summary['finished'], \PhpOffice\PhpSpreadsheet\Cell\DataType::TYPE_STRING);

        $titleSheet->getColumnDimension('A')->setWidth(6);
        $titleSheet->getColumnDimension('B')->setWidth(14);
        for ($i=3; $i <= 12; $i++) {
            $titleSheet->getColumnDimensionByColumn($i)->setWidth(20);
        }

        $titleSheet->getStyle('C5:C7')->applyFromArray($this->styles['keyvalue']);

        $titleSheet->mergeCells('C5:F5');
        $titleSheet->mergeCells('C6:F6');
        $titleSheet->mergeCells('C7:F7');

        $titleSheet->setSelectedCell('A1');

        // Delete default sheet

        $this->workbook->removeSheetByIndex(0);

        // Write render group sheets

        $rg = 1;
        $rgOutcomeIndex = 0;

        $groupListRow = 9;
        foreach ($this->rendergroups as $renderGroupID => $renderGroup) {
            $renderGroupName = $this->writeRenderGroup($renderGroupID);

            $rgStartRow = $groupListRow + $rgOutcomeIndex;
            foreach ($renderGroup['outcomes'] as $outcomeID) {
                $outcomeRow = $groupListRow + $rgOutcomeIndex;
                $renderGroupOutcome = $this->outcomes[$outcomeID];
                $titleSheet->setCellValue("B{$outcomeRow}", $renderGroupOutcome['title']);
                $rgOutcomeIndex++;
            }
            $rgEndRow = $groupListRow + ($rgOutcomeIndex-1);

            $titleSheet->setCellValue("H{$rgStartRow}", $renderGroupName);
            $titleSheet->mergeCells("H{$rgStartRow}:H{$rgEndRow}");
            $titleSheet->getStyle("F{$rgStartRow}")->applyFromArray($this->styles['rendergroup']);

            $titleSheet->getStyle("B{$rgStartRow}:H{$rgEndRow}")->applyFromArray($this->styles['box']);
            $titleSheet->getStyle("H{$rgStartRow}:H{$rgEndRow}")->applyFromArray($this->styles['box']);

            $rg++;
        }
    }

    protected function writeRenderGroup($renderGroupID) {

        $q = true;

        $renderGroup = &$this->rendergroups[$renderGroupID];

        $rg = substr(hash('sha256',uniqid("",true)), 0, 8);
        if (!$q) echo "Writing render group $renderGroupID ($rg)\n";

        $renderGroupName = "Render Group {$rg}";
        $groupSheet = new Worksheet($this->workbook, "Render Group {$rg}");
        $this->workbook->addSheet($groupSheet);
        $renderGroup['sheet'] = $groupSheet;

        $groupSheet->getColumnDimension('A')->setWidth(6);
        $groupSheet->getColumnDimension('B')->setWidth(24);
        $groupSheet->freezePane('C1');

        for ($i=3; $i <= 12; $i++) {
            $groupSheet->getColumnDimensionByColumn($i)->setWidth(14);
        }

        $groupSheet->setCellValue("B2", "Render Group {$rg}");
        $groupSheet->getStyle('B2:Z2')->applyFromArray($this->styles['title']);

        $groupSheet->setCellValue("B3", "Grouped Outcomes");
        $groupSheet->getStyle('B3:Z3')->applyFromArray($this->styles['subtitle']);

        $topRow = 5;
        $headerSize = 6;

        // Write classes down column B

        if (!$q) echo "  classes:\n";

        $classRows = [];

        $clr = $topRow + count($renderGroup['outcomes']) + $headerSize + 2;
        $firstClassRow = $clr-2;

        // Write analyzed result class

        $arRowStart = $clr-2;
        $arRowEnd = $clr-1;
        $classRows['Analyzed Result'] = $clr-2;
        $groupSheet->setCellValue("B{$arRowStart}", "Analyzed Result");
        $groupSheet->mergeCells("B{$arRowStart}:B{$arRowEnd}");
        $groupSheet->getStyle("B{$arRowStart}:B{$arRowEnd}")->applyFromArray($this->styles['arclass']);

        foreach ($renderGroup['classes'] as $class) {
            if ($class) {
                $groupSheet->setCellValue("B{$clr}", $class);
                if (!$q) echo "    B{$clr}: {$class}\n";

                $classRows[$class] = $clr;
            }
            $clr++;
        }
        $lastClassRow = $clr-1;

        $groupSheet->getStyle("B{$firstClassRow}:B{$lastClassRow}")->applyFromArray($this->styles['class']);

        // Write outcomes

        if (!$q) echo "  outcomes:\n";

        $outcomeRow = $topRow;
        $outcomeStartColumn = 3;

        $dataTop = $topRow + count($renderGroup['outcomes']);
        $outcomeHeaderRow = $dataTop + 2;
        $outcomeUnitsRow = $dataTop + 3;
        $outcomeGroupRow = $dataTop + 4;
        $outcomeMetricRow = $dataTop + 5;

        foreach ($renderGroup['outcomes'] as $outcomeID) {
            $outcome = $this->outcomes[$outcomeID];
            $outcomeTitle = $this->outcomes[$outcomeID]['title'];

            $groupWidth = $outcome['measure']['width'];
            $outcomeWidth = count($outcome['groups']) * $groupWidth;

            $minSize = 14;
            if ($groupWidth == 1) {
                $minSize = 28;
            }

            for ($i=$outcomeStartColumn; $i < $outcomeStartColumn+$outcomeWidth; $i++) {
                $groupSheet->getColumnDimensionByColumn($i)->setWidth($minSize);
            }

            // Write outcome into list of outcomes in render group

            if (!$q) echo "    B{$outcomeRow}: {$outcomeTitle}\n";
            $groupSheet->setCellValue("B{$outcomeRow}", $outcomeTitle);
            $groupSheet->getStyle("B{$outcomeRow}:Z{$outcomeRow}")->applyFromArray($this->styles['outcome']);
            $outcomeRow++;

            // Write outcome name (column group)

            $sColIndex = $outcomeStartColumn;
            $eColIndex = $outcomeStartColumn + ($outcomeWidth - 1);
            $sCol = $this->getCol($sColIndex);
            $eCol = $this->getCol($eColIndex);

            if (!$q) echo "    {$sCol}{$outcomeHeaderRow}: {$outcomeTitle}\n";
            $groupSheet->setCellValue("{$sCol}{$outcomeHeaderRow}", $outcomeTitle);
            $groupSheet->getStyle("{$sCol}{$outcomeHeaderRow}:{$eCol}{$outcomeHeaderRow}")->applyFromArray($this->styles['header']);
            $groupSheet->mergeCells("{$sCol}{$outcomeHeaderRow}:{$eCol}{$outcomeHeaderRow}");

            // Write outcome units

            $unitsLabel = $outcome['measure']['param'];
            if (!empty($outcome['measure']['units'])) {
                $unitsLabel .= ", as {$outcome['measure']['units']}";
            }
            if (!empty($outcome['measure']['dispersion'])) {
                $unitsLabel .= ", with {$outcome['measure']['dispersion']}";
            }
            $groupSheet->setCellValue("{$sCol}{$outcomeUnitsRow}", $unitsLabel);
            $groupSheet->getStyle("{$sCol}{$outcomeUnitsRow}:{$eCol}{$outcomeUnitsRow}")->applyFromArray($this->styles['headerunits']);
            $groupSheet->mergeCells("{$sCol}{$outcomeUnitsRow}:{$eCol}{$outcomeUnitsRow}");

            // Borderlize header row

            $groupSheet->getStyle("{$sCol}{$outcomeHeaderRow}:{$eCol}{$outcomeUnitsRow}")->applyFromArray($this->styles['headerbox']);

            // Write groups

            if (!$q) echo "    groups:\n";

            $groupColumns = [];

            $groupColumn = $outcomeStartColumn;
            foreach ($outcome['groups'] as $outcomeGroupID => $outcomeGroupName) {
                $groupStartColIndex = $groupColumn;
                $groupEndColIndex = $groupColumn + ($groupWidth - 1);
                $groupStartCol = $this->getCol($groupStartColIndex);
                $groupEndCol = $this->getCol($groupEndColIndex);

                // Write group name

                if (!$q) echo "      {$groupStartCol}{$outcomeGroupRow}: [{$outcomeGroupID}] {$outcomeGroupName}\n";
                $groupSheet->setCellValue("{$groupStartCol}{$outcomeGroupRow}", $outcomeGroupName);

                $groupColumns[$outcomeGroupID] = $groupStartColIndex;

                $groupSheet->getStyle("{$groupStartCol}{$outcomeGroupRow}:{$groupEndCol}{$outcomeGroupRow}")->applyFromArray($this->styles['group']);
                if ($groupWidth > 1) {
                    $groupSheet->mergeCells("{$groupStartCol}{$outcomeGroupRow}:{$groupEndCol}{$outcomeGroupRow}");
                }

                // Write class metrics

                $k = 0;
                $nKeys = 0;
                foreach ($outcome['measure']['keys'] as $groupMetricLabelKey) {
                    if ($groupMetricLabelKey == 'group_id') {
                        continue;
                    }
                    $nKeys++;

                    $metricColIndex = $groupStartColIndex+$k;
                    $metricCol = $this->getCol($metricColIndex);

                    // Write metric label
                    $groupSheet->setCellValue("{$metricCol}{$outcomeMetricRow}", ucwords(str_replace('_',' ',$groupMetricLabelKey)));
                    $groupSheet->getStyle("{$metricCol}{$outcomeMetricRow}")->applyFromArray($this->styles['metric']);
                    $k++;
                }

                $groupSheet->getStyle("{$groupStartCol}{$outcomeMetricRow}:{$groupEndCol}{$outcomeMetricRow}")->applyFromArray($this->styles['box']);
                $groupSheet->getStyle("{$groupStartCol}{$firstClassRow}:{$groupEndCol}{$lastClassRow}")->applyFromArray($this->styles['box'])->applyFromArray($this->styles['metricvalue']);

                $groupColumn += $groupWidth;
            }

            // Write class data (metrics)

            if (!$q) echo "      metrics:\n";

            // Write analyzed results

            if (!empty($outcome['measure']['analyzed'])) {

                // Labels

                foreach ($outcome['measure']['analyzed']['data'] as $analyzedGroupID => $analyzedGroupData) {

                    $analyzedGroupLookupStartColIndex = $groupColumns[$analyzedGroupID];
                    $analyzedGroupLookupStartCol = $this->getCol($analyzedGroupLookupStartColIndex);
                    $analyzedGroupLookupEndColIndex = $analyzedGroupLookupStartColIndex + ($groupWidth-1);
                    $analyzedGroupLookupEndCol = $this->getCol($analyzedGroupLookupEndColIndex);

                    $groupSheet->mergeCells("{$analyzedGroupLookupStartCol}{$arRowStart}:{$analyzedGroupLookupEndCol}{$arRowStart}");
                    $groupSheet->setCellValue("{$analyzedGroupLookupStartCol}{$arRowStart}", $outcome['measure']['analyzed']['units']);
                    $groupSheet->getStyle("{$analyzedGroupLookupStartCol}{$arRowStart}")->applyFromArray($this->styles['analyzedresult']);

                    $groupSheet->mergeCells("{$analyzedGroupLookupStartCol}{$arRowEnd}:{$analyzedGroupLookupEndCol}{$arRowEnd}");
                    $groupSheet->setCellValue("{$analyzedGroupLookupStartCol}{$arRowEnd}", $analyzedGroupData['value']);
                    $groupSheet->getStyle("{$analyzedGroupLookupStartCol}{$arRowEnd}")->applyFromArray($this->styles['analyzedresult']);

                    $groupSheet->getStyle("{$analyzedGroupLookupStartCol}{$arRowStart}:{$analyzedGroupLookupEndCol}{$arRowEnd}")->applyFromArray($this->styles['box']);
                }
            }

            foreach ($outcome['measure']['raw'] as $rawMeasureClassMetrics) {
                $rawMeasureClassMetricsClassName = $rawMeasureClassMetrics['class'];
                $measureItemRow = $classRows[$rawMeasureClassMetricsClassName] ?? 0;

                if (!$q) echo "        $rawMeasureClassMetricsClassName (row {$measureItemRow}):\n";

                $metricGroupIndex = 0;
                $metricGroupColumn = $outcomeStartColumn;
                $metricGroupStartColIndex = $metricGroupColumn;

                foreach ($rawMeasureClassMetrics['data'] as $groupID => $groupMetricList) {

                    $metricGroupLookupStartCol = $groupColumns[$groupID];

                    if (!$q) echo "          {$groupID}: ".implode(', ', array_slice(array_values($groupMetricList), 1))."\n";

                    $metricGroupColIndex = $metricGroupStartColIndex + ($metricGroupIndex * $groupWidth);
                    $k = 0;
                    foreach ($groupMetricList as $groupMetricKey => $groupMetricValue) {
                        if ($groupMetricKey == 'group_id') {
                            continue;
                        }

                        $metricColIndex = $metricGroupLookupStartCol + $k;
                        $metricCol = $this->getCol($metricColIndex);
                        $groupSheet->setCellValue("{$metricCol}{$measureItemRow}", $groupMetricValue);
                        $k++;
                    }

                    $metricGroupIndex++;
                }
            }

            $outcomeStartColumn += $outcomeWidth;
        }

        $groupSheet->setSelectedCell('A1');

        return $renderGroupName;
    }

    public function save($file) {
        $writer = new Xlsx($this->workbook);
        $writer->save($file);
    }

    protected function getCol($index) {
        $chunk = floor($index/26);
        if (!($index % 26)) {
            $chunk--;
        }

        if ($chunk) {
            $prefN = 64 + $chunk;
            $pref = chr($prefN);
        } else {
            $pref = '';
        }

        $chrOff = $index % 26 ? $index % 26 : 26;
        $chr = chr(64 + $chrOff);
        return "{$pref}{$chr}";
    }

}
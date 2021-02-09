<?php

/**
 * General purpose XML parser and formatter for studies published on clinicaltrials.gov
 *
 * @license MIT
 * @copyright 2021 Tim Gunter
 * @author Tim Gunter <gunter.tim@gmail.com>
 */

namespace Kaecyra\ClinicalTrials;

class ClinicalTrialParser {

    /**
     * @var string
     */
    protected $file;

    /**
     * @var \SimpleXMLElement
     */
    protected $xml;

    /**
     * @var string
     */
    protected $title;

    /**
     * @var array
     */
    protected $summary;

    /**
     * @var array
     */
    protected $groups;

    /**
     * @var array
     */
    protected $periods;

    /**
     * @var array
     */
    protected $outcomes;

    /**
     * @var array
     */
    protected $classes;

    /**
     * @var array
     */
    protected $classUsage;

    /**
     * @var array
     */
    protected $rendergroups;

    /**
     * @var numeric
     */
    protected $threshold;

    /**
     * @var array
     */
    protected $parity;


    protected $debug;

    public function __construct() {
        $this->setSimilarityThreshold(75);
        $this->setUnitParityRequirement(true);
        $this->setGroupParityRequirement(true);
        $this->setClassParityRequirement(true);
        $this->setCommonClassUsageThreshold(2);
        $this->setDebug(false);
    }

    /**
     * @param $file
     * @throws \Exception
     */
    public function parse($file) {

        $this->groups = [];
        $this->periods = [];
        $this->outcomes = [];
        $this->classes = [];
        $this->classUsage = [];

        $this->rendergroups = [];

        $this->file = $file;
        if (!file_exists($this->file)) {
            throw new \Exception("Supplied trial data file '{$this->file}' does not exist.");
        }

        libxml_use_internal_errors(true);
        $this->xml = simplexml_load_file($this->file);
        if ($this->xml === false) {
            echo "Supplied input file is not valid XML and could not be parsed.\n";

            $data = file_get_contents($this->file);
            $splitxml = explode("\n", $data);
            foreach (libxml_get_errors() as $error) {
                echo $this->display_xml_error($error, $splitxml);
            }
            libxml_clear_errors();
            exit;
        }

        echo "Successfully opened {$this->file}\n";
        echo "\n";

        // Get and display title
        $this->title = (string)$this->xml->brief_title;
        echo "Trial: {$this->title}\n";

        $this->summary = [
            'title' => $this->title,
            'official' => (string)$this->xml->official_title,
            'url' => (string)$this->xml->required_header->url,
            'nct' => (string)$this->xml->id_info->nct_id,
            'started' => (string)$this->xml->start_date,
            'finished' => (string)(string)$this->xml->completion_date,
        ];

        // Get groups/arms
        echo "Groups:\n";
        foreach ($this->xml->xpath("clinical_results/participant_flow/group_list/group") as $group) {
            $groupID = (string)$group['group_id'];
            $groupTitle = (string)$group->title;
            $this->groups[$groupID] = $groupTitle;
            echo "  [{$groupID}] {$groupTitle}\n";
        }
        echo "\n";

        // Get periods
        echo "Periods:\n";
        $periodCounter = 0;
        foreach ($this->xml->xpath("clinical_results/participant_flow/period_list/period") as $period) {
            $periodCounter++;
            $periodID = "PERIOD{$periodCounter}";
            $periodTitle = (string)$period->title;
            $this->periods[$periodCounter] = $periodTitle;
            echo "  [{$periodID}] {$periodTitle}\n";

            // Period groups
            $periodGroups = [];
            foreach ($period->xpath("milestone//participants") as $periodGroup) {
                $periodGroupID = (string)$periodGroup['group_id'] ?? null;
                if ($periodGroup) {
                    $periodGroups[$periodGroupID] = true;
                }
            }

            ksort($periodGroups, SORT_NATURAL);
            foreach (array_keys($periodGroups) as $groupID) {
                $groupTitle = $this->groups[$groupID];
                echo "    + {$groupTitle}\n";
            }

        }
        echo "\n";

        // Get outcomes
        unset($outcome);
        foreach ($this->xml->xpath("clinical_results/outcome_list/outcome") as $outcome) {
            $outcomeType = strtolower($outcome->type);
            $outcomeTitle = (string)$outcome->title;
            $outcomeDescription = (string)$outcome->description;
            $outcomePopulation = (string)$outcome->population;
            $outcomeTimeframe = (string)$outcome->time_frame;

            $measure = $outcome->measure;

            $outcomeMeasureUnits = (string)$measure->units;
            $outcomeMeasureParam = (string)$measure->param;
            $outcomeMeasureDispersion = strtolower($measure->dispersion ?? 'none');

            $outcomeID = 'outcome-' . hash('sha256', strtolower(implode('-', [
                $outcomeType,
                $outcomeTitle,
                $outcomeTimeframe,
                uniqid()
            ])));

            // Outcome groups
            $outcomeGroups = [];
            unset($outcomeGroup);
            foreach ($outcome->xpath(".//group_list/group") as $outcomeGroup) {
                $outcomeGroupID = (string)$outcomeGroup['group_id'] ?? null;
                $outcomeGroupTitle = (string)$outcomeGroup->title ?? null;
                $outcomeGroups[$outcomeGroupID] = $outcomeGroupTitle;
            }

            // Outcome classes
            $outcomeClasses = [];
            $rawData = [];
            $classKeys = [];
            $groupWidth = 1;
            unset($outcomeClass);
            foreach ($measure->xpath(".//class_list/class") as $outcomeClass) {

                $outcomeClassTitle = (string)$outcomeClass->title ?? null;
                if (!$outcomeClassTitle) {
                    $outcomeClassTitle = "Scalar";
                }

                $outcomeClassID = 'class-' . hash('sha256', strtolower($outcomeClassTitle));
                $this->classes[$outcomeClassID] = $outcomeClassTitle;
                $outcomeClasses[$outcomeClassID] = $outcomeClassTitle;

                if (!array_key_exists($outcomeClassID, $this->classUsage)) {
                    $this->classUsage[$outcomeClassID] = [];
                }
                $this->classUsage[$outcomeClassID][] = $outcomeID;

                $classData = [
                    'class' => $outcomeClassTitle
                ];
                unset($measurement);
                foreach ($outcomeClass->xpath(".//measurement_list/measurement") as $measurement) {
                    $classGID = (string)$measurement['group_id'];
                    $classData['data'][$classGID] = ((array)$measurement->attributes())['@attributes'];
                    $thisGroupWidth = count($classData['data'][$classGID])-1;
                    $groupWidth = $thisGroupWidth > $groupWidth ? $thisGroupWidth : $groupWidth;
                    if (empty($classKeys)) {
                        $classKeys = array_keys($classData['data'][$classGID]);
                    }
                }
                $rawData[] = $classData;
            }

            // Analyzed
            $analyzed = $measure->xpath("analyzed_list/analyzed")[0];
            $analyzedData = [
                'units' => (string)$analyzed->units,
                'scope' => (string)$analyzed->scope,
                'data' => []
            ];
            unset($count);
            foreach ($analyzed->xpath(".//count") as $count) {
                $analyzedGID = (string)$count['group_id'];
                $analyzedData['data'][$analyzedGID] = ((array)$count->attributes())['@attributes'];
            }

            $this->outcomes[$outcomeID] = [
                'id' => $outcomeID,
                'type' => $outcomeType,
                'title' => $outcomeTitle,
                'description' => $outcomeDescription,
                'population' => $outcomePopulation,
                'comparetitle' => $outcomeType .'/'.$outcomeMeasureUnits."/".($outcomeTitle),
                'timeframe' => $outcomeTimeframe,
                'measure' => [
                    'units' => $outcomeMeasureUnits,
                    'param' => $outcomeMeasureParam,
                    'dispersion' => $outcomeMeasureDispersion,
                    'width' => $groupWidth,
                    'keys' => $classKeys,
                    'analyzed' => $analyzedData,
                    'raw' => $rawData
                ],
                'groups' => $outcomeGroups,
                'classes' => $outcomeClasses
            ];

            // Raw
        }

        echo"Associate similar outcomes...\n";

        $this->similarByTitles();
        $this->similarByClassUsage();
        echo "\n";

        // Collect single outcomes

        foreach ($this->outcomes as $outcome) {
            if (empty($outcome['rendergroup'])) {
                $this->associate([$outcome['id']]);
            }
        }

        // Show

        if ($this->debug) {
            print_r($this->outcomes);
            echo "\n";
        }

        printf("Outcomes (%d):\n", count($this->outcomes));
        reset($this->outcomes);
        foreach ($this->outcomes as $outcome) {
            $this->renderOutcome($outcome);
        }

        echo "\n";
    }

    /**
     * Set similarity comparison threshold
     * @param $threshold
     */
    public function setSimilarityThreshold($threshold) {
        $this->threshold = $threshold;
    }

    public function setUnitParityRequirement(bool $required) {
        $this->parity['units'] = $required;
    }

    public function setGroupParityRequirement(bool $required) {
        $this->parity['groups'] = $required;
    }

    public function setClassParityRequirement(bool $required) {
        $this->parity['classes'] = $required;
    }

    public function setCommonClassUsageThreshold(int $commonality) {
        $this->parity['classusage'] = $commonality;
    }

    public function setDebug(bool $debug) {
        $this->debug = $debug;
    }

    /**
     * Try to find similar outcomes by title and units
     *
     */
    protected function similarByTitles() {
        // Connect similar outcomes
//        echo "Similarity: By Titles and Units...\n";
//        echo "\n";
        foreach ($this->outcomes as $outcomeID => &$outcome) {
//            echo "  {$outcome['title']} ({$outcome['measure']['units']})\n";

            $renderGroup = [
                $outcome['id']
            ];
            foreach ($this->outcomes as $testOutcomeID => $testOutcome) {
                similar_text($outcome['comparetitle'], $testOutcome['comparetitle'], $similarity);
                if ($similarity >= $this->threshold) {

                    // Run similarity checks

                    // Don't render with ourselves, that would be silly
                    if ($testOutcome['id'] == $outcome['id']) {
                        continue;
                    }

                    // Require units parity
                    if ($this->parity['units'] && $testOutcome['measure']['units'] != $outcome['measure']['units']) {
                        continue;
                    }

//                    echo "    {$testOutcome['title']} ({$testOutcome['measure']['units']})\n";

                    // Require group parity
                    if ($this->parity['groups'] && $testOutcome['groups'] != $outcome['groups']) {
//                        echo "    > groups mismatch\n";
//                        echo "      (".implode(',', $testOutcome['groups']).") vs. (".implode(',',$outcome['groups']).")\n\n";
                        continue;
                    }

                    // Require class parity
                    if ($this->parity['classes'] && $testOutcome['classes'] != $outcome['classes']) {
//                        echo "    > classes mismatch\n";
//                        echo "      (".implode(',', $testOutcome['classes']).") vs. (".implode(',',$outcome['classes']).")\n\n";
                        continue;
                    }

//                    echo "    > associating\n";

                    // Matched, add to render group
                    $renderGroup[] = $testOutcomeID;
                }
            }

            if (count($renderGroup) > 1) {
                $this->associate($renderGroup, $outcome['measure']['units']);
            }
//            echo "\n";
        }
//        echo "\n";
    }

    /**
     * Try to find similar outcomes by class usage
     *
     */
    protected function similarByClassUsage() {
//        echo "Similarity: By Class Usage...\n";
//        echo "\n";

        $shareTracker = [];

        foreach ($this->classUsage as $classID => $classOutcomes) {

            $class = $this->classes[$classID];
//            echo "  {$class}:\n";

            // No overlap with other outcomes
            if (count($classOutcomes) == 1) {
//                echo "    skipped, unique\n\n";
                continue;
            }

            $renderGroups = [];
            foreach ($classOutcomes as $outcomeID) {
                $outcomeRenderGroupID = $this->outcomes[$outcomeID]['rendergroup'] ?? null;
                if ($outcomeRenderGroupID) {
                    $renderGroups[$outcomeRenderGroupID] = ($renderGroups[$outcomeRenderGroupID] ?? 0) + 1;
                }
            }

            // Already matched with all similars
            if (count($renderGroups) == 1 && array_sum($renderGroups)) {
//                echo "    < skipped, fully matched\n\n";
                continue;
            }

            foreach ($classOutcomes as $outcomeID) {
                $outcome = $this->outcomes[$outcomeID];
                $outcomeRenderGroup = $outcome['rendergroup'] ?? 'no group';
//                echo "    [{$outcomeRenderGroup}] {$outcome['title']}\n";

                if (empty($shareTracker[$outcomeID])) {
                    $shareTracker[$outcomeID] = [];
                }

                // Track common class usage
                foreach ($classOutcomes as $sharedOutcomeID) {
                    if ($sharedOutcomeID != $outcomeID) {
                        $shareTracker[$outcomeID][$sharedOutcomeID] = ($shareTracker[$outcomeID][$sharedOutcomeID] ?? 0) + 1;
                    }
                }
            }
//            echo "\n";

        }
//        echo "\n";

//        echo "Associate: By Class Usage\n";
//        echo "\n";
        foreach ($shareTracker as $outcomeID => $sharedUsageIDs) {
            // No commons
            if (!count($sharedUsageIDs)) {
                continue;
            }

            $outcome = $this->outcomes[$outcomeID];
//            echo "  {$outcome['title']}\n";
//            array_walk($outcome['classes'], function ($class){ echo "    | {$class}\n";});
            foreach ($sharedUsageIDs as $sharedUsageID => $count) {
                $commonOutcome = $this->outcomes[$sharedUsageID];
//                echo "    {$count} classes in common with {$commonOutcome['title']}\n";
                if ($count >= $this->parity['classusage']) {
//                    array_walk($commonOutcome['classes'], function ($class){ echo "      | {$class}\n";});
//                    echo "      > associating\n";
                    $this->associate([$outcomeID, $sharedUsageID]);
                }
            }
        }
    }

    public function renderOutcome($outcome) {
        printf("  [%9s] %s\n", $outcome['type'], $outcome['title']);
        printf("%14s: %s\n", 'timeframe', $outcome['timeframe']);
        printf("%14s: %s\n", 'units', $outcome['measure']['units']);
        printf("%14s: %s\n", 'param', $outcome['measure']['param']);
        printf("%14s: %s\n", 'disperson', $outcome['measure']['dispersion']);
        echo "\n";

        foreach ($outcome['groups'] as $groupID => $groupTitle) {
            echo "    + {$groupTitle}\n";
        }
        echo "\n";

        foreach ($outcome['classes'] as $classID => $classTitle) {
            echo "    - {$classTitle}\n";
        }
        echo "\n";

        if (!empty($outcome['rendergroup'])) {
            echo "    Render with:\n";
            echo "      group: {$outcome['rendergroup']}\n";
            $groupTitles = [];
            foreach ($this->rendergroups[$outcome['rendergroup']]['outcomes'] as $similarOutcomeID) {
                $groupTitles[] = $this->outcomes[$similarOutcomeID]['title'];
            }
            sort($groupTitles, SORT_NATURAL);
            array_walk($groupTitles, function ($title){ echo "      {$title}\n";});
            echo "\n";
        }
    }

    public function associate(array $outcomeIDs, string $commonUnits = null) {
        sort($outcomeIDs, SORT_NATURAL);
        $renderGroupID = 'rendergroup-'.hash('sha256', implode("\n", $outcomeIDs).uniqid());

        $this->rendergroups[$renderGroupID] = [
            'id' => $renderGroupID,
            'outcomes' => [],
            'classes' => []
        ];

        if ($commonUnits) {
            $this->rendergroups[$renderGroupID]['commonunits'] = $commonUnits;
        }

        // Cannibalize existing render groups with these outcome ids
        $defunctRenderGroups = [];
        foreach ($outcomeIDs as $outcomeID) {
            $existingGroupID = $this->outcomes[$outcomeID]['rendergroup'] ?? null;
            if ($existingGroupID && $existingGroupID != $renderGroupID) {
                $defunctRenderGroups[] = $existingGroupID;
            }

            $this->associateWithRenderGroup($outcomeID, $renderGroupID);
        }

        foreach ($defunctRenderGroups as $defunctRenderGroupID) {
            if (!array_key_exists($defunctRenderGroupID, $this->rendergroups)) {
                continue;
            }

            $defunctGroupOutcomes = $this->rendergroups[$defunctRenderGroupID]['outcomes'];
            foreach ($defunctGroupOutcomes as $defunctOutcomeID) {
                $this->associateWithRenderGroup($defunctOutcomeID, $renderGroupID);
            }
            unset($this->rendergroups[$defunctRenderGroupID]);
        }

        return $renderGroupID;
    }


    public function associateWithRenderGroup($outcomeID, $renderGroupID) {
        $this->rendergroups[$renderGroupID]['outcomes'] = array_unique(array_merge($this->rendergroups[$renderGroupID]['outcomes'], [$outcomeID]));
        $this->outcomes[$outcomeID]['rendergroup'] = $renderGroupID;

        $this->rendergroups[$renderGroupID]['classes'] = array_unique(array_merge($this->rendergroups[$renderGroupID]['classes'], array_values($this->outcomes[$outcomeID]['classes'])));
        //natsort($this->rendergroups[$renderGroupID]['classes']);
        if (in_array('Baseline', $this->rendergroups[$renderGroupID]['classes'])) {
            $baselineIndex = array_search('Baseline', $this->rendergroups[$renderGroupID]['classes']);
            if ($baselineIndex != 0) {
                unset($this->rendergroups[$renderGroupID]['classes'][$baselineIndex]);
                array_unshift($this->rendergroups[$renderGroupID]['classes'], 'Baseline');
            }
        }
    }

    /**
     * XML parsing error output
     *
     * @param $error
     * @param $xml
     * @return string
     */
    private function display_xml_error($error, $xml) {
        $return  = $xml[$error->line - 1] . "\n";
        $return .= str_repeat('-', $error->column) . "^\n";

        switch ($error->level) {
            case LIBXML_ERR_WARNING:
                $return .= "Warning $error->code: ";
                break;
            case LIBXML_ERR_ERROR:
                $return .= "Error $error->code: ";
                break;
            case LIBXML_ERR_FATAL:
                $return .= "Fatal Error $error->code: ";
                break;
        }

        $return .= trim($error->message) .
            "\n  Line: $error->line" .
            "\n  Column: $error->column";

        if ($error->file) {
            $return .= "\n  File: $error->file";
        }

        return "$return\n\n--------------------------------------------\n\n";
    }

    public function save() {
        echo "Saving to disk\n";
        $saver = new Save();
        $saver->build($this->summary, $this->rendergroups, $this->outcomes);

        $file = $this->file.".xlsx";
        echo "  {$file}\n";
        $saver->save($file);
    }

}
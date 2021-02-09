#!/usr/bin/env php
<?php

/**
 * General purpose XML parser and formatter for studies published on clinicaltrials.gov
 *
 * @license MIT
 * @copyright 2021 Tim Gunter
 * @author Tim Gunter <gunter.tim@gmail.com>
 */

use \Kaecyra\ClinicalTrials\ClinicalTrialParser;

// Switch to root directory
chdir(dirname($argv[0]));

// Include the core autoloader.

$paths = [
    getcwd().'/vendor/autoload.php',
    __DIR__.'/../vendor/autoload.php',  // locally
    __DIR__.'/../../../autoload.php',   // dependency
];
foreach ($paths as $path) {
    if (file_exists($path)) {
        require_once $path;
        break;
    }
}

$cmd = basename($argv[0]);
if ($argc < 2) {
    echo "Usage: {$cmd} <xml file>\n";
    exit;
}

// Get input file
$f = $argv[1];

$trial = new ClinicalTrialParser();
$trial->setSimilarityThreshold(70);
$trial->setClassParityRequirement(false);
$trial->setCommonClassUsageThreshold(2);
$trial->parse($f);
$trial->save();
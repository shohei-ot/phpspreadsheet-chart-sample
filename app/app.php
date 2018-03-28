<?php
require dirname(__FILE__).'/../vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\Writer\Exception as WriterException;

use PhpOffice\PhpSpreadsheet\Chart\Chart;
use PhpOffice\PhpSpreadsheet\Chart\Title;
use PhpOffice\PhpSpreadsheet\Chart\Layout;
use PhpOffice\PhpSpreadsheet\Chart\Legend;
use PhpOffice\PhpSpreadsheet\Chart\PlotArea;
use PhpOffice\PhpSpreadsheet\Chart\Axis;
use PhpOffice\PhpSpreadsheet\Chart\DataSeries;
use PhpOffice\PhpSpreadsheet\Chart\DataSeriesValues;

$book = new Spreadsheet();
$sheet = $book->getActiveSheet();

// chart

// Chart Title
$titleLayoutOpts = [
  'layoutTarget' => '',
  'xMode' => '',
  'yMode' => '',
  'x' => 100,
  'y' => 0,
  'w' => 100,
  'h' => 60,
];
$chartTitleLayout = new Layout($titleLayoutOpts);
$chartTitle = new Title('chart-title', $chartTitleLayout);

// Chart Legend
$legendLayoutOpts = [
  'layoutTarget' => '',
  'xMode' => '',
  'yMode' => '',
  'x' => 0,
  'y' => 60,
  'w' => 80,
  'h' => 340,
];
$chartLegendLayout = new Layout($legendLayoutOpts);
$isOverlay = true;
$chartLegend = new Legend('r', $chartLegendLayout, $isOverlay);

// PlotArea
$plotLayoutOpts = [
  'layoutTarget' => '',
  'xMode' => '',
  'yMode' => '',
  'x' => 80,
  'y' => 60,
  'w' => 140,
  'h' => 340,
];
$plotAreaLayout = new Layout($plotLayoutOpts);

// ProtSeries
$labels = ['東京都', '大阪府', '京都府'];
$data = [100, 80, 40];
// $datasets = [
//   [
//     'label' => '電車',
//     'data' => [100, 80, 60],
//     'backgroundColor' => '#ff0000'
//   ],
//   [
//     'label' => 'バス',
//     'data' => [20, 30, 40],
//     'backgroundColor' => '#00ff00'
//   ],
//   [
//     'label' => 'その他',
//     'data' => [50, 40, 80],
//     'backgroundColor' => '#0000ff'
//   ],
// ];
$plotSeries = []; // DataSeriesの配列

// $plotType = 'bar'; // あってるか分からない
$plotType = DataSeries::TYPE_BARCHART;
// $plotGrouping = null;
$plotGrouping = DataSeries::GROUPING_STACKED;
$plotOrder = [];

$plotLabels = [];
$pl1 = new DataSeriesValues('String', 'plot-label-1', null, 0, $labels);
$plotLabels[] = $pl1;

$plotValues = [];
$fillColor = '#ffaaaa';
$pv1 = new DataSeriesValues('Number', 'plot-value-1', null, 0, $data, null, $fillColor);
$plotValues[] = $pv1;

$dataSeries = new DataSeries(
  $plotType,
  $plotGrouping,
  $plotOrder,
  $plotLabels,
  []
);


$plotArea = new PlotArea($plotAreaLayout, $plotSeries);

$plotVisibleOnly = true;
$displayBlanksAs = '0';

// xAxis Title
$xAxisLabelOpts = [
  'layoutTarget' => '',
  'xMode' => '',
  'yMode' => '',
  'x' => 110,
  'y' => 370,
  'w' => 110,
  'h' => 30,
];
$xAxisLabelLayout = new Layout($xAxisLabelOpts);
$xAxisLabel = new Title('xAxisLabel', $xAxisLabelLayout);

// yAxis Title
$yAxisLabelOpts = [
  'layoutTarget' => '',
  'xMode' => '',
  'yMode' => '',
  'x' => 80,
  'y' => 60,
  'w' => 30,
  'h' => 310,
];
$yAxisLabelLayout = new Layout($yAxisLabelOpts);
$yAxisLabel = new Title('yAxisLabel', $yAxisLabelLayout);

// xAxis
$xAxis = new Axis();
$xAxis->setAxisOptionsProperties('xAxis!');

// yAxis
$yAxis = new Axis();
$yAxis->setAxisOptionsProperties('yAxis!');

// Chart
$chart = new Chart(
  'chart-name',
  $chartTitle,
  $chartLegend,
  $plotArea,
  $plotVisibleOnly,
  $displayBlanksAs,
  $xAxisLabel,
  $yAxisLabel,
  $xAxis,
  $yAxis
);

// chart/end?


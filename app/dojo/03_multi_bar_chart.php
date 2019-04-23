<?php

require_once "./common.php";

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Reader\Xlsx as Reader;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx as Writer;
use PhpOffice\PhpSpreadsheet\Writer\Exception as WriterException;

use PhpOffice\PhpSpreadsheet\Chart\DataSeries;
use PhpOffice\PhpSpreadsheet\Chart\DataSeriesValues;
use PhpOffice\PhpSpreadsheet\Chart\PlotArea;
use PhpOffice\PhpSpreadsheet\Chart\Title;
use PhpOffice\PhpSpreadsheet\Chart\Legend;
use PhpOffice\PhpSpreadsheet\Chart\Chart;

$tmpPath = dirname(__FILE__).'/tmp';
$srsFilePath = dirname(__FILE__).'/src/03_multi_bar_chart_base.xlsx';

$reader = new Reader();
$spreadSheet = $reader->load($srsFilePath);

$sheet = $spreadSheet->getActiveSheet();

// 横軸の各要素のラベル達
$xAxisTickValues = [
    new DataSeriesValues(
        DataSeriesValues::DATASERIES_TYPE_STRING,
        "{$sheet->getTitle()}!A21:B43",
        null,
        23
    )
];

// グラフで表現するデータ
$dataSeriesValues = [
    new DataSeriesValues(
        DataSeriesValues::DATASERIES_TYPE_NUMBER,
        "{$sheet->getTitle()}!C21:C43"
    ),
    new DataSeriesValues(
        DataSeriesValues::DATASERIES_TYPE_NUMBER,
        "{$sheet->getTitle()}!D21:D43"
    )
];

// データに対応する凡例
$plotLegendLabels = [
    new DataSeriesValues(
        DataSeriesValues::DATASERIES_TYPE_STRING,
        "{$sheet->getTitle()}!C20"
    ),
    new DataSeriesValues(
        DataSeriesValues::DATASERIES_TYPE_STRING,
        "{$sheet->getTitle()}!D20"
    ),
];

$series = new DataSeries(
    DataSeries::TYPE_BARCHART,
    DataSeries::GROUPING_STANDARD,
    range(0, count($dataSeriesValues) - 1),
    $plotLegendLabels,
    $xAxisTickValues,
    $dataSeriesValues
);

$series->setPlotDirection(DataSeries::DIRECTION_COL);

$plotArea = new PlotArea(null, [$series]);

$title = new Title('ここにグラフタイトル');

$legend = new Legend(Legend::POSITION_RIGHT,null,false);

$chart = new Chart(
    'bar chart',
    $title,
    $legend,
    $plotArea
);

$chart->setTopLeftPosition('F16')->setBottomRightPosition('S43');

$sheet->addChart($chart);

$writer = new Writer($spreadSheet);
$writer->setIncludeCharts(true);
$filename = '03_multi_bar_chart_output.xlsx';
$filepath = $tmpPath.'/'.$filename;

if (file_exists($filepath)) {
    unlink($filepath);
}
$writer->save($filepath);

try{
    $size = filesize($filepath);
    $src = fopen($filepath, 'r');

    header('Content-disposition: attachment; filename='.$filename);
    header('Content-type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    header('Content-Length: '.$size);
    header('Content-Transfer-Encoding: binary');
    header('Cache-Control: must-revalidate');
    header('Pragma: public');

    ob_start();

    while(!feof($src)){
        echo fread($src, 1024);
    }

    fclose($src);
    ob_end_flush();

    unlink($filepath); // tmpのファイルを削除
}catch(\Exception $e){
    if(isset($src)){
        fclose($src);
    }
    ob_end_clean();
    echo "ダウンロード処理に失敗";
    echo "<hr>";
    echo $e->getMessage();
    exit;
}
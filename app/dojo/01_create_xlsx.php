<?php

require_once "./common.php";

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\Writer\Exception as WriterException;

$tmpPath = dirname(__FILE__).'/tmp';

$spreadsheet = new Spreadsheet();

$spreadsheet->getProperties()
    ->setTitle('タイトル')
    ->setSubject('サブタイトル')
    ->setCreator('作成者')
    ->setCompany('会社名')
    ->setManager('管理者')
    ->setCategory('分類')
    ->setDescription('コメント')
    ->setKeywords('キーワード');

$sheet = $spreadsheet->getActiveSheet();
$sheet->setCellValue('A1', 'A1です!!'); // 内部では `getCell()->setValue()` をやってるっぽい
$sheet->getCell('A2')->setValue('A2です');

$sheet->setCellValue('A4', '勤怠管理');
$data = [
  ['日付', '出勤', '退勤', '勤務時間', 'メモ'],
  ['2018/03/27', '10:00', '19:00', null, '関数って使えるのかな'],
  ['2018/03/28', '10:00', '19:00', null, '=1+2'],
  ['2018/03/29', '10:00', '19:00', null, '=2+3'],
];
$sheet->fromArray($data, null, 'A5', true);

$writer = new Xlsx($spreadsheet);
$xlsFileName = 'test_01.xlsx';
$filePath = $tmpPath.'/'.$xlsFileName;

try{
  if(file_exists($filePath)){
    unlink($filePath);
  }
}catch(\Exception $e){
  echo $e->getMessage();
  exit;
}

try{
  $writer->save($filePath);
}catch(WriterException $e){
  echo 'saveに失敗';
  echo '<hr>';
  echo $e->getMessage();
  exit;
}

try{
  $size = filesize($filePath);
  $src = fopen($filePath, 'r');

  header('Content-disposition: attachment; filename='.$xlsFileName);
  header('Content-type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  header('Content-Length: '.$size);
  header('Content-Transfer-Encoding: binary');
  header('Cache-Control: must-revalidate');
  header('Pragma: public');

  ob_start();

  while(!feof($src)){
    echo fread($src, 4);
  }

  fclose($src);
  ob_end_flush();

  unlink($filePath); // tmpのファイルを削除
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
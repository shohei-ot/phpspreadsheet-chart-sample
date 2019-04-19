<?php

require_once "./common.php";

use PhpOffice\PhpSpreadsheet\Writer\Xlsx as Writer;
use PhpOffice\PhpSpreadsheet\Reader\Xlsx as Reader;

$tempPath = dirname(__FILE__).'/template';
$tempFilePath = $tempPath.'/template.xlsx';

$tmpPath = dirname(__FILE__).'/tmp';
$tmpFileName = 'test_02.xlsx';
$tmpFilePath = $tmpPath.'/'.$tmpFileName;


$reader = new Reader();
$book = $reader->load($tempFilePath);

$sheet = $book->getActiveSheet();

$students = [
  [null, '田中太郎', '', 20],
  [null, '田中太郎', '', 21],
  [null, '田中太郎', '', 22],
  [null, '田中太郎', '', 23],
  [null, '田中太郎', '', 24],
  [null, '田中太郎', '', 25],
  [null, '田中太郎', '', 26],
  [null, '田中太郎', '', 27],
  [null, '田中太郎', '', 28],
];
$students = array_map(function($row){
  $row[1] = $row[1].' '.rand(1,100);
  $row[2] = rand(0,1) ? '男' : '女';
  $row[3] = $row[3] + rand(5, 20);
  return $row;
}, $students);

$sheet->fromArray($students, null, 'A2', true);

$writer = new Writer($book);


try{
  $writer->save($tmpFilePath);
}catch(WriterException $e){
  echo 'saveに失敗';
  echo '<hr>';
  echo $e->getMessage();
  exit;
}

try{
  $size = filesize($tmpFilePath);
  $src = fopen($tmpFilePath, 'r');

  header('Content-disposition: attachment; filename='.$tmpFileName);
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

  unlink($tmpFilePath); // tmpのファイルを削除
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
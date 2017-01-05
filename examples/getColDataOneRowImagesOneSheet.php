<?php
/**
 * Created by PhpStorm.
 * User: lejianwen
 * Date: 2017/1/4
 * Time: 18:12
 * QQ: 84855512
 */
//按列 读某个Sheet中所有数据，并上传某行的图片
//一行一组数据
require_once __DIR__ .'/../../../autoload.php';
$file = '/test.xlsx';
$image_path = '/data/upload/';
$excel = new \Ljw\Excel\Excel();
$excel->loadFile($file);
$excel->loadDataFromSheetCol(1);
//将sheet中某一列的图片存到$path中
$excel->saveRowImagesFromSheet(1, '5', $image_path);
$excel->combineExcelData();
$data = $excel->getSheetData(1);
var_dump($data);
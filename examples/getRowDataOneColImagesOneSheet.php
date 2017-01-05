<?php
/**
 * Created by PhpStorm.
 * User: lejianwen
 * Date: 2017/1/4
 * Time: 18:12
 * QQ: 84855512
 */
//按行 读某个Sheet中所有数据，并上传某列的图片
//一行一组数据
require_once __DIR__ .'/../../../autoload.php';
$file = '/test.xlsx';
$image_path = '/data/upload/';
$excel = new \Ljw\Excel\Excel();
$excel->loadFile($file);
$excel->loadDataFromSheetRow(0);
//将sheet中某一列的图片存到$path中
$excel->saveColImagesFromSheet(0, 'B', $image_path);
$excel->combineExcelData();
$data = $excel->getSheetData(0);
var_dump($data);
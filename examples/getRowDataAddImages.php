<?php
/**
 * Created by PhpStorm.
 * User: lejianwen
 * Date: 2017/1/4
 * Time: 18:12
 * QQ: 84855512
 */
//按行 读某个Excel中所有数据，并上传图片
//一行一组数据
require_once __DIR__ .'/../../../autoload.php';
$file = '/test.xlsx';
$image_path = '/data/upload/';
$excel = new \Ljw\Excel\Excel();
$excel->loadFile($file);
//请保证每个sheet的格式一致否则请用单一sheet数据合并
$excel->loadDataFromExcelRow();
//将excel中的图片存到$path中
$excel->saveImagesFromExcel($image_path);
$excel->combineExcelData();
$data = $excel->getExcelData();
var_dump($data);
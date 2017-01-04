<?php
/**
 * Created by PhpStorm.
 * User: lejianwen
 * Date: 2017/1/4
 * Time: 18:12
 * QQ: 84855512
 */
//按列 读某个Sheet中所有数据，并上传图片
//一列一组数据
require_once __DIR__ .'/../../../autoload.php';
$file = '/test.xlsx';
$image_path = '/data/upload/';
$excel = new \Ljw\Excel\Excel();
$excel->loadFile($file);
//读取第2个sheet的数据，按列组合
$data = $excel->loadDataFromSheetCol(1);
//将excel中的图片存到$path中
$imageData = $excel->saveImagesFromSheet(1, $image_path);
//2行的是图片,组合
$re = $excel->combineSheetData($data, $imageData, '2');
//3行的是图片,组合
$re = $excel->combineSheetData($re, $imageData, '7');
var_dump($re);
<?php
/**
 * Created by PhpStorm.
 * User: lejianwen
 * Date: 2017/1/4
 * Time: 18:12
 * QQ: 84855512
 */
//按行 读某个Sheet中所有数据，并上传图片
//一行一组数据
require_once __DIR__ .'/../../../autoload.php';
$file = '/test.xlsx';
$image_path = '/data/upload/';
$excel = new \Ljw\Excel\Excel();
$excel->loadFile($file);
$data = $excel->loadDataFromSheetRow(0);
//将excel中的图片存到$path中
$imageData = $excel->saveImagesFromSheet(0, $image_path);
//B列的是图片，组合
$re = $excel->combineSheetData($data, $imageData, 'B');
//G列的是图片，组合
$re = $excel->combineSheetData($re, $imageData, 'G');
var_dump($re);
<?php
/**
 * Created by PhpStorm.
 * User: lejianwen
 * Date: 2017/1/4
 * Time: 18:12
 * QQ: 84855512
 */
//按行 读某个Excel中所有数据
//一行一组数据
require_once __DIR__ .'/../../../autoload.php';
$file = '/test.xlsx';
$excel = new \Ljw\Excel\Excel();
$excel->loadFile($file);
$excel->loadDataFromExcelRow();
$data = $excel->getExcelData();
var_dump($data);
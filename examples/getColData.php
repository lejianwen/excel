<?php
/**
 * Created by PhpStorm.
 * User: lejianwen
 * Date: 2017/1/4
 * Time: 18:12
 * QQ: 84855512
 */
//按列 读excel中所有数据
//一列一组数据
require_once __DIR__ .'/../../../autoload.php';
$file = '/test.xlsx';
$excel = new \Ljw\Excel\Excel();
$excel->loadFile($file);
$data = $excel->loadDataFromExcelCol();
var_dump($data);
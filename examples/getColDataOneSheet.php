<?php
/**
 * Created by PhpStorm.
 * User: lejianwen
 * Date: 2017/1/4
 * Time: 18:12
 * QQ: 84855512
 */

//按列 读某个Sheet中所有数据
//一列一组数据
require_once __DIR__ .'/../../../autoload.php';
$file = '/test.xlsx';
$excel = new \Ljw\Excel\Excel();
$excel->loadFile($file);
//读取第一个sheet
$data = $excel->loadDataFromSheetCol(0);
//读取第二个sheet
$data2 = $excel->loadDataFromSheetCol(1);
var_dump($data);
<?php
/**
 * Created by PhpStorm.
 * User: lejianwen
 * Date: 2016/11/25
 * Time: 12:25
 * QQ: 84855512
 */
namespace Ljw\Excel;

use \Exception;

class Excel
{
    protected $objPHPExcel;
    protected $excelData;
    protected $images;

    public function __construct()
    {
        $this->excelData = [];
        $this->images = [];
    }

    public function getExcelData()
    {
        return $this->excelData;
    }


    public function getSheetData($i)
    {
        return $this->excelData[$i];
    }

    /**
     * @return array
     */
    public function getImages()
    {
        return $this->images;
    }

    public function loadFile($filename)
    {
        if ($this->objPHPExcel)
            return $this->objPHPExcel;
        if (!is_file($filename))
            throw new Exception('文件不存在');
        $this->objPHPExcel = \PHPExcel_IOFactory::load($filename);
        return $this->objPHPExcel;
    }

    /**按行分组读取sheet数据
     * @param $i int sheet的索引值
     */
    public function loadDataFromSheetRow($i)
    {
        if(!empty($this->excelData[$i]))
            return;
        $objWorksheet = $this->objPHPExcel->getSheet($i);
        //获取总行数
        $highestRow = $objWorksheet->getHighestRow();
        //获取总列数
        $highestColumn = $objWorksheet->getHighestColumn();
        $highestColumnIndex = \PHPExcel_Cell::columnIndexFromString($highestColumn);
        for ($row = 1; $row <= $highestRow; $row++)
        {
            for ($col = 0; $col < $highestColumnIndex; $col++)
            {
                $col_str = \PHPExcel_Cell::stringFromColumnIndex($col);
                $this->excelData[$i][$row][$col_str] = (string)$objWorksheet->getCell($col_str . $row)->getValue();
                //公式值
                //$objWorksheet->getCell($col_str . $row)->getCalculatedValue();
            }
        }
    }

    public function sheetToArray()
    {

    }

    /**按列分组读取sheet数据
     * @param $i int sheet的索引值
     */
    public function loadDataFromSheetCol($i)
    {
        if(!empty($this->excelData[$i]))
            return;
        $objWorksheet = $this->objPHPExcel->getSheet($i);
        //获取总行数
        $highestRow = $objWorksheet->getHighestRow();
        //获取总列数
        $highestColumn = $objWorksheet->getHighestColumn();
        $highestColumnIndex = \PHPExcel_Cell::columnIndexFromString($highestColumn);
        for ($col = 0; $col < $highestColumnIndex; $col++)
        {
            $col_str = \PHPExcel_Cell::stringFromColumnIndex($col);
            for ($row = 1; $row <= $highestRow; $row++)
            {
                $this->excelData[$i][$col_str][$row] = (string)$objWorksheet->getCell($col_str . $row)->getValue();
            }
        }
    }

    /**按行分组读取整个excel数据，包括所有的sheet
     * @throws Exception
     */
    public function loadDataFromExcelRow()
    {
        if (!$this->objPHPExcel)
            throw new Exception('PHPExcel is not be load');
        $sheetCount = $this->objPHPExcel->getSheetCount();
        for ($i = 0; $i < $sheetCount; $i++)
        {
            $this->loadDataFromSheetRow($i);
        }
    }

    /**按列分组读取excel数据，包括所有的sheet
     * @throws Exception
     */
    public function loadDataFromExcelCol()
    {
        if (!$this->objPHPExcel)
            throw new Exception('PHPExcel is not be load');
        $sheetCount = $this->objPHPExcel->getSheetCount();
        for ($i = 0; $i < $sheetCount; $i++)
        {
            $this->loadDataFromSheetCol($i);
        }
    }

    /**合并数据数组和图片
     */
    public function combineExcelData()
    {
        foreach ($this->excelData as $i => $sheetData)
        {
            foreach ($sheetData as $key => $_data)
            {
                foreach ($_data as $k => $value)
                {
                    if($this->images[$i][$key][$k])
                        $this->excelData[$i][$key][$k] = $this->images[$i][$key][$k] ?: '';
                }
            }
        }
    }

    /**将excel中的图片保存到本地
     * @param $path string 图片保存路径
     * @throws Exception
     */
    public function saveImagesFromExcel($path)
    {
        if (!$this->objPHPExcel)
            throw new Exception('PHPExcel is not be load');
        $sheetCount = $this->objPHPExcel->getSheetCount();
        for ($i = 0; $i < $sheetCount; $i++)
        {
            $this->saveImagesFromSheet($i, $path);
        }
    }

    /**将excel中某个sheet的图片保存到本地
     * @param $i int sheet的索引值
     * @param $path
     */
    public function saveImagesFromSheet($i, $path)
    {
        if($this->images[$i])
            return;
        $sheet = $this->objPHPExcel->getSheet($i);
        foreach ($sheet->getDrawingCollection() as $key => $drawing)
        {
            //坐标 比如 A1 B2 B3等
            $xy = $drawing->getCoordinates();
            $x = preg_replace('/[A-Z]/', '', $xy);
            $y = preg_replace('/[0-9]/', '', $xy);
            $this->images[$i][$x][$y] = $this->images[$i][$x][$y] ?: $this->saveImage($drawing, $path);
            $this->images[$i][$y][$x] = $this->images[$i][$x][$y] ?: $this->saveImage($drawing, $path);
        }
    }

    /**只将excel中某个sheet的某行的图片保存到本地
     * @param $i int sheet的索引值
     * @param $row int 行
     * @param $path
     */
    public function saveRowImagesFromSheet($i, $row, $path)
    {
        $sheet = $this->objPHPExcel->getSheet($i);
        foreach ($sheet->getDrawingCollection() as $key => $drawing)
        {
            $xy = $drawing->getCoordinates();
            $x = preg_replace('/[A-Z]/', '', $xy);
            $y = preg_replace('/[0-9]/', '', $xy);
            if ($x == $row)
            {
                $this->images[$i][$x][$y] = $this->images[$i][$x][$y] ?: $this->saveImage($drawing, $path);
                $this->images[$i][$y][$x] = $this->images[$i][$x][$y] ?: $this->saveImage($drawing, $path);
            }

        }
    }

    /**只将excel中某个sheet的某列的图片保存到本地
     * @param $i int sheet的索引值
     * @param $col string 列
     * @param $path
     */
    public function saveColImagesFromSheet($i, $col, $path)
    {
        $sheet = $this->objPHPExcel->getSheet($i);
        foreach ($sheet->getDrawingCollection() as $key => $drawing)
        {
            $xy = $drawing->getCoordinates();
            $x = preg_replace('/[A-Z]/', '', $xy);
            $y = preg_replace('/[0-9]/', '', $xy);
            if ($y == $col)
            {
                $this->images[$i][$x][$y] = $this->images[$i][$x][$y] ?: $this->saveImage($drawing, $path);
                $this->images[$i][$y][$x] = $this->images[$i][$x][$y] ?: $this->saveImage($drawing, $path);
            }

        }
    }



    /**根据图片格式保存图片
     * @param $drawing \PHPExcel_Worksheet_MemoryDrawing excel的sheet对象
     * @param $path
     * @return string 文件名
     */
    protected function saveImage($drawing, $path)
    {
        $this->createDir($path);
        if ($drawing instanceof \PHPExcel_Worksheet_MemoryDrawing)
        {
            $image = $drawing->getImageResource();
            // $filename = md5(time()).$drawing->getIndexedFilename();
            $filename = md5(uniqid(time() . rand(1, 9999)) . time() . rand(1, 9999)) . $drawing->getIndexedFilename();
            $file = $path . $filename;
            $renderingFunction = $drawing->getRenderingFunction();
            switch ($renderingFunction)
            {
                case \PHPExcel_Worksheet_MemoryDrawing::RENDERING_JPEG:
                    imagejpeg($image, $file);
                    break;

                case \PHPExcel_Worksheet_MemoryDrawing::RENDERING_GIF:
                    imagegif($image, $file);
                    break;

                case \PHPExcel_Worksheet_MemoryDrawing::RENDERING_PNG:
                    imagegif($image, $file);
                    break;

                case \PHPExcel_Worksheet_MemoryDrawing::RENDERING_DEFAULT:
                    imagegif($image, $file);
                    break;
            }
            return $file;
        }
        //xlsx
        if ($drawing instanceof \PHPExcel_Worksheet_Drawing)
        {
            $excel_file = $drawing->getPath();
            //$filename = $drawing->getIndexedFilename();
            $filename = md5(uniqid(time() . rand(1, 9999)) . time() . rand(1, 9999)) . '.' . $drawing->getExtension();
            $file = $path . $filename;
            copy($excel_file, $file);
            return $file;
        }
        return '';
    }

    /**创建目录
     * @param $dirName
     * @param int $rights
     * @deprecated  Use createDir instead.
     */
    protected function mkdir_r($dirName, $rights = 0777)
    {
        if (is_dir($dirName))
            return;
        $dirs = explode('/', $dirName);
        $dir = '';
        foreach ($dirs as $part)
        {
            $dir .= $part . '/';
            if (!is_dir($dir) && strlen($dir) > 0)
                mkdir($dir, $rights);
        }
    }

    protected function createDir($destinationFolder)
    {
        if (!$destinationFolder) {
            return $this;
        }

        if (substr($destinationFolder, -1) == DIRECTORY_SEPARATOR) {
            $destinationFolder = substr($destinationFolder, 0, -1);
        }

        if (!(@is_dir($destinationFolder) || @mkdir($destinationFolder, 0777, true))) {
            throw new Exception("Unable to create directory '{$destinationFolder}'.");
        }
    }
}

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
    public function loadPhpExcel($filename)
    {
        if($this->objPHPExcel)
            return $this->objPHPExcel;
        if(!is_file($filename))
            throw new Exception('文件不存在');
        $this->objPHPExcel = \PHPExcel_IOFactory::load($filename);
        return $this->objPHPExcel;
    }

    /**按行分组读取sheet数据
     * @param $i int sheet的索引值
     * @return array
     */
    public function loadDataFromSheetRow($i)
    {
        $data = [];
        $objWorksheet = $this->objPHPExcel->getSheet($i);
        //获取总行数
        $highestRow = $objWorksheet->getHighestRow();
        //获取总列数
        $highestColumn = $objWorksheet->getHighestColumn();
        $highestColumnIndex = \PHPExcel_Cell::columnIndexFromString($highestColumn);
        for ($row = 1; $row <= $highestRow; $row++) {
            for ($col = 0; $col < $highestColumnIndex; $col++) {
                $data[$row][] = (string)$objWorksheet->getCellByColumnAndRow($col, $row)->getValue();
            }
        }
        return $data;
    }

    /**按列分组读取sheet数据
     * @param $i int sheet的索引值
     * @return array
     */
    public function loadDataFromSheetCol($i)
    {
        $data = [];
        $objWorksheet = $this->objPHPExcel->getSheet($i);
        //获取总行数
        $highestRow = $objWorksheet->getHighestRow();
        //获取总列数
        $highestColumn = $objWorksheet->getHighestColumn();
        $highestColumnIndex = \PHPExcel_Cell::columnIndexFromString($highestColumn);
        for ($col = 0; $col < $highestColumnIndex; $col++) {
            $col_str= \PHPExcel_Cell::stringFromColumnIndex($col);
            for ($row = 1; $row <= $highestRow; $row++) {
                $data[$col_str][] = (string)$objWorksheet->getCell($col_str.$row)->getValue();
            }
        }
        return $data;
    }

    /**按行分组读取整个excel数据，包括所有的sheet
     * @return array
     * @throws Exception
     */
    public function loadDataFromExcelRow()
    {
        if(!$this->objPHPExcel)
            throw new Exception('PHPExcel is not be load');
        $sheetCount = $this->objPHPExcel->getSheetCount();
        $data = [];
        for ($i = 0; $i < $sheetCount; $i++) {
            $data[$i] = $this->loadDataFromSheetRow($i);
        }
        return $data;
    }

    /**按列分组读取excel数据，包括所有的sheet
     * @return array
     * @throws Exception
     */
    public function loadDataFromExcelCol()
    {
        if(!$this->objPHPExcel)
            throw new Exception('PHPExcel is not be load');
        $sheetCount = $this->objPHPExcel->getSheetCount();
        $data = [];
        for ($i = 0; $i < $sheetCount; $i++) {
            $data[$i] = $this->loadDataFromSheetCol($i);
        }
        return $data;
    }

    /**合并数据数组和图片
     * @param $data array 整个excel数据，包括所有sheet
     * @param $imgData array 所有sheet的图片数据
     * @param $position string||int 图片所在的位置,如果是字母 则表示是一列，数字则代表是一行
     * @return array
     */
    public function combineExcelData($data, $imgData, $position)
    {
        foreach ($data as $sheet_k => $sheets)
        {
            foreach ($sheets as $key => $_data)
            {
                if(is_string($position))
                    $data[$sheet_k][$key]['_excel_image'] = $imgData[$sheet_k][$position.$key];
                elseif (is_numeric($position))
                    $data[$sheet_k][$key]['_excel_image'] = $imgData[$sheet_k][$key.($position+1)];
            }
        }
        return $data;
    }

    /**合并数据数组和图片
     * @param $data array 一个sheet的数据
     * @param $imgData array 所有sheet的图片数据
     * @param $sheet_index int sheet的索引值
     * @param $position string||int 图片所在的位置,如果是字母 则表示是一列，数字则代表是一行
     * @return array
     */
    public function combineSheetData($data, $imgData, $sheet_index, $position)
    {
        foreach ($data as $key => $_data) {
            if (is_string($position))
                $data[$key]['_excel_image'] = $imgData[$sheet_index][$position . $key];
            elseif (is_numeric($position))
                $data[$key]['_excel_image'] = $imgData[$sheet_index][$key . ($position + 1)];
        }
        return $data;
    }


    /**将excel中的图片保存到本地
     * @param $path string 图片保存路径
     * @return array
     * @throws Exception
     */
    public function saveImagesFromExcel($path)
    {
        if(!$this->objPHPExcel)
            throw new Exception('PHPExcel is not be load');
        $sheetCount = $this->objPHPExcel->getSheetCount();
        $data = [];
        for($i=0;$i<$sheetCount;$i++)
        {
            $sheet = $this->objPHPExcel->getSheet($i);
            foreach ($sheet->getDrawingCollection() as $key => $drawing)
            {
                //坐标 比如 A1 B2 B3等
                $xy = $drawing->getCoordinates();
                $data[$i][$xy] = $this->saveImagesFromExcelSheet($drawing, $path);
            }
        }
        return $data;
    }


    /**将excel中的图片保存到本地
     * @param $drawing \PHPExcel_Worksheet_MemoryDrawing excel的sheet对象
     * @param $path
     * @return string
     */
    public function saveImagesFromExcelSheet($drawing, $path)
    {
        $full_path = realpath($path);
        $this->mkdir_r($full_path);
        if ($drawing instanceof \PHPExcel_Worksheet_MemoryDrawing) {
            $image = $drawing->getImageResource();
            // $filename = md5(time()).$drawing->getIndexedFilename();
            $filename = md5(uniqid(time() . rand(1, 9999)) . time() . rand(1, 9999)) . $drawing->getIndexedFilename();
            $file = $full_path . $filename;
            $renderingFunction = $drawing->getRenderingFunction();
            switch ($renderingFunction) {
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
            return $path.$filename;
        }
        //xlsx
        if ($drawing instanceof \PHPExcel_Worksheet_Drawing) {
            $excel_file = $drawing->getPath();
            //$filename = $drawing->getIndexedFilename();
            $filename = md5(uniqid(time() . rand(1, 9999)) . time() . rand(1, 9999)) . '.' . $drawing->getExtension();
            $file = $full_path . $filename;
            copy($excel_file, $file);
            return $path.$filename;
        }
    }

    public function mkdir_r($dirName, $rights = 0777)
    {
        if(is_dir($dirName))
            return;
        $dirs = explode('/', $dirName);
        $dir = '';
        foreach ($dirs as $part) {
            $dir .= $part . '/';
            if (!is_dir($dir) && strlen($dir) > 0)
                mkdir($dir, $rights);
        }
    }
}

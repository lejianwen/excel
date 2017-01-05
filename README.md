# excel
利用phpexcel读取数据, 依赖phpexcel
# 安装
composer require "ljw/excel":"dev-master"
# 示例
    //先实例化对象
    $excel = new \Ljw\Excel\Excel();    
    $excel->loadFile($file);
## 读取数据
 ###读取整个excel数据
    //按列读取整个excel中的数据  
    //请保证每个sheet中的数据都是按列分组的，即一列一组数据
    $excel->loadDataFromExcelCol();  
    //按列读取整个excel中的数据  
    //请保证每个sheet中的数据都是按行分组的，即一行一组数据  
    $excel->loadDataFromExcelRow();
    $data = $excel->getExcelData();
 ###读取某个sheet中的数据
    //读取sheet中的数据
    $i=0; //sheet的索引值,从0开始
    //按行分组读取，即每行一组数据
    $excel->loadDataFromSheetRow($i);
    //按列分组读取，即每列一组数据
    $excel->loadDataFromSheetCol($i);
    $data = $excel->getSheetData($i);
##保存图片
 ###保存整个excel中的图片
    //图片保存路径
    $path = '/data/';
    //保存整个excel中的图片，返回图片路径数组
    $excel->saveImagesFromExcel($path);
    //只保存某个sheet的图片
    $excel->saveImagesFromSheet(1,$path);
    //如果只需要保存某列的图片或者某行的图片
    
    $excel->saveRowImagesFromSheet(1,5,$image_path);  //保存第二个sheet中第5行的图片
 
    $excel->saveColImagesFromSheet(0,'B',$image_path);   //保存第一个sheet中B列的图片
    //获取图片数据
    $images = $excel->getImages();
 ###将图片数据加入到数据中
    //图片数据加入到数据中
    $excel->combineExcelData();
    $all_data = $excel->getExcelData();
    //某个sheet的数据
    $sheet_data = $excel->getSheetData(1);



<?php
//引用套件，這邊要特別注意路徑
require_once 'reader.php';  
//建立excel檔的物件
$data = new Spreadsheet_Excel_Reader();  
//設定輸出編碼，指的是從excel讀取後再進行編碼
$data->setOutputEncoding('UTF-8');  
//載入要讀取的檔案
$data->read('../txt/txt.xlsx');  
//這行可加可不加，因為有時候會出現錯誤，錯誤的原因是因為可能在excel的表格內含有空白
error_reporting(E_ALL ^ E_NOTICE);  
//以下則是以迴圈的方式讀取資料
//下面範例則是先讀取欄位再讀取列，因此i代表列的數目，j則代表欄位的數目
      for ($i = 1; $i <= $data->sheets[0]['numRows']; $i++) {
          //如下圖因為excel的表格第一列都會寫上欄位名稱，所以這邊預設不會讀取第一列
          if($i!=1){
              for ($j = 1; $j <= $data->sheets[0]['numCols']; $j++) { 
                  $value[0] = $data->sheets[0]['cells'][$i][1];
                  $value[1] = $data->sheets[0]['cells'][$i][2];
                  $value[2] = $data->sheets[0]['cells'][$i][3];
              }
          }
     }  
     
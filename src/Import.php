<?php
/**
 * Created by PhpStorm.
 * User: yuelin
 * Date: 2018/3/7
 * Time: 上午11:57
 */

namespace Linyuee\Excel;


use Linyuee\Exception\ApiException;

class Import
{
    private $path;
    private static $support_ext = ['xls','xlsx','csv'];
    public function setPath($path)
    {
        $info = pathinfo($path);
        if (!isset($info['extension'])||!is_string($path)){
            throw new ApiException('文件不合法');
        }
        if (!in_array($info['extension'],self::$support_ext)){
            throw new ApiException('不支持该扩展文件');
        }
        $this->path =$path;
        return $this;
    }
    
    public function read()
    {
        if (empty($this->path)){
            throw new ApiException('还没有设置导入文件路径');
        }
        $type = 'Excel2007';//设置为Excel5代表支持2003或以下版本，Excel2007代表2007版
        $xlsReader = \PHPExcel_IOFactory::createReader($type);
        $xlsReader->setReadDataOnly(true);
        $xlsReader->setLoadSheetsOnly(true);
        $Sheets = $xlsReader->load($this->path);
        //开始读取上传到服务器中的Excel文件，返回一个二维数组
        $dataArray = $Sheets->getSheet(0)->toArray();
        return $dataArray;
    }
}
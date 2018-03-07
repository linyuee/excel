<?php
/**
 * Created by PhpStorm.
 * User: yuelin
 * Date: 2018/3/7
 * Time: 上午11:57
 */

namespace Linyuee\Excel;


use Linyuee\Exception\ApiException;

class Export
{
    private $fileName;
    private $title;
    private $data;
    private $startIndex = 'A';
    private $ext = 'xls';//文件扩展名
    private static $support_ext = ['xls','xlsx','csv'];
    public function setData(array $data){
        $this->data = $data;
        return $this;
    }

    public function setTitle(array $title){
        $this->title = $title;
        return $this;
    }

    public function setFileName($name){
        if (!is_string($name)){
            throw new \Linyuee\Exception\ApiException('文件名必须是字符串类型');
        }
        $this->fileName = $name;
        return $this;
    }

    public function setStartIndex($index){
        if (!preg_match('/^[A-Z]+$/', $index)){
            throw new \Linyuee\Exception\ApiException('必须是A-Z的大写字母');
        };
        $this->startIndex = $index;
        return $this;
    }

    public function setFileExt($ext){
        if (!in_array($ext,self::$support_ext)){
            throw new ApiException('不支持该扩展类型');
        }
        $this->ext = $ext;
        return $this;
    }

    public function export(){
        $sheetName = $this->fileName;
        $sheettitle = $this->title;
        $data = $this->data;
        $sheetkey = array();
        $startIndex = $this->startIndex;
        foreach ($sheettitle as $key=>$item) {
            $sheetkey[$key] = $startIndex;
            $startIndex++;
        }
        $objPHPExcel = new \PHPExcel();
        $objPHPExcel->getProperties()->setCreator("phpexcel")
            ->setLastModifiedBy("phpexcel")
            ->setTitle($sheetName)
            ->setSubject($sheetName)
            ->setDescription($sheetName)
            ->setKeywords($sheetName);
        $objPHPExcel->setActiveSheetIndex(0);
        $objSheet = $objPHPExcel->getActiveSheet();
        $objSheet->setTitle($sheetName);
        foreach ($sheettitle as $key => $value) {

            $objSheet->setCellValue($sheetkey[$key].'1',$value);
        }
        unset($value);
        $data = array_values($data);
        for ($i=0;$i<count($data);$i++) {
            foreach ($data[$i] as $key => $value) {
                if (!array_key_exists($key,$sheetkey)){
                    throw new ApiException('导出的数据'.$key.'与标题不匹配');
                }
                $objSheet->setCellValue($sheetkey[$key].($i+2),$value);
            }
            unset($value);
        }

        header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        header('Content-Disposition: attachment;filename='.$sheetName.'.'.$this->ext);
        header('Cache-Control: max-age=0');
        header ('Expires: Mon, 26 Jul 1997 05:00:00 GMT'); // Date in the past
        header ('Last-Modified: '.gmdate('D, d M Y H:i:s').' GMT'); // always modified
        header ('Cache-Control: cache, must-revalidate'); // HTTP/1.1
        header ('Pragma: public'); // HTTP/1.0

        $objWriter = \PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
        $objWriter->save('php://output');
        exit;
    }
}
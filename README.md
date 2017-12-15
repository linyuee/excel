PHPExcel工具包
===================

### 1、安装
composer require linyuee/excel

### 2、使用
```
//导出的数据
$data = array(
            ['id'=>1,'username'=>'test1','phone'=>'18881515151'],
            ['id'=>2,'username'=>'test2','phone'=>'18881515152'],
            ['id'=>3,'username'=>'test3','phone'=>'18881515153'],
        );
//标题
$sheettitle = array(  
            'id'=>'ID',
            'username'=>'答题人',
            'phone'=>'联系方式',
        );
//导出的数据和标题数组的key要一致
$excel = new \Linyuee\Excel();
        $excel->setData($data)
            ->setFileName('问卷答题情况')
            ->setTitle($sheettitle)
            ->setStartIndex('B')
            ->setFileExt('csv')  //目前支持csv，xls，xlsx
            ->export();        
```
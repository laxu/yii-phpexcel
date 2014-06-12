PHPExcel wrapper for Yii
========================

This is a simple wrapper for PHPExcel.

Configuration
-------------

Add this to your Yii application config:

```php
'components' => array(
  .....
  'excel' => array(
    'class' => '\laxu\yii_phpexcel\ExcelManager',
    'filePath' => 'app.files.excel'
  ),
),
```

* **filePath** the directory alias where you want excel files to be saved

Examples
--------

```php
public function readExcelFile($filename)
{
    $manager = Yii::app()->getComponent('excel');
    $excel = $manager->get($filename);
    return $excel->read();
}

public function writeExcelFile()
{
    $manager = Yii::app()->getComponent('excel');
    //Create empty instance
    $excel = $manager->create();
    .....
    //Note that setData doesn't care about the actual keys in the data, only the order
    $data = array(
        array(
            'header1',
            'header2',
            'header3'
        ),
        array(
            'data1',
            'data2',
            'data3',
        ),
        array(
            'id' => 1,
            'name' => 'Example',
            'moreData' => 'Something'
        )
    )
    .....
    $excel->setData($data);
    $excel->save();
}

public function downloadExcelFile($filename)
{
    $manager = Yii::app()->getComponent('excel');
    $excel = $manager->get($filename);
    $excel->download();
}
```




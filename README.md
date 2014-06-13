Yii-PHPExcel
=============
**PHPExcel wrapper for Yii**

This is a simple wrapper for the PHPExcel library. It allows reading and writing Excel files. 

Yii-PHPExcel consists of a _manager component_ and _a class representing the excel you are working with_. The manager is only used to create instances and read excel files, everything else is done using the _Excel_ class.

The _Excel_ class doesn't include many of the more advanced functions of PHPExcel. These can be used by either extending the _Excel_ class or by calling the ** getInstance ** method to directly get the PHPExcel instance of an Excel object.

Excel files are always written in _Office Open XML (.xlsx)_ format because those old Excel formats are crap.

Installation
------------

The easiest way is to install thru [Composer](https://getcomposer.org/). At the time of writing this is not yet in the Packagist repo, so you need to add the following to your _composer.json_:

```json
repositories:
[
    {
        "type": "vcs",
        "url": "https://github.com/laxu/yii-phpexcel.git"
    }
]
```
Then you can run the following command in your terminal:
```
php composer.phar require laxu/yii-phpexcel
```

Configuration
-------------

Add this to your Yii application config:

```php
'components' => array(
  .....
  'excel' => array(
    'class' => '\laxu\yii_phpexcel\ExcelManager',
    'savePath' => 'app.files.excel'
  ),
),
```

* **savePath** the directory or directory alias where you want Excel files to be saved.
* **excelClass** (optional) class to use for creating _Excel_ objects. Change this if you extend the _Excel_ class.

Examples
--------

```php
public function readExcelFile($filepath)
{
    $manager = Yii::app()->getComponent('excel');
    $excel = $manager->get($filepath);
    return $excel->read();
}

public function writeExcelFile()
{
    $manager = Yii::app()->getComponent('excel');
    //Create empty instance
    $excel = $manager->create();
    .....
    //Note that setData doesn't care about the actual keys in the data, only the order of values
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

    /* This would generate something like:
    header1 | header2 | header3
    data1 | data2 | data 3
    1 | Example | Something */
}

public function downloadExcelFile($filepath)
{
    $manager = Yii::app()->getComponent('excel');
    $excel = $manager->get($filepath);
    $excel->download();
}
```




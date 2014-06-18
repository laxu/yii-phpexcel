Yii-PHPExcel
=============
**PHPExcel wrapper for Yii**

This is a wrapper for the PHPExcel library to make it easy to use in Yii. It allows reading and writing Excel files and is made to cover the most
common use cases for importing and exporting data in Yii apps.

Yii-PHPExcel consists of a _manager component_ and _a class representing the excel you are working with_. The manager is only used to create instances and read
excel files, everything else is done using the _Excel_ class.

The _Excel_ class doesn't include many of the more advanced functions of PHPExcel. These can be ADDED by extending the _Excel_ class or
by calling the ** getInstance ** method and working directly with the PHPExcel instance of an _Excel_ object.

Excel files are always written in _Office Open XML (.xlsx)_ format because those old Excel formats are crap.

Installation
------------

The easiest way is to install through [Composer](https://getcomposer.org/). At the time of writing this is not yet in the Packagist repo, so you need to add the following to your _composer.json_:

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
  'yii-phpexcel' => array(
    'class' => '\laxu\yii_phpexcel\ExcelManager',
    'savePath' => 'app.files.excel'
  ),
),
```

* **savePath** the directory or directory alias where you want Excel files to be saved.
* **excelClass** (optional) class to use for creating _Excel_ objects. Change this if you extend the _Excel_ class.

The _Excel_ objects have their own settings to control various aspects of the document easily. Here's how you can set them:
```
$manager = Yii::app()->getComponent('yii-phpexcel');
$excel = $manager->create();
$excel->setSettings(array(
    'autoSize' => array('A', 'C', 'E'),
    'useTempDir' => true
));
```
* **autoSize** Controls autosizing of columns. Accepted values: true to autosize all columns, array of column names to autosize specific columns. Applied before saving a document.
* **useTempDir** If this is set, yii-phpexcel will save files to a temporary dir.
This is mainly useful when you want to generate a report and immediately send it to the user for download.
After your PHP script has finished the temporary file will be removed.

Examples
--------

Read an Excel file

```php
public function readExcelFile($filepath)
{
    $manager = Yii::app()->getComponent('yii-phpexcel');
    $excel = $manager->get($filepath);
    return $excel->read();
}
```
By default the **read** function will try to read a header row from the first row of the document. If you add false as the first parameter you'll get the raw data as it comes from PHPExcel

---

Write an Excel file

```
public function writeExcelFile()
{
    $manager = Yii::app()->getComponent('yii-phpexcel');
    //Create empty instance
    $excel = $manager->create();
    .....
    //Add a header row with a grey background
    $headerStyle = array(
        'fill' => array(
            'type' => \PHPExcel_Style_Fill::FILL_SOLID,
            'color' => array('rgb' => 'CCCCCC')
        )
    );
    $excel->addHeaderRow(
        array(
            'header1',
            'header2',
            'header3'
        ),
        $headerStyle
    );
    .....
    //Add a few rows of data to the document
    //Note that addData doesn't care about the actual keys in the data, only the order of values
    $data = array(
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
    $excel->addData($data);
    $excel->save();
}
```
This would generate something like:

| header1 | header2 | header3   |
| ------- | ------- | --------- |
| data1   | data2   | data 3    |
| 1       | Example | Something |
---

Download an Excel file
```
public function downloadExcelFile($filepath)
{
    $manager = Yii::app()->getComponent('yii-phpexcel');
    $excel = $manager->get($filepath);
    $excel->download();
}
```

Additional info
---------------

The difference between the add and set functions in the _Excel_ class is that the add functions will add to the end of the document whereas
the set functions can change content in a less linear fashion, working directly with specific rows. The add functions use the set functions themselves.

To add empty rows or columns you can simply use an empty array.
<?php

namespace laxu\yii_phpexcel\components;

/**
 * Class ExcelManager
 * @package laxu\yii_phpexcel\components
 */
class ExcelManager extends \CApplicationComponent
{
    /**
     * @var string directory where files are stored
     */
    public $filePath;

    public function init()
    {
        parent::init();
    }


    /**
     * Get an instance of Excel class
     * @param string $filename
     * @return Excel
     */
    public function get($filename)
    {
        $excel = new Excel($filename, $this->filePath);
        return $excel;
    }

    /**
     * Create a new instance of Excel class
     * @return Excel
     */
    public function create()
    {
        return new Excel(null, $this->filePath);
    }
}
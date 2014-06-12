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
     * @return \Excel
     */
    public function get($filename)
    {
        return $this->buildInstance($filename);
    }

    /**
     * Create a new instance of Excel class
     * @return \Excel
     */
    public function create()
    {
        return $this->buildInstance();
    }

    /**
     * Create Excel component
     * @param null|string $filename
     * @return \Excel
     */
    protected function buildInstance($filename = null) {
        $excel = Yii::createComponent(array(
                'class' => 'Excel',
                'filename' => $filename,
                'filePath' => $this->filePath
            ));
        return $excel;
    }
}
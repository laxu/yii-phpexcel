<?php

namespace laxu\yii_phpexcel\components;

/**
 * Class ExcelManager
 * @package laxu\yii_phpexcel\components
 */
class ExcelManager extends \CApplicationComponent
{
    /**
     * @var string directory alias where files are stored
     */
    public $savePath;

    /**
     * @var string class used for creating Excel instances
     */
    public $excelClass = '\laxu\yii_phpexcel\components\Excel';

    /**
     * @inheritdoc
     */
    public function init()
    {
        parent::init();
    }

    /**
     * Get an instance of Excel class
     * @param string $filePath
     * @throws \CException
     * @return Excel
     */
    public function get($filePath)
    {
        $excel = $this->buildInstance();
        $excel->loadFromFile($filePath);
        return $excel;
    }

    /**
     * Create a new instance of Excel class
     * @return Excel
     */
    public function create()
    {
        return $this->buildInstance();
    }

    /**
     * Create Excel component
     * @return Excel
     */
    protected function buildInstance()
    {
        $excel = \Yii::createComponent(
            array(
                'class' => $this->excelClass,
                'savePath' => $this->savePath
            )
        );
        $excel->init();
        return $excel;
    }
}
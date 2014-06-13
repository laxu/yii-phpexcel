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
        if (empty($filePath)) {
            throw new \CException(\Yii::t('excel', 'Empty filename'));
        }
        return $this->buildInstance($filePath);
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
     * @param null|string $filePath file location
     * @return Excel
     */
    protected function buildInstance($filePath = null)
    {
        $excel = \Yii::createComponent(
            array(
                'class' => '\laxu\yii_phpexcel\components\Excel',
                'filePath' => $filePath,
                'savePath' => $this->savePath
            )
        );
        $excel->init();
        return $excel;
    }
}
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
    public $filePath;

    public function init()
    {
        parent::init();
    }

    /**
     * Get an instance of Excel class
     * @param string $filename
     * @param string $filePath
     * @throws \CException
     * @return \Excel
     */
    public function get($filename, $filePath = null)
    {
        if (empty($filename)) {
            throw new \CException('Empty filename');
        }
        return $this->buildInstance($filename, $filePath);
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
     * @param null|string $filePath filepath alias
     * @return \Excel
     */
    protected function buildInstance($filename = null, $filePath = null)
    {
        if($filePath === null) {
            $filePath = $this->filePath;
        }
        $excel = \Yii::createComponent(
            array(
                'class' => '\laxu\yii_phpexcel\components\Excel',
                'filename' => $filename,
                'filePath' => $filePath
            )
        );

        $excel->init();
        return $excel;
    }
}
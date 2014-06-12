<?php


namespace nordsoftware\yii_phpexcel\components;

/**
 * Class Excel
 * @package nordsoftware\yii_phpexcel\components
 * @var $objPHPExcel PHPExcel;
 */
class Excel {
    /**
     * @var string PHPExcel location
     */
    public $libPath = 'vendor.yii-phpexcel.lib.phpexcel';

    /**
     * @var PHPExcel instance
     */
    public $phpExcel;

    function __construct($file = null)
    {
        Yii::import($this->libPath);
        $this->setInstance($file);
    }


    /**
     * Read an Excel file
     * @param string $file filename
     * @return array
     */
    public function read($file)
    {
        return $this->phpExcel->getActiveSheet()->toArray(null,true,true,true);
    }

    /**
     * Write to an Excel file
     * @param $file
     * @param $data
     * @throws CException
     * @return bool
     */
    public function write($file, $data)
    {
        return $this->phpExcel->write($data);
    }

    /**
     * Set PHPExcel instance
     * @param null|string|PHPExcel $file
     * @throws CException
     */
    public function setInstance($file = null){
        if($file === null) {
            //Create new workbook
            $this->phpExcel = new PHPExcel();
        }
        elseif(is_string($file)) {
            //Load an existing one
            $this->phpExcel = PHPExcel_IOFactory::load($file);
        }
        elseif($file instanceof PHPExcel) {
            //Use another PHPExcel instance passed to this one
            $this->phpExcel = $file;
        }
        else {
            throw new CException('$file should be null, filename or PHPExcel instance');
        }
    }

    /**
     * Get current PHPExcel instance
     * @return PHPExcel
     */
    public function getInstance() {
        return $this->phpExcel;
    }
} 
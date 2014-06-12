<?php


namespace laxu\yii_phpexcel\components;

/**
 * Class Excel
 * @package nordsoftware\yii_phpexcel\components
 * @var $objPHPExcel PHPExcel;
 */
class Excel
{
    /**
     * @var string PHPExcel location
     */
    public $libPath = 'vendor.phpexcel';

    /**
     * @var string directory where files are stored
     */
    public $filePath;

    /**
     * @var string filename
     */
    public $filename;

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
        return $this->phpExcel->getActiveSheet()->toArray(null, true, true, true);
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
        $this->phpExcel->write($data);
        $this->phpExcel->save($this->getFilePath());
    }

    /**
     * Set PHPExcel instance
     * @param null|string|PHPExcel $filename
     * @throws CException
     */
    public function setInstance($filename = null)
    {
        if ($filename === null) {
            //Create new workbook
            $this->phpExcel = new PHPExcel();
        } elseif (is_string($filename)) {
            //Load an existing one
            $this->filename = $filename;
            $this->phpExcel = PHPExcel_IOFactory::load($this->getFilePath());
        } elseif ($filename instanceof PHPExcel) {
            //Use another PHPExcel instance passed to this one
            $this->phpExcel = $filename;
        } else {
            throw new CException('Filename should be null, filename or PHPExcel instance');
        }
    }

    /**
     * Get current PHPExcel instance
     * @return PHPExcel
     */
    public function getInstance()
    {
        return $this->phpExcel;
    }

    public function getFilePath()
    {
        return $this->filePath . $this->filename;
    }
} 
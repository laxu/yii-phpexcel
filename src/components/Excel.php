<?php


namespace laxu\yii_phpexcel\components;

/**
 * Class Excel
 * @package laxu\yii_phpexcel\components
 * @var $objPHPExcel \PHPExcel;
 */
class Excel extends CComponent
{
    /**
     * @var string directory where files are stored
     */
    public $filePath;

    /**
     * @var string filename
     */
    public $filename;

    /**
     * @var \PHPExcel instance
     */
    public $phpExcel;

    /**
     * @var bool Whether file has been saved
     */
    protected $stored = false;

    /**
     * @var string|array Characters for columns
     */
    protected $columnCharSet = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ';

    public function init()
    {
        parent::init();
        $this->columnCharSet = str_split($this->columnCharSet);
        if($this->filename === null) {
            $this->createNewInstance();
        }
        else {
            $this->loadFromFile();
        }
    }

    /**
     * Get raw data as it comes from PHPExcel
     * @param mixed $nullValue Value returned in the array entry if a cell doesn't exist
     * @param boolean $calculateFormulas Should formulas be calculated?
     * @param boolean $formatData Should formatting be applied to cell values?
     * @param boolean $returnCellRef False - Return a simple array of rows and columns indexed by number counting from zero
     *                               True - Return rows and columns indexed by their actual row and column IDs
     * @return array
     */
    public function readRaw($nullValue = null, $calculateFormulas = true, $formatData = true, $returnCellRef = false)
    {
        return $this->phpExcel->getActiveSheet()->toArray($nullValue, $calculateFormulas, $formatData, $returnCellRef);
    }

    /**
     * Read an Excel file
     * @return array
     */
    public function read()
    {
        $rawData = $this->readRaw();
        //Assume first row contains headers
        $headers = $rawData[0];
        unset($rawData[0]);
        $outputData = array();
        foreach ($rawData as &$row) {
            //Loop thru rows and create associative arrays based on headers
            $rowData = array();
            foreach ($row as $idx => $value) {
                if (isset($headers[$idx])) {
                    if (is_string($value)) {
                        $value = utf8_decode($value);
                    }
                    $rowData[$headers[$idx]] = $value;
                }
            }
            $outputData[] = $rowData;
        }
        return $outputData;
    }

    /**
     * Set data
     * @param array $dataSet Data as an array of arrays, first element should be header row
     * @return bool
     */
    public function setData($dataSet)
    {
        $workSheet = $this->phpExcel->getActiveSheet();
        foreach ($dataSet as $rowIdx => $rowData) {
            $rowData = array_values($rowData);
            foreach ($rowData as $cellIdx => $cellData) {
                $cellKey = $this->getCellKey($cellIdx, $rowIdx);
                $workSheet->setCellValue($cellKey, $cellData);
            }
        }
    }

    /**
     * Save data to an Excel file
     */
    public function save()
    {
        $objWriter = \PHPExcel_IOFactory::createWriter($this->phpExcel, "Excel2007");
        $objWriter->save($this->resolveFilePath());
        $this->stored = true;
    }

    /**
     * Download the Excel file
     * @throws \CHttpException
     */
    public function download()
    {
        $filePath = $this->resolveFilePath();

        if (file_exists($filePath)) {
            \Yii::app()->getRequest()->sendFile($this->filename, file_get_contents($filePath));
        } else {
            throw new \CHttpException(404, 'File not found');
        }
    }

    /**
     * Create a new PHPExcel instance
     */
    public function createNewInstance() {
        $this->phpExcel = new \PHPExcel();
        $this->filename = $this->generateFilename();
    }

    /**
     * Create PHPExcel instance by loading the file
     * @throws \CException
     */
    public function loadFromFile()
    {
        if(!is_string($this->filename)) {
            throw new \CException('Filename should be a string containing a filename');
        }

        $filePath = $this->resolveFilePath();
        if(!file_exists($filePath)) {
            throw new \CException('File not found');
        }

        $this->phpExcel = \PHPExcel_IOFactory::load($filePath);
    }

    /**
     * Get current PHPExcel instance
     * @return \PHPExcel
     */
    public function getInstance()
    {
        return $this->phpExcel;
    }

    /**
     * Resolve full path for file
     * @return string
     */
    public function resolveFilePath()
    {
        return \Yii::getPathOfAlias($this->filePath) . "/" . $this->filename;
    }

    /**
     * Get Excel column name
     * @param int $cellIdx
     * @param int $rowIdx
     * @return string
     */
    protected function getCellKey($cellIdx, $rowIdx)
    {
        return $this->columnCharSet[$cellIdx] . ($rowIdx + 1);
    }

    /**
     * Generate filename
     * @param string $extension file extension
     * @return string
     */
    protected function generateFilename($extension = 'xlsx')
    {
        $filename = uniqid() . '.' . $extension;
        while (file_exists($this->filePath . $filename)) {
            $filename = uniqid() . '.' . $extension;
        }
        return $filename;
    }
}

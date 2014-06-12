<?php


namespace laxu\yii_phpexcel\components;

/**
 * Class Excel
 * @package laxu\yii_phpexcel\components
 * @var $objPHPExcel \PHPExcel;
 */
class Excel
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

    function __construct($file = null, $filePath)
    {
        $this->columnCharSet = str_split($this->columnCharSet);
        $this->filePath = $filePath;
        $this->setInstance($file);
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
     * Set PHPExcel instance
     * @param null|string $filename
     * @throws \CException
     */
    public function setInstance($filename = null)
    {
        if ($filename === null) {
            //Create new workbook
            $this->phpExcel = new \PHPExcel();
            $this->filename = $this->generateFilename();
        } elseif (is_string($filename)) {
            //Load an existing one
            $this->filename = $filename;
            $this->phpExcel = \PHPExcel_IOFactory::load($this->resolveFilePath());
            $this->stored = true;
        } else {
            throw new \CException('Filename should be null or filename');
        }
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

<?php


namespace laxu\yii_phpexcel\components;

/**
 * Class Excel
 * @package laxu\yii_phpexcel\components
 * @var $objPHPExcel \PHPExcel;
 */
class Excel extends \CComponent
{
    /**
     * @var string alias of directory where files are stored
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
     * @var int Currently active worksheet
     */
    public $activeSheet = 0;

    /**
     * @var bool Whether file has been saved
     */
    protected $stored = false;

    /**
     * @var array Document settings
     */
    public $settings = array();

    /**
     * @var string|array Characters for columns
     */
    protected $columnCharSet = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ';

    /**
     * Initialize component
     */
    public function init()
    {
        //Turn columnCharSet into an array
        $this->columnCharSet = str_split($this->columnCharSet);
        if ($this->filename === null) {
            //No filename found, create an empty instance
            $this->createNewInstance();
        } else {
            //Filename, try to load from file
            $this->loadFromFile();
        }

        //Set active worksheet
        // $this->setWorksheet($this->activeSheet);
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
        return $this->getWorksheet()->toArray($nullValue, $calculateFormulas, $formatData, $returnCellRef);
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
     * Set document properties
     * @param array $data as key => value pairs, see below for allowed keys
     */
    public function setDocumentProperties($data)
    {
        $properties = $this->phpExcel->getProperties();
        foreach ($data as $key => $value) {
            switch ($key) {
                case 'creator':
                    $properties->setCreator($value);
                    break;
                case 'lastModifiedBy':
                    $properties->setLastModifiedBy($value);
                    break;
                case 'title':
                    $properties->setTitle($value);
                    break;
                case 'subject':
                    $properties->setSubject($value);
                    break;
                case 'description':
                    $properties->setDescription($value);
                    break;
                case 'company':
                    $properties->setCompany($value);
                    break;
                case 'category':
                    $properties->setCategory($value);
                    break;
                case 'manager':
                    $properties->setManager($value);
                    break;
                case 'keywords':
                    $properties->setKeywords($value);
                    break;
                default:
                    $properties->setCustomProperty($key, $value);
            }
        }
    }

    /**
     * Set data
     * @param array $dataSet Data as an array of arrays, first element should be header row
     * @return bool
     */
    public function setData($dataSet)
    {
        $workSheet = $this->getWorksheet();
        foreach ($dataSet as $rowIdx => $rowData) {
            $rowData = array_values($rowData);
            foreach ($rowData as $cellIdx => $cellData) {
                $cellKey = $this->getCellKey($cellIdx, $rowIdx);
                $workSheet->setCellValue($cellKey, $cellData);
            }
        }
    }

    /**
     * Set settings
     * @param array $data
     * @throws \CException
     */
    public function setSettings($data)
    {
        if (!is_array($data)) {
            throw new \CException('$data is not an array');
        }
        $this->settings = array_merge($this->settings, $data);
        $this->applySettings();
    }

    /**
     * Get settings
     * @param null|string $setting Get a specific setting, null returns all settings
     * @return array|null
     */
    public function getSettings($setting = null)
    {
        if ($setting === null) {
            return $this->settings;
        }
        return array_key_exists($setting, $this->settings) ? $this->settings[$setting] : null;
    }

    /**
     * Apply settings to worksheet
     */
    public function applySettings()
    {
        $workSheet = $this->getWorksheet();
        foreach ($this->settings as $key => $value) {
            switch ($key) {
                case 'autoSize':
                    if (!$value) {
                        continue;
                    }
                    if (is_array($key)) {
                        //Set specific columns to size automatically
                        foreach ($value as $column) {
                            $workSheet->getColumnDimension($column)->setAutoSize(true);
                        }
                    } else {
                        //Set all columns to size automatically
                        $columns = $workSheet->getColumnDimensions();
                        foreach ($columns as $column) {
                            $column->setAutoSize(true);
                        }
                    }
                    break;
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
    public function createNewInstance()
    {
        $this->phpExcel = new \PHPExcel();
        $this->filename = $this->generateFilename();
    }

    /**
     * Create PHPExcel instance by loading the file
     * @throws \CException
     */
    public function loadFromFile()
    {
        if (!is_string($this->filename)) {
            throw new \CException('Filename should be a string containing a filename');
        }

        $filePath = $this->resolveFilePath();
        if (!file_exists($filePath)) {
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
     * Set active worksheet
     * @param $idx
     */
    public function setWorksheet($idx)
    {
        $this->phpExcel->setActiveSheetIndex($idx);
        $this->activeSheet = $idx;
    }

    /**
     * Get active worksheet
     * @return \PHPExcel_Worksheet
     */
    public function getWorksheet()
    {
        return $this->phpExcel->getActiveSheet();
    }

    /**
     * Create a new worksheet and set is active
     * @param int $idx Index where the new worksheet should go
     */
    public function createWorksheet($idx = null)
    {
        $this->phpExcel->createSheet($idx);
        $this->setWorksheet($idx);
    }

    /**
     * Set title of active worksheet
     * @param string $title
     */
    public function setWorksheetTitle($title)
    {
        $this->getWorksheet()->setTitle($title);
    }

    /**
     * Resolve full path for file
     * @return string
     */
    public function resolveFilePath()
    {
        return \Yii::getPathOfAlias($this->filePath) . '/' . $this->filename;
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
        $filePath = \Yii::getPathOfAlias($this->filePath);
        $filename = uniqid() . '.' . $extension;
        while (file_exists($filePath . '/' . $filename)) {
            $filename = uniqid() . '.' . $extension;
        }
        return $filename;
    }
}

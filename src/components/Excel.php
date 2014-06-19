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
     * @var string directory where the file tied to this component is located
     */
    public $filePath;

    /**
     * @var string filename of the file tied to this component
     */
    public $filename;

    /**
     * @var string Where new files are to be saved, can be different from the one in filepath
     */
    public $savePath;

    /**
     * @var \PHPExcel instance
     */
    public $phpExcel;

    /**
     * @var int Currently active worksheet index
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
    protected $columnCharSet;

    /**
     * @var array Scenarios where to apply different settings
     */
    protected $settingsScenarios = array(
        'init' => array('useTempDir'),
        'beforeSave' => array('autoSize')
    );

    /**
     * @var \PHPExcel_Worksheet Reference to current worksheet
     */
    private $_currentWorkSheet;

    /**
     * @var int Last row in document
     */
    private $_currentRow = 0;

    /**
     * @var array List of indices for rows that are headers
     */
    private $_headerRows = array();

    /**
     * Initialize component
     */
    public function init()
    {
        $this->columnCharSet = range('A', 'Z');
        $this->createNewInstance();

        $this->filePath = $this->savePath;
        $this->generateFilename();

        //Set active worksheet
        $this->setWorksheet($this->activeSheet);
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
        $workSheet = $this->getWorksheet();
        return $workSheet->toArray($nullValue, $calculateFormulas, $formatData, $returnCellRef);
    }

    /**
     * Read an Excel file
     * @param bool $useHeadersAsKeys Use headers as keys
     * @return array
     */
    public function read($useHeadersAsKeys = true)
    {
        $rawData = $this->readRaw();
        $rawData = $this->removeEmptyRows($rawData);
        if (!$useHeadersAsKeys) {
            return $rawData;
        }
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
     * Remove empty rows
     * @param $data
     * @return array
     */
    public function removeEmptyRows($data)
    {
        foreach($data as $idx => $row) {
            if(!array_filter($row)) {
                unset($data[$idx]);
            }
        }
        return $data;
    }

    /**
     * Set document properties
     * @param array $data as key => value pairs, see below for allowed keys
     */
    public function setDocumentProperties($data)
    {
        $properties = $this->phpExcel->getProperties();
        foreach ($data as $key => $value) {
            $method = 'set' . ucfirst($key);
            if (method_exists($properties, $method)) {
                $properties->$method($value);
            } else {
                $properties->setCustomProperty($key, $value);
            }
        }
    }

    /**
     * Set cell style
     * @param array|string $cells e.g. array('A1','A2','B2') or string 'A1' or 'A1:G5'
     * @param array $style See PHPExcel documentation
     */
    public function setStyle($cells, $style)
    {
        $workSheet = $this->getWorksheet();
        $workSheet->getStyle($cells)->applyFromArray($style);
    }

    /**
     * Set style of row
     * @param int $rowIdx zero based row index
     * @param array $style See PHPExcel documentation
     */
    public function setRowStyle($rowIdx, $style)
    {
        $workSheet = $this->getWorksheet();
        $cells = $this->getCellRange(
            0,
            \PHPExcel_Cell::columnIndexFromString($workSheet->getHighestColumn($rowIdx + 1)) - 1,
            $rowIdx,
            $rowIdx
        );
        $this->setStyle($cells, $style);
    }

    /**
     * Add data to end of document
     * @param array $dataSet Data as an array of arrays
     * @return bool
     */
    public function addData($dataSet)
    {
        $data = array();
        foreach ($dataSet as &$rowData) {
            $data[$this->_currentRow] = $rowData;
            $this->_currentRow++;
        }

        $this->setData($data);
    }

    /**
     * Add a header row
     * @param array $data
     * @param null|array $style See PHPExcel documentation
     */
    public function addHeaderRow($data, $style = null)
    {
        $this->setHeaderRow($this->_currentRow, $data, $style);
        $this->_currentRow++;
    }

    /**
     * Set data from a dataset containing data in "row idx => data" format
     * @param $dataSet
     */
    public function setData($dataSet)
    {
        foreach ($dataSet as $rowIdx => $rowData) {
            $this->setRowContent($rowIdx, $rowData);
        }
    }

    /**
     * Set a row as a header
     * @param int $rowIdx Zero-based row index where to add header
     * @param array $data Cell data for row
     * @param null|array $style Style the header row. See PHPExcel documentation
     */
    public function setHeaderRow($rowIdx, $data, $style = null)
    {
        $this->setRowContent($rowIdx, $data);

        if ($style) {
            $this->setRowStyle($rowIdx, $style);
        }

        $this->_headerRows[] = $rowIdx;
    }

    /**
     * Set content for a specific row
     * @param int $rowIdx The row where you want to set data
     * @param array $data
     */
    public function setRowContent($rowIdx, $data)
    {
        $workSheet = $this->getWorksheet();
        $cellKey = $this->getCellKey(0, $rowIdx);
        $workSheet->fromArray(array_values($data), null, $cellKey);
    }

    /**
     * Set settings
     * @param array $newSettings
     * @throws \CException
     */
    public function setSettings($newSettings)
    {
        if (!is_array($newSettings)) {
            throw new \CException('$data is not an array');
        }
        $this->settings = array_merge($this->settings, $newSettings);
        $this->applySettings('init');
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
     * Get the scenario where to apply settings
     * @param string $key Scenario key
     * @return array
     * @throws \CException
     */
    protected function getSettingsScenario($key)
    {
        if (!isset($this->settingsScenarios[$key])) {
            throw new \CException(t('yii-phpexcel', 'Settings scenario not found'));
        }

        return $this->settingsScenarios[$key];
    }

    /**
     * Apply settings to worksheet
     * @param string $scenarioKey Scenario where to apply settings
     */
    public function applySettings($scenarioKey)
    {
        $scenario = $this->getSettingsScenario($scenarioKey);
        $workSheet = $this->getWorksheet();
        foreach ($this->settings as $key => $value) {
            if (!in_array($key, $scenario)) {
                continue;
            }
            switch ($key) {
                case 'autoSize':
                    if (!$value) {
                        continue;
                    }
                    if (is_array($value)) {
                        //Set specific columns to size automatically
                        foreach ($value as $column) {
                            $workSheet->getColumnDimension($column)->setAutoSize(true);
                        }
                    } else {
                        //Set all columns to size automatically
                        $max = $workSheet->getHighestColumn();
                        foreach (range('A', $max) as $column) {
                            $workSheet->getColumnDimension($column)->setAutoSize(true);
                        }
                    }
                    break;
                case 'useTempDir':
                    //Use PHP temp dir to save files
                    $path = sys_get_temp_dir();
                    if (!preg_match('/\/$/', $path)) {
                        $path .= '/';
                    }
                    $this->filePath = $path;
                    break;
            }
        }
    }

    /**
     * Save data to an Excel file
     */
    public function save()
    {
        $this->applySettings('beforeSave');
        $objWriter = \PHPExcel_IOFactory::createWriter($this->phpExcel, "Excel2007");
        $objWriter->save($this->getFullPath());
        $this->stored = true;
    }

    /**
     * Download the Excel file
     * @throws \CHttpException
     */
    public function download()
    {
        $filePath = $this->getFullPath();

        if (file_exists($filePath)) {
            \Yii::app()->getRequest()->sendFile($this->filename, file_get_contents($filePath));
        } else {
            throw new \CHttpException(404, \Yii::t('yii-phpexcel', 'File not found'));
        }
    }

    /**
     * Create a new PHPExcel instance
     */
    public function createNewInstance()
    {
        $this->phpExcel = new \PHPExcel();
    }

    /**
     * Create PHPExcel instance by loading the file
     * @param string $filePath Full path to file
     * @throws \CException
     */
    public function loadFromFile($filePath)
    {
        $this->setFullPath($filePath);

        if (!file_exists($filePath)) {
            throw new \CException('File not found');
        }

        $this->phpExcel = \PHPExcel_IOFactory::load($this->getFullPath());
        $this->setWorksheet($this->activeSheet);
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
        $this->_currentWorkSheet = $this->phpExcel->getActiveSheet();
        $this->activeSheet = $idx;
    }

    /**
     * Get active worksheet
     * @return \PHPExcel_Worksheet
     */
    public function getWorksheet()
    {
        if ($this->_currentWorkSheet === null) {
            $this->_currentWorkSheet = $this->phpExcel->getActiveSheet();
        }
        return $this->_currentWorkSheet;
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
     * Set filename and path from a full path
     * @param string $filePath full path including filename
     * @throws \CException
     */
    public function setFullPath($filePath)
    {
        $data = pathinfo($filePath);
        $filename = $data['basename'];
        $path = $data['dirname'];
        if (empty($filename) || empty($path)) {
            throw new \CException(t('yii-phpexcel', 'Could not find filename or path'));
        }

        $this->filePath = $path;
        $this->filename = $filename;
    }

    /**
     * Get full path for filename tied to this instance
     * @return string
     * @throws \CException
     */
    public function getFullPath()
    {
        if (empty($this->filename)) {
            throw new \CException(t('yii-phpexcel', 'Filename is undefined'));
        }

        return $this->filePath . '/' . $this->filename;
    }

    /**
     * Generate filename
     * @param string $extension file extension
     * @throws \CException
     */
    protected function generateFilename($extension = 'xlsx')
    {
        if (empty($this->filePath)) {
            throw new \CException(\Yii::t('yii-phpexcel', 'filePath is undefined'));
        }
        $filename = uniqid() . '.' . $extension;
        while (file_exists($this->filePath . '/' . $filename)) {
            $filename = uniqid() . '.' . $extension;
        }
        $this->filename = $filename;
    }

    /**
     * Get cell range using numbers
     * @param int $start
     * @param int $end
     * @param int $startRow
     * @param int $endRow
     * @return string
     */
    public function getCellRange($start, $end, $startRow = 0, $endRow = 0)
    {
        $startChar = $this->columnCharSet[$start];
        $endChar = $this->columnCharSet[$end];

        return $startChar . ($startRow + 1) . ':' . $endChar . ($endRow + 1);
    }

}

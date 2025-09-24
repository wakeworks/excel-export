<?php

/**
 * ExcelDataFormatter provides a DataFormatter allowing an {@link SS_link} of
 * {@link DataObjectInterface} to be exported to be to Excel 2007 Spreadsheet
 * (XLSX).
 *
 * This class can be extended to export to other format supported by
 * {@link https://github.com/PHPOffice/PHPExcel PHPExcel}.
 *
 * @author Firebrand <hello@firebrand.nz>
 * @license MIT
 * @package silverstripe-excel-export
 */

namespace ExcelExport;

use SilverStripe\Model\List\ArrayList;
use SilverStripe\View\TemplateEngine;

use SilverStripe\Model\List\SS_List;
use Override;
use PhpOffice\PhpSpreadsheet\Cell\Coordinate;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;
use SilverStripe\Control\Controller;
use SilverStripe\ORM\DataObjectInterface;
use SilverStripe\Security\Security;
use SilverStripe\Core\Config\Config;

class ExcelDataFormatter extends DataFormatter
{


    private static $api_base = "api/v1/";

    /**
     * Determined what we will use as headers for the spread sheet.
     * @var bool
     */
    protected $useLabelsAsHeaders = null;

    /**
     * @inheritdoc
     */
    public function supportedExtensions()
    {
        return [
            'xlsx',
        ];
    }

    /**
     * @inheritdoc
     */
    public function supportedMimeTypes()
    {
        return [
            'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        ];
    }

    /**
     * @inheritdoc
     */
    public function convertDataObject(DataObjectInterface $do)
    {
        return $this->convertDataObjectSet(ArrayList::create([$do]));
    }

    /**
     * @inheritdoc
     */
    public function convertDataObjectSet(SS_List $set)
    {
        $this->setHeader();

        $excel = $this->getPhpExcelObject($set);

        $fileData = $this->getFileData($excel, 'Excel2007');

        return $fileData;
    }

    /**
     * Set the HTTP Content Type header to the appropriate Mime Type.
     */
    protected function setHeader()
    {
        Controller::curr()->getResponse()
            ->addHeader("Content-Type", $this->supportedMimeTypes()[0]);
    }

    /**
     * @inheritdoc
     */
    #[Override]
    protected function getFieldsForObj($obj)
    {
        $dbFields = [];

        // if custom fields are specified, only select these
        if(is_array($this->customFields)) {
            foreach($this->customFields as $fieldName) {
                // @todo Possible security risk by making methods accessible - implement field-level security
                if($obj->hasField($fieldName) || $obj->hasMethod('get' . $fieldName)) {
                    $dbFields[$fieldName] = $fieldName;
                }
            }
        } elseif ($obj->hasMethod('getExcelExportFields')) {
            $dbFields = $obj->getExcelExportFields();
        } else {
            // by default, all database fields are selected
            $dbFields = $obj->inheritedDatabaseFields();
        }

        if(is_array($this->customAddFields)) {
            foreach($this->customAddFields as $fieldName) {
                // @todo Possible security risk by making methods accessible - implement field-level security
                if($obj->hasField($fieldName) || $obj->hasMethod('get' . $fieldName)) {
                    $dbFields[$fieldName] = $fieldName;
                }
            }
        }

        // Make sure our ID field is the first one.
        $dbFields = ['ID' => 'Int'] + $dbFields;

        if(is_array($this->removeFields)) {
            $dbFields = array_diff_key($dbFields, array_combine($this->removeFields,$this->removeFields));
        }

        return $dbFields;
    }

    /**
     * Generate a {@link Spreadsheet} for the provided DataObject List
     * @param SS_List $set List of DataObjects
     * @return Spreadsheet
     */
    public function getPhpExcelObject(SS_List $set)
    {
        // Get the first object. We'll need it to know what type of objects we
        // are dealing with
        $first = $set->first();

        // Get the Excel object
        $excel = $this->setupExcel($first);
        $sheet = $excel->setActiveSheetIndex(0);

        // Make sure we have at lease on item. If we don't, we'll be returning
        // an empty spreadsheet.
        if ($first) {
            // Set up the header row
            $fields = $this->getFieldsForObj($first);
            $this->headerRow($sheet, $fields, $first);

            // Add a new row for each DataObject
            foreach ($set as $item) {
                $this->addRow($sheet, $item, $fields);
            }

            // Freezing the first column and the header row
            $sheet->freezePane("B2");

            // Auto sizing all the columns
            $col = count($fields);
            for ($i = 1; $i <= $col; $i++) {
                $sheet
                    ->getColumnDimension(
                        Coordinate::stringFromColumnIndex($i)
                    )
                    ->setAutoSize(true);
            }

        }

        return $excel;
    }

    /**
     * Initialize a new {@link PHPExcel} object based on the provided
     * {@link DataObjectInterface} interface.
     * @param  DataObjectInterface $do
     * @return Spreadsheet
     */
    protected function setupExcel(DataObjectInterface $do)
    {
        // Try to get the current user
        $member = Security::getCurrentUser();
        $creator = $member ? $member->getName() : '';

        // Get information about the current Model Class
        $singular = $do ? $do->i18n_singular_name() : '';
        $plural = $do ? $do->i18n_plural_name() : '';

        // Create the Spread sheet
        $excel = new Spreadsheet();

        $excel->getProperties()
            ->setCreator($creator)
            ->setTitle(_t(
                'firebrandhq.EXCELEXPORT',
                '{singular} export',
                'Title for the spread sheet export',
                ['singular' => $singular]
            ))
            ->setDescription(_t(
                'firebrandhq.EXCELEXPORT',
                'List of {plural} exported out of a SilverStripe website',
                'Description for the spread sheet export',
                ['plural' => $plural]
            ));

        // Give a name to the sheet
        if ($plural) {
            $excel->getActiveSheet()->setTitle($plural);
        }

        return $excel;
    }

    /**
     * Add an header row to a {@link PHPExcel_Worksheet}.
     * @param  Worksheet $sheet
     * @param  array              $fields List of fields
     * @param  DataObjectInterface  $do
     * @return Worksheet
     */
    protected function headerRow(Worksheet &$sheet, array $fields, DataObjectInterface $do)
    {
        // Counter
        $row = 1;
        $col = 1;

        $useLabelsAsHeaders = $this->getUseLabelsAsHeaders();

        // Add each field to the first row
        foreach (array_keys($fields) as $field) {
            $header = $useLabelsAsHeaders ? $do->fieldLabel($field) : $field;
            $sheet->setCellValue([$col, $row], $header);
            $col++;
        }

        // Get the last column
        $col--;
        $endcol = Coordinate::stringFromColumnIndex($col);

        // Set Autofilters and Header row style
        $sheet->setAutoFilter(sprintf('A1:%s1', $endcol));
        $sheet->getStyle(sprintf('A1:%s1', $endcol))->getFont()->setBold(true);


        return $sheet;
    }

    /**
     * Add a new row to a {@link PHPExcel_Worksheet} based of a
     * {@link DataObjectInterface}
     * @param Worksheet  $sheet
     * @param DataObjectInterface $item
     * @param array               $fields List of fields to include
     * @return Worksheet
     */
    protected function addRow(
        Worksheet &$sheet,
        DataObjectInterface $item,
        array $fields
    ) {
        $row = $sheet->getHighestRow() + 1;
        $col = 1;

        foreach (array_keys($fields) as $field) {
            if ($item->hasField($field) || $item->hasMethod('get' . $field)) {
                $value = $item->$field;
            } else {
                $viewer = singleton(TemplateEngine::class)->renderString('$' . $field . '.RAW');
                $value = $item->renderWith($viewer, true);
            }

            $sheet->setCellValue([$col, $row], $value);
            $col++;
        }

        return $sheet;
    }

    /**
     * Generate a string representation of an {@link PHPExcel} spread sheet
     * suitable for output to the browser.
     * @param  Spreadsheet $excel
     * @param  string   $format Format to use when outputting the spreadsheet.
     * Must be compatible with the format expected by
     * {@link PHPExcel_IOFactory::createWriter}.
     * @return string
     */
    protected function getFileData(Spreadsheet $excel, $format)
    {
        $writer = IOFactory::createWriter($excel, $format);
        ob_start();
        $writer->save('php://output');
        $fileData = ob_get_clean();

        return $fileData;
    }

    /**
     * Accessor for UseLabelsAsHeaders. If this is `true`, the data formatter will call {@link DataObject::fieldLabel()} to pick the header strings. If it's set to false, it will use the raw field name.
     *
     * You can define this for a specific ExcelDataFormatter instance with `setUseLabelsAsHeaders`. You can set the default for all ExcelDataFormatter instance in your YML config file:
     *
     * ```
     * ExcelDataFormatter:
     *   UseLabelsAsHeaders: true
     * ```
     *
     * Otherwise, the data formatter will default to false.
     *
     * @return bool
     */
    public function getUseLabelsAsHeaders()
    {
        if ($this->useLabelsAsHeaders !== null) {
            return $this->useLabelsAsHeaders;
        }

        $useLabelsAsHeaders = Config::inst()->get(self::class, 'UseLabelsAsHeaders');
        if ($useLabelsAsHeaders !== null) {
            return $useLabelsAsHeaders;
        }

        return false;
    }

    /**
     * Setter for UseLabelsAsHeaders. If this is `true`, the data formatter will call {@link DataObject::fieldLabel()} to pick the header strings. If it's set to false, it will use the raw field name.
     *
     * If `$value` is `null`, the data formatter will fall back on whatevr the default is.
     * @param bool $value
     * @return ExcelDataFormatter
     */
    public function setUseLabelsAsHeaders($value)
    {
        $this->useLabelsAsHeaders = $value === null ? null : (bool)$value;

        return $this;
    }
}

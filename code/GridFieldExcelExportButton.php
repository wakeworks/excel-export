<?php

/**
 * Enhanced GridField export button that allows the list to be exported to:
 *  * Excel 2007,
 *  * Excel 5,
 *  * CSV
 *
 * The button appears has a Split button exposing the 3 possible export format.
 *
 * @author Firebrand <hello@firebrand.nz>
 * @license MIT
 * @package silverstripe-excel-export
 */

namespace ExcelExport;


use SilverStripe\Model\List\ArrayList;
use SilverStripe\Model\List\SS_List;
use SilverStripe\Control\Controller;
use SilverStripe\Control\HTTPRequest;
use SilverStripe\Forms\GridField\GridField;
use SilverStripe\Forms\GridField\GridField_ActionProvider;
use SilverStripe\Forms\GridField\GridField_FormAction;
use SilverStripe\Forms\GridField\GridField_HTMLProvider;
use SilverStripe\Forms\GridField\GridField_URLHandler;
use SilverStripe\Forms\GridField\GridFieldFilterHeader;
use SilverStripe\Forms\GridField\GridFieldSortableHeader;

class GridFieldExcelExportButton implements
    GridField_HTMLProvider,
    GridField_ActionProvider,
    GridField_URLHandler
{

    /**
     * Whatever to override the default $useFieldLabelsAsHeaders value for the DataFormatter.
     * @var bool
     */
    protected $useLabelsAsHeaders = null;

    /**
     * Instanciate GridFieldExcelExportButton.
     * @param string $targetFragment
     */
    public function __construct(protected $targetFragment = "before")
    {
    }

    /**
     * @inheritdoc
     *
     * Create the split button with all the export options.
     *
     * @param  GridField $gridField
     * @return array
     */
    public function getHTMLFragments($gridField)
    {
        // Set up the split button
        $splitButton = SplitButton::create('Export', 'Export');
        $splitButton->setAttribute('data-icon', 'download-csv');

        // XLSX option
        $button = GridField_FormAction::create($gridField, 'xlsxexport', _t('firebrandhq.EXCELEXPORT', 'Export to Excel (XLSX)'), 'xlsxexport', null);
        $button->addExtraClass('no-ajax');

        $splitButton->push($button);

        // XLS option
        $button = GridField_FormAction::create($gridField, 'xlsexport', _t('firebrandhq.EXCELEXPORT', 'Export to Excel (XLS)'), 'xlsexport', null);
        $button->addExtraClass('no-ajax');

        $splitButton->push($button);

        // CSV option
        $button = GridField_FormAction::create($gridField, 'csvexport', _t('firebrandhq.EXCELEXPORT', 'Export to CSV'), 'csvexport', null);
        $button->addExtraClass('no-ajax');

        $splitButton->push($button);

        // Return the fragment
        return [
            $this->targetFragment =>
                $splitButton->Field()
        ];
    }

    /**
     * @inheritdoc
     */
    public function getActions($gridField)
    {
        return ['xlsxexport', 'xlsexport', 'csvexport'];
    }

    /**
     * @inheritdoc
     */
    public function handleAction(
        GridField $gridField,
        $actionName,
        $arguments,
        $data
    ) {
        if ($actionName == 'xlsxexport') {
            return $this->handleXlsx($gridField);
        }

        if ($actionName == 'xlsexport') {
            return $this->handleXls($gridField);
        }

        if ($actionName == 'csvexport') {
            return $this->handleCsv($gridField);
        }

        return null;
    }

    /**
     * @inheritdoc
     */
    public function getURLHandlers($gridField)
    {
        return [
            'xlsxexport' => 'handleXlsx',
            'xlsexport' => 'handleXls',
            'csvexport' => 'handleCsv',
        ];
    }

    /**
     * Action to export the GridField list to an Excel 2007 file.
     * @param  GridField $gridField
     * @param  HTTPRequest    $request
     * @return string
     */
    public function handleXlsx(GridField $gridField, $request = null)
    {
        return $this->genericHandle('ExcelDataFormatter', 'xlsx', $gridField, $request);
    }

    /**
     * Action to export the GridField list to an Excel 5 file.
     * @param  GridField $gridField
     * @param  HTTPRequest    $request
     * @return string
     */
    public function handleXls(GridField $gridField, $request = null)
    {
        return $this->genericHandle('OldExcelDataFormatter', 'xls', $gridField, $request);
    }

    /**
     * Action to export the GridField list to an CSV file.
     * @param  GridField $gridField
     * @param  HTTPRequest    $request
     * @return string
     */
    public function handleCsv(GridField $gridField, $request = null)
    {
        return $this->genericHandle('CsvDataFormatter', 'csv', $gridField, $request);
    }

    /**
     * Generic Handle request that will return a Spread Sheet in the requested format
     * @param  string    $dataFormatterClass
     * @param  string    $ext
     * @param  GridField $gridField
     * @param  HTTPRequest    $request
     * @return string
     */
    protected function genericHandle($dataFormatterClass, $ext, GridField $gridField, $request = null)
    {
        $items = $this->getItems($gridField);

        $this->setHeader($gridField, $ext);


        $formater = new $dataFormatterClass();
        $formater->setUseLabelsAsHeaders($this->useLabelsAsHeaders);

        $fileData = $formater->convertDataObjectSet($items);

        return $fileData;
    }

    /**
     * Set the HTTP header to force a download and set the filename.
     * @param GridField $gridField
     * @param string $ext Extension to use in the filename.
     */
    protected function setHeader($gridField, $ext)
    {
        $do = singleton($gridField->getModelClass());

        Controller::curr()->getResponse()
            ->addHeader(
                "Content-Disposition",
                'attachment; filename="' .
                $do->i18n_plural_name() .
                '.' . $ext . '"'
            );
    }

    /**
     * Helper function to extract the item list out of the GridField.
     * @param  GridField $gridField
     * @return SS_List
     */
    protected function getItems(GridField $gridField)
    {
        $gridField->getConfig()->removeComponentsByType('GridFieldPaginator');

        $items = $gridField->getManipulatedList();

        foreach ($gridField->getConfig()->getComponents() as $component) {
            if ($component instanceof GridFieldFilterHeader || $component instanceof GridFieldSortableHeader) {
                $items = $component->getManipulatedData($gridField, $items);
            }
        }

        $arrayList = ArrayList::create();

        foreach ($items->limit(null) as $item) {
            if (!$item->hasMethod('canView') || $item->canView()) {
                $arrayList->add($item);
            }
        }

        return $arrayList;
    }

    /**
     * Set the DataFormatter's UseFieldLabelsAsHeaders property
     * @param bool $value
     * @return GridFieldExcelExportButton
     */
    public function setUseLabelsAsHeaders($value)
    {
        $this->useLabelsAsHeaders = $value === null ? null : (bool)$value;

        return $this;
    }

    /**
     * Return the value that will be assigned to the DataFormatter's UseFieldLabelsAsHeaders property
     *
     * If null, will fallback on the default.
     *
     * @return bool|null
     */
    public function getUseLabelsAsHeaders()
    {
        return $this->useLabelsAsHeaders;
    }
}

<?php

/**
 * CsvDataFormatter extends {@link ExcelDataFormatter} to provide a DataFormatter
 * suitable for exporting an {@link SS_link} of {@link DataObjectInterface} to
 * a CSV spreadsheet.
 *
 * @author Firebrand <hello@firebrand.nz>
 * @license MIT
 * @package silverstripe-excel-export
 */
namespace ExcelExport;

use Override;
use SilverStripe\Model\List\SS_List;

class CsvDataFormatter extends ExcelDataFormatter
{

    /**
     * @inheritdoc
     */
    #[Override]
    public function supportedExtensions()
    {
        return [
            'csv',
        ];
    }

    /**
     * @inheritdoc
     */
    #[Override]
    public function supportedMimeTypes()
    {
        return [
            'text/csv',
        ];
    }

    /**
     * @inheritdoc
     */
    #[Override]
    public function convertDataObjectSet(SS_List $set)
    {
        $this->setHeader();

        $excel = $this->getPhpExcelObject($set);

        $fileData = $this->getFileData($excel, 'CSV');

        return $fileData;
    }
}


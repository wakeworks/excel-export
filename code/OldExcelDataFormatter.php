<?php

/**
 * OldExcelDataFormatter extends {@link ExcelDataFormatter} to provide a DataFormatter
 * suitable for exporting an {@link SS_link} of {@link DataObjectInterface} to
 * a Excel5 spreadsheet (XLS).
 *
 * @author Firebrand <hello@firebrand.nz>
 * @license MIT
 * @package silverstripe-excel-export
 */

namespace ExcelExport;

use Override;
use SilverStripe\Model\List\SS_List;

class OldExcelDataFormatter extends ExcelDataFormatter
{

    /**
     * @inheritdoc
     */
    #[Override]
    public function supportedExtensions()
    {
        return [
            'xls',
        ];
    }

    /**
     * @inheritdoc
     */
    #[Override]
    public function supportedMimeTypes()
    {
        return [
            'application/vnd.ms-excel',
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

        $fileData = $this->getFileData($excel, 'Excel5');

        return $fileData;
    }
}

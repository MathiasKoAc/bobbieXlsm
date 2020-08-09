<?php
//error_reporting(E_ALL);
//ini_set('display_errors', 1);
require __DIR__ . '/vendor/autoload.php';
require __DIR__ . '/models/UploadFileUserdata.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\IOFactory;
use models\UploadFileUserdata;

$loadUrl = "https://bau.bobbie.de/Product-Upload.xlsm";
//TODO in eine Verzeichnis ohne direkten Zugriff legen
//$tmpFileName = "load/Product-Upload-Manuel.xlsm";
$tmpFileName = "load/Product-Upload.xlsm";
$filePart = "Product-Upload";

$userdata = new UploadFileUserdata();
$userdata->seller_id = 172;
$userdata->export_api_key = "bobbieDeveloper";
$userdata->seller_prefix = "bobDev";
$userdata->base_url = "bau.bobbie.de";

if (!file_exists($tmpFileName)
    || (isset($_GET["reload"]) && $_GET["reload"] == "true")) {
    originalDateiLaden($loadUrl, $tmpFileName);
}

$spreadsheet = dateiLaden($tmpFileName);
dateiModifizieren($spreadsheet, $userdata);
$filename = namenErmittlen($spreadsheet, $userdata, $filePart);
dateiHerausgeben($spreadsheet, $filename);


/**
 * @param $spreadsheet
 * @param $userdata UploadFileUserdata
 * @param $filePart
 * @return string
 */
function namenErmittlen($spreadsheet, $userdata, $filePart) {
    $worksheet = $spreadsheet->getSheetByName("Hilfe!");
    $version = "0.0";
    $cell = $worksheet->getCell('I2');
    if(isset($cell)) {
        $version = $cell->getValue();
    }
    return $filePart ."_". $userdata->seller_prefix ."_". $version . ".xlsm";
}

/**
 * @param $spreadsheet Spreadsheet
 * @param $filename string
 */
function dateiHerausgeben($spreadsheet, $filename) {
    $writer = IOFactory::createWriter($spreadsheet, "Xlsx");
    $writer->setPreCalculateFormulas(false);

    header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    header('Content-Disposition: attachment;filename="'.$filename.'"');
    header('Cache-Control: max-age=0');

    $writer->save('php://output');
}

/**
 * @param $tmpFileName string
 * @return Spreadsheet
 * @throws \PhpOffice\PhpSpreadsheet\Reader\Exception
 */
function dateiLaden($tmpFileName) {
    $reader = IOFactory::createReader("Xlsx");
    return $spreadsheet = $reader->load($tmpFileName);
}

/**
 * @param $loadUrl string
 * @param $fileSaveName string
 */
function originalDateiLaden($loadUrl, $fileSaveName) {
    $file = file_get_contents($loadUrl);
    file_put_contents($fileSaveName, $file);
}

/**
 * @param &$spreadsheet Spreadsheet
 * @param $userdata UploadFileUserdata
 */
function dateiModifizieren(&$spreadsheet, $userdata) {
    $worksheet = $spreadsheet->getSheetByName('Werte');
    $worksheet->setCellValue('A1', 'SELLER_ID');
    $worksheet->setCellValue('B1', 'config');
    $worksheet->setCellValue('C1', $userdata->seller_id);

    $worksheet->setCellValue('A2', 'SELLER_PREFIX');
    $worksheet->setCellValue('B2', 'config');
    $worksheet->setCellValue('C2', $userdata->seller_prefix);

    $worksheet->setCellValue('A3', 'SELLER_NAME');
    $worksheet->setCellValue('B3', 'config');

    $worksheet->setCellValue('A4', 'EXPORT_API_KEY');
    $worksheet->setCellValue('B4', 'config');
    $worksheet->setCellValue('C4', $userdata->export_api_key);

    $worksheet->setCellValue('A5', 'BASE_URL');
    $worksheet->setCellValue('B5', 'config');
    $worksheet->setCellValue('C5', $userdata->base_url);
}
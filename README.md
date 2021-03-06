# ExcelPHPHelper
##Custom Helper class for ExcelPHP

Place the ExcelHelper.php in same place as PHPExcel.php

if using CakePHP:
add this line to your Controller
```
App::import('Vendor','PHPExcel',array('file' => 'excel/ExcelHelper.php'));
```

Some examples:
```
$excel = new ExcelHelper("TitleOfExcel", "MyWorkSheetName");

$sheet1 = $excel->getSheetHelper("MyWorkSheetName");

$sheet2 = $excel->newSheet("Sheet2Name");

$sheet1->write("Hi there"); //default String type
$sheet1->write(3, PHPExcel_Cell_DataType::TYPE_NUMERIC);

$sheet1->wrap()->write("This is a very long string and the height of cell is wrapped")->wrap()->write("chain writing to next column cell");

$sheet1->nextRow()->nextRow(); //move the pointer to next 2 row with column 0

$sheet1->setBgColor(null,"yellow");

$sheet1->setBgColor()->setTextColor(null,"dark_grey")->write("Amount:");

$sheet1->setAlign("A1","right");

$sheet1->setColumnWidth("A", 15);
```
Get the PHPExcel Worksheet Object, so official methods could be called.
```
$excel->getPHPExcel();

$sheet1->getSheet();
```

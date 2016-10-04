<?php
/*
* This is a custom helper class to facilitate the use of PHPExcel for self use only.
* No need take this class serious. Of course you can use if you find this class helpful.
*
* v0.1 2016-09-26
*
* Authors: nubotz (Kel)
*/
require_once('PHPExcel.php');

class ExcelHelper{
	private $objPHPExcel;
	private $sheets;

	//constructor
	function __construct($titleName, $defaultSheetName="WorkSheet"){
		$this->objPHPExcel = new PHPExcel();
		$this->objPHPExcel->getProperties()->setCreator("Owner")
				 	->setLastModifiedBy("Owner")
				 	->setTitle($titleName)
				 	->setSubject($titleName);
		$this->objPHPExcel->setActiveSheetIndex(0);
		$defaultSheet = $this->objPHPExcel->getActiveSheet();
		$defaultSheet->setTitle($defaultSheetName);
		$defaultSheetHelper = new SheetHelper($defaultSheet);
		$this->sheets = array($defaultSheetName => $defaultSheetHelper);
	}
	public function getPHPExcel(){
		return $this->objPHPExcel;
	}

	public function getSheetHelper($sheetName){
		return $this->sheets[$sheetName];

	}
	public function newSheet($sheetName=null){
		$newSheet = $this->objPHPExcel->createSheet();
		if($sheetName != null){
			$newSheet->setTitle($sheetName);
		}
		$newSheetHelper = new SheetHelper($newSheet);
		$this->sheets[$sheetName] = $newSheetHelper;
		return new SheetHelper($newSheet);
	}
	public function save($filePath){
		$objWriter = new PHPExcel_Writer_Excel2007($this->objPHPExcel);	
		try{
			$objWriter->save($filePath);
			return true;
		}catch(Exception $e){
			debug($e->getMessage());
			return false;
		}
	}
}

class SheetHelper{
	private $mySheet;
	private $myPointer;

	function __construct($sheet){
		$this->mySheet = $sheet;
		$this->myPointer = array("col"=>0,"row"=>1);
	}

	public function setHeader($headerArr=array()){
		foreach($headerArr as $value){
			$this->setBold("");
			$this->write($value);
		}
		$this->nextRow();
	}
	public function write($value, $type=TYPE_STR){
		$pointer = $this->myPointer;
		$sheet = $this->mySheet;
		$sheet->setCellValueExplicitByColumnAndRow($pointer["col"],$pointer["row"], $value, $type);
		$col = $pointer["col"]+1;
		$row = $pointer["row"];
		$this->myPointer = array("col"=>$col,"row"=>$row);
	}
	public function writeArray($array=array()){
		foreach($array as $value){
			$this->write($value);
		}
		$this->nextRow();
	}
	//write specific cell with third arg the cell location
	//without moving the pointer
	public function jumpWriteCell($coord="A1", $value="", $type=TYPE_STR){
		$this->mySheet->setCellValueExplicit($coord, $value, $type);
	}
	public function jumpWrite($col=0, $row=0, $value="", $type=TYPE_STR){
		$this->mySheet->setCellValueExplicitByColumnAndRow($col, $row, $value, $type);
	}
	public function setAlign($coord=null,$alignment="right"){
		$align_map=array(
			"right"=>PHPExcel_Style_Alignment::HORIZONTAL_RIGHT,
			"left"=>PHPExcel_Style_Alignment::HORIZONTAL_LEFT,
			"center"=>PHPExcel_Style_Alignment::HORIZONTAL_CENTER
			);
		if($coord == null){
			$this->getStyle()->getAlignment()->setHorizontal($align_map[$alignment]);
		}else{
			$this->mySheet->getStyle($coord)->getAlignment()->setHorizontal($align_map[$alignment]);
		}
	}
	public function setBold($coord=null){
		$style;
		if($coord == null){
			$pointer = $this->myPointer;
			$style = $this->mySheet->getStyleByColumnAndRow($pointer["col"],$pointer["row"]);
		}else{
			$style = $this->mySheet->getStyle($coord);
		}
		$style->getFont()->setBold(true);
		return $this;
	}
	public function setBorder($coord=null){
		$styleArray =          array(
             'allborders' => array(
                 'style' => PHPExcel_Style_Border::BORDER_THIN
             )
         );
		$style;
		if($coord == null){
			$pointer = $this->myPointer;
			$style = $this->mySheet->getStyleByColumnAndRow($pointer["col"],$pointer["row"]);
		}else{
			$style = $this->mySheet->getStyle($coord);
		}
		$style->getBorders()->applyFromArray($styleArray);
		return $this;
	}
	//if first parameter is empty string, then just use current cell
	public function setBgColor($coord=null,$color="pink"){
		$rgb_map = array(
			"pink"=>"F28A8C",
			"black"=>"000000",
			"white"=>"FFFFFFFF",
			"yellow"=>"FFFF00",
			"blue"=>"3399FF",
			"grey"=>"A9A9A9",
			"dark_grey"=>"464646"
			);
		$myBgColorStyle = array(
						'fill' => array(
							'type' => PHPExcel_Style_Fill::FILL_SOLID,
							'color' => array('rgb' => $rgb_map[$color])
						));
		$style;
		if($coord == null){
			$pointer = $this->myPointer;
			$style = $this->mySheet->getStyleByColumnAndRow($pointer["col"],$pointer["row"]);
		}else{
			$style = $this->mySheet->getStyle($coord);
		}
		$style->applyFromArray($myBgColorStyle);
		return $this;
	}
	public function setTextColor($coord=null,$color="grey"){
		$rgb_map = array(
			"pink"=>"F28A8C",
			"black"=>"000000",
			"white"=>"FFFFFFFF",
			"yellow"=>"FFFF00",
			"blue"=>"3399FF",
			"grey"=>"A9A9A9",
			"dark_grey"=>"464646"
			);
		$styleArray = array(
						'font' => array(
							'color' => array('rgb' => $rgb_map[$color])
						));

						// 'font'  => array(
						// 	'bold'=>true,
						// 	'color'=>array('rgb' => 'FF0000'),
						// 	'size'=>15,
						// 	'name'=>'Verdana')
		$style;
		if($coord == null){
			$pointer = $this->myPointer;
			$style = $this->mySheet->getStyleByColumnAndRow($pointer["col"],$pointer["row"]);
		}else{
			$style = $this->mySheet->getStyle($coord);
		}
		$style->applyFromArray($styleArray);
		return $this;
	}
	//set format code for the current cell
	public function setFormat($format, $formatCode=null){
		$format_map = array(
			""=>'for custom format code',
			"date"=>'yyyy/mm/dd;@',
			"dollar"=>'#,##0.00;[Red]-#,##0.00'
			);
		$pointer = $this->myPointer;
		$style = $this->mySheet->getStyleByColumnAndRow($pointer["col"],$pointer["row"]);

		if($formatCode == null){
			$style->getNumberFormat()->setFormatCode($format_map[$format]);
		}else{
			$style->getNumberFormat()->setFormatCode($formatCode);
		}
		return $this;
	}
	public function setColumnWidth($column="A", $width=15){
		$this->mySheet->getColumnDimension($column)->setWidth($width);
		return $this;
	}
	public function nextRow(){
		$pointer = $this->myPointer;
		$sheet = $this->mySheet;
		$col = 0;
		$row = $pointer["row"]+1 ;
		$this->myPointer = array("col"=>$col,"row"=>$row);
		return $this;
	}
	public function getCell(){
		$pointer = $this->myPointer;
		return $this->mySheet->getCellByColumnAndRow($pointer["col"],$pointer["row"]);
	}
	public function getStyle(){
		$pointer = $this->myPointer;
		return $this->mySheet->getStyleByColumnAndRow($pointer["col"],$pointer["row"]);
	}
	public function setPointer($col, $row){
		$this->myPointer = array("col"=>$col,"row"=>$row);
	}
	public function getPointer(){
		return $this->myPointer;
	}
	public function getSheet(){
		return $this->mySheet;
	}
}

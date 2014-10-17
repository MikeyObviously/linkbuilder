<?php

/** Error reporting */
error_reporting(E_ALL);
ini_set('display_errors', TRUE);
ini_set('display_startup_errors', TRUE);
define('EOL',(PHP_SAPI == 'cli') ? PHP_EOL : '<br />');

/** PHPExcel */
require_once 'Classes/PHPExcel.php';
require_once 'Classes/PHPExcel/IOFactory.php';

/** Functions */
require_once 'zip.php';
require_once 'autolink.php';

if ($_FILES["file"]["error"] > 0) {
  echo "Error: " . $_FILES["file"]["error"] . "<br>";
} 

/** Create Excel Object **/
$inputFileName = $_FILES["file"]["tmp_name"];
$objPHPExcel = PHPExcel_IOFactory::load($inputFileName);
$objWorksheet = $objPHPExcel->getActiveSheet();

/** Get how many rows **/
$highestRow = $objWorksheet->getHighestRow();

/** Keeps track of chapters to make new files on change **/
$chapter = 0;

/** For keeping track of the files created to zip & delete. **/
$files_to_zip = array();

/** Iterates through every row, writing data from cells into HTML files. **/
for ($row = 2; $row <= $highestRow; ++$row) {

	/** Checks first column to see if the chapter has changed. 
	If it has, the current file will be closed and a new one opened. **/
	if ($chapter != $objWorksheet->getCellByColumnAndRow(0, $row)->getValue()){
		fwrite ($f, "</div>");
		fclose($f);
		$filename = 'chapter' . $objWorksheet->getCellByColumnAndRow(0, $row)->getValue() . '.htm';
		$f = fopen($filename, 'a');
		/** Add to the list of files to zip and delete **/
		array_push($files_to_zip, $filename);
		fwrite ($f, "<div id='#links'>\n");
		fwrite ($f, "<h2>Chapter " . $objWorksheet->getCellByColumnAndRow(0, $row)->getValue() . "</h2>\n");
		/** Update chapter value **/
		$chapter = $objWorksheet->getCellByColumnAndRow(0, $row)->getValue();
	}
	/** Write links **/
	$autolinked_description = auto_link_text($objWorksheet->getCellByColumnAndRow(3, $row)->getValue());
	$data = 
		"<p><strong><a target='_blank' href='" . 
		$objWorksheet->getCellByColumnAndRow(2, $row)->getValue() . 
		"'>" . 
		$objWorksheet->getCellByColumnAndRow(1, $row)->getValue() . 
		"</a></strong><br>" . 
		$autolinked_description;
	
	fwrite ($f, $data);
}

fclose($f);

/** Sample style sheet **/
$f = fopen('style-sample.css', 'w');
fwrite ($f, "#links p, #links ol li {\n
	font-family: georgia, serif;\n
	line-height: 1.5em;\n
}\n
\n
#links ol li {\n
	margin-bottom:10px;\n
}\n
\n
#links p a {\n
	font-family: arial, san-serif;\n
	font-size: 110%;\n
}");
fclose($f);
array_push($files_to_zip,'style-sample.css');

/** Removes blank chapter.htm file. Still not sure why it is being created. 
The if statement is evaluating as true on the first run through (even when $chapter = 1) **/
if(($key = array_search('chapter.htm', $files_to_zip)) !== false) {
    unset($files_to_zip[$key]);
	unlink('chapter.htm');
}

$zip = create_zip($files_to_zip, 'weblinks.zip');

foreach($files_to_zip as &$del){
	unlink($del);
}

header("Content-disposition: attachment; filename=weblinks.zip");
header("Content-type: application/zip");
readfile("weblinks.zip");
unlink("weblinks.zip");

?>

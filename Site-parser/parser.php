<?php 
require_once __DIR__ . '/vendor/autoload.php';
//wrapper for domrequests jquery-style
use PhpQuery\PhpQuery;

//for spreadsheets readwrite
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

//going through links to parse info from. vals from the spreadsheet
$reader = \PhpOffice\PhpSpreadsheet\IOFactory::createReader('Xlsx');
$reader->setReadDataOnly(TRUE);
$spreadsheet = $reader->load("links.xlsx");
$worksheet = $spreadsheet->getActiveSheet();

//spreadsheet for collecting info + pics
$spreadsheetWrite = new Spreadsheet();

foreach ($worksheet->getRowIterator() as $row) {
    $cellIterator = $row->getCellIterator();
    $cellIterator->setIterateOnlyExistingCells(true); 
    foreach ($cellIterator as $cell) {
            linkGetData($cell->getValue(), $spreadsheetWrite); 
    }
}

//save images etc.
$writer = new Xlsx($spreadsheetWrite);
$writer->save('data.xlsx');


/*collect the data from $url and the subsequent pages. Upgrade the virt. sheet $sheet*/
function linkGetData($url, $sheet) {
	$curl = curl_init();
	curl_setopt($curl, CURLOPT_URL, $url);
	curl_setopt($curl, CURLOPT_CONNECTTIMEOUT, 20);
	curl_setopt($curl, CURLOPT_USERAGENT, 'Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/97.0.4692.99 Safari/537.36');
	curl_setopt($curl, CURLOPT_RETURNTRANSFER, true);
	curl_setopt($curl, CURLOPT_FOLLOWLOCATION, true);
	$str = curl_exec($curl); 
	curl_close($curl);
	
	$page = new PhpQuery;
	$page->load_str($str);
	
	//find the number pages and entries
	preg_match_all('!\d+!', $page->innerHTML($page->query('.pgCount')[0]), $arr);
	$currentAmount = ($arr[0][1] - $arr[0][0] + 1);
	$maxAmount = $arr[0][2];
	
	//get the trip name
	$tripName = $page->innerHTML($page->query('#HEADING')[0]);
	
	//filling the sheet
	$highestRow = $sheet->getActiveSheet()->getHighestDataRow();
	$sheet->getActiveSheet()->setCellValueByColumnAndRow(1, $highestRow, $tripName);
	
	$tourInfo; 
	
	//get pictures+names to save in .xlsx spreadsheet later on
	for($i = 0; $i < $currentAmount; $i++) {
		$review = $page->innerHTML($page->query('.more')[$i]);
		preg_match('/<a[^>]*>(.*?)<\/a>/', $review, $res);
		if(isset($res[1])) {
			preg_match('!\d+!', $res[1], $arr);
			if($arr[0] < 100) continue;
			$name;
			//jquery-style images
			if(!(str_contains($page->innerHTML($page->query('.near_listing_content')[$i]), 'photo_image'))) {
				$highestRow++; 
				$name = "no_pic.png";
			} else {
				$source = $page->outerHTML($page->query('.photo_image')[$i]);
				preg_match('/<img.*?src=[\'"](.*?)[\'"].*?>/i', $source, $arrPict);
				$ch = curl_init($arrPict[1]);
				preg_match('/.*\/(.*)/', $arrPict[1], $names);
				$name = $names[1];
				if($name !== "") {
					$fp = fopen($name, 'w+b');
					curl_setopt($ch, CURLOPT_FILE, $fp);
					curl_setopt($ch, CURLOPT_HEADER, 0);
					curl_exec($ch);
					fclose($fp);
					$highestRow++;
				}		else $name = 'no_pic.png';
					curl_close($ch);
			}
				//saving information to .xlsx 
				$drawing = new \PhpOffice\PhpSpreadsheet\Worksheet\Drawing();
				$drawing->setPath($name); // path to image
				$drawing->setCoordinates('B' . $highestRow);
				$drawing->setWidthAndHeight(50, 50);
				$drawing->setWorksheet($sheet->getActiveSheet());
			
				//jquering names of atractions + removing <a> tag
				$val = $page->innerHtml($page->query('.location_name')[$i]);
				preg_match('/<a[^>]*>(.*?)<\/a>/', $val, $res);
				$res[1] = htmlspecialchars_decode($res[1]);
				$sheet->getActiveSheet()->setCellValueByColumnAndRow(1, $highestRow, $res[1]);
				$sheet->getActiveSheet()->setCellValueByColumnAndRow(3, $highestRow, $arr[0]);
		}
	
	//for REGEX method
	/* $highestRow = $sheet->getActiveSheet()->getHighestDataRow();
	$sheet->getActiveSheet()->setCellValueByColumnAndRow(1, $highestRow, $tripName);
	//get pictures and save in .xlsx sreadsheet
	foreach($elements as $element) {
		$ch = curl_init($element);
		$posLast = strrpos($element, '/'); 
		if(str_contains($element, '.jpg') || str_contains($element, '.jpeg') || str_contains($element, '.png')) {
			$highestRow++;
			$name = substr($element, ($posLast+1));
			$fp = fopen($name, 'w+b');
			curl_setopt($ch, CURLOPT_FILE, $fp);
			curl_setopt($ch, CURLOPT_HEADER, 0);
			curl_exec($ch);
			fclose($fp);
			$drawing = new \PhpOffice\PhpSpreadsheet\Worksheet\Drawing();
			$drawing->setPath($name); // put your path and image here
			$drawing->setCoordinates('B' . $highestRow);
			$drawing->setWidthAndHeight(50, 50);
			$drawing->setWorksheet($sheet->getActiveSheet());
		}
		curl_close($ch);
	} */
	
	
}
}
?>
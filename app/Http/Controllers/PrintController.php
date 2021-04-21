<?php

namespace App\Http\Controllers;

use Illuminate\Http\Request;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\Writer\Xls;
// use PhpOffice\PhpSpreadsheet\IOFactory;
use App\Models\PrintTable;
class PrintController extends Controller
{
    //

    public function index(){
		// $emp = Emp::paginate(5);
		// return view('employee', ['emp' => $emp]);
		// $printtable = new PrintTable;
		// dd(111);
		return view('');

	}	

	public function printtable(){

		$printtable = new PrintTable;

		// $users = PrintTable::all()->sortBy('point_id')->get();
		$d=strtotime("-3 months");
		$startdate = date("Y-m-d h:i:s", $d);

		$d2=strtotime("-2 months");
		$startdate_2 = date("Y-m-d h:i:s", $d2);

		$d3=strtotime("-1 month");
		$startdate_3 = date("Y-m-d h:i:s", $d3);

		$d4 = strtotime("today");
		$today = date("Y-m-d h:i:s", $d4);

		// dd($today);

		// $e=strtotime("today");
		// echo ($e-$d);

		$users_3 = printtable::where('hyosyo_from', '>', $startdate)->where('hyosyo_from', '<' , $startdate_2)->orderBy('created')->orderBy('point_id')->get();
		$users_2 = printtable::where('hyosyo_from', '>', $startdate_2)->where('hyosyo_from', '<' , $startdate_3)->orderBy('created')->orderBy('point_id')->get();
		$users_1 = printtable::where('hyosyo_from', '>', $startdate_3)->orderBy('created')->orderBy('point_id')->get();


		// var_dump($users_3);
		// exit();	
		// return;
		// $employees = Emp::all();
		// $objPHPExcel = PHPExcel_IOFactory::load("excelTemplate.xls");
		$inputFileName = storage_path("table.xlsx");

		/** Load $inputFileName to a Spreadsheet Object  **/
		$spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load($inputFileName);

		// $spreadsheet = new Spreadsheet();
        $sheet = $spreadsheet->getActiveSheet();
                    
		$rows = 3;
		foreach($users_3 as $user){ 
			$sheet->setCellValue('A' . $rows, $user->point_id);
            $sheet->setCellValue('C' . $rows, $user->receivePoint);
            $sheet->setCellValue('D' . $rows, $user->sendPoint);
            $sheet->setCellValue('E' . $rows, $user->messageCount);
			$sheet->setCellValue('F' . $rows, $user->messageCount+$user->receivePoint+$user->sendPoint);
            
            $rows++;		
		}
		$rows = 3;
		foreach($users_2 as $user){ 
            $sheet->setCellValue('G' . $rows, $user->receivePoint);
            $sheet->setCellValue('H' . $rows, $user->sendPoint);
			$sheet->setCellValue('I' . $rows, $user->messageCount);
            $sheet->setCellValue('J' . $rows, $user->messageCount+$user->receivePoint+$user->sendPoint);
            
            $rows++;		
		}
		$rows = 3;

		foreach($users_1 as $user){ 
            $sheet->setCellValue('K' . $rows, $user->receivePoint);
            $sheet->setCellValue('L' . $rows, $user->sendPoint);
			$sheet->setCellValue('M' . $rows, $user->messageCount);
            $sheet->setCellValue('N' . $rows, $user->messageCount+$user->receivePoint+$user->sendPoint);
            $rev_total = $spreadsheet->getActiveSheet()->getCell('C'.$rows)->getValue()+$spreadsheet->getActiveSheet()->getCell('G'.$rows)->getValue()+$spreadsheet->getActiveSheet()->getCell('K'.$rows)->getValue();
            $sheet->setCellValue('O' . $rows, $rev_total);

            $send_total = $spreadsheet->getActiveSheet()->getCell('D'.$rows)->getValue()+$spreadsheet->getActiveSheet()->getCell('H'.$rows)->getValue()+$spreadsheet->getActiveSheet()->getCell('L'.$rows)->getValue();
            $sheet->setCellValue('P' . $rows, $send_total);

            $msg_total = $spreadsheet->getActiveSheet()->getCell('E'.$rows)->getValue()+$spreadsheet->getActiveSheet()->getCell('I'.$rows)->getValue()+$spreadsheet->getActiveSheet()->getCell('M'.$rows)->getValue();
            $sheet->setCellValue('Q' . $rows, $msg_total);

            $overall_total = $spreadsheet->getActiveSheet()->getCell('F'.$rows)->getValue()+$spreadsheet->getActiveSheet()->getCell('J'.$rows)->getValue()+$spreadsheet->getActiveSheet()->getCell('N'.$rows)->getValue();
            $sheet->setCellValue('R' . $rows, $overall_total);


            $rows++;		
		}
		// $cellValue = $spreadsheet->getActiveSheet()->getCell('A5')->getValue();
		// dd($cellValue);



	    $fileName = "table."."xlsx";
		// if($type == 'xlsx') {
		// 	$writer = new Xlsx($spreadsheet);
		// } else if($type == 'xls') {
		// 	$writer = new Xls($spreadsheet);			
		// }		
		$writer = new Xlsx($spreadsheet);
		$writer->save(storage_path("app\\".$fileName));
		header("Content-Type: application/vnd.ms-excel");
        return redirect(url('/')."/export/".$fileName);  

		// foreach ($users_3 as $user) {
  //   		$t_point_id = $user->point_id;
  //   		$t_created = $user->created;
  //   		$t_hyosyo_id = $user->hyosyo_id;
  //   		$t_hyosyo_from = $user->hyosyo_from;
  //   		$t_hyosyo_to = $user->hyosyo_to;
  //   		$t_no = $user->no;
  //   		$t_person_id = $user->person_id;
  //   		$t_point = $user->point;
  //   		$t_messageCount = $user->messageCount;
  //   		$t_receivePoint = $user->receivePoint;
  //   		$t_sendPoint = $user->sendPoint;
  //   		$t_created = $user->created;
  //   		$t_modified = $user->created;
    		// $t_messageCount = $user->messageCount;


    		// $result = $table_created->format('Y-m-d-H-i-s');
			// var_dump($t_created);
    	// }
		







		// dd($user);


		
	}	




	


}

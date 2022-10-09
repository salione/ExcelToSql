<?php
include 'excel_reader.php';     // include the class
error_reporting(1);
function con_db(){
    $pdo = new PDO('mysql:dbname=tolidatmelli;host=localhost', 'root', '');
   // $pdo->exec("set names utf8");
    return $pdo;
}
// creates an object instance of the class, and read the excel file data
$excel = new PhpExcelReader;
$excel->setOutputEncoding("UTF-8");
$excel->read('test.xls');
$pdo = con_db();
// Excel file data is stored in $sheets property, an Array of worksheets
/*
The data is stored in 'cells' and the meta-data is stored in an array called 'cellsInfo'

Example (firt_sheet - index 0, second_sheet - index 1, ...):

$sheets[0]  -->  'cells'  -->  row --> column --> Interpreted value
         -->  'cellsInfo' --> row --> column --> 'type' (Can be 'date', 'number', or 'unknown')
                                            --> 'raw' (The raw data that Excel stores for that data cell)
*/

// this function creates and returns a HTML table with excel rows and columns data
// Parameter - array with excel worksheet data
function sheetData($sheet) {
   echo $sheet['cells'][1][1];
//  $re = '<table>';     // starts html table

  $x = 65536;
  while($x <= 70634) {
//      for($x=1;$x++;$x<190000){
//    $re .= "<tr>\n";
//          if(!empty($sheet['cells'][$x][1])){


    $y = 1;
//    while($y <= $sheet['numCols']) {
  //    $cell = isset($sheet['cells'][$x][$y]) ? $sheet['cells'][$x][$y] : '';
//      $re .= " <td>$cell</td>\n";
      $pdo = con_db();
      $sth2 = $pdo->prepare('UPDATE  `attributesvalue`  SET
      `title_en`="'.$sheet['cells'][$x][2].'" where id = "'.$sheet['cells'][$x][1].'" ');
    //  $sth2->execute();

        echo $sheet['cells'][$x][1].'->';
        echo $sheet['cells'][$x][2].'-> imported<br/>';

      $y++;
//    }
//    $re .= "</tr>\n";
    $x++;
//          }
  }

//  return $re .'</table>';     // ends and returns the html table
}

$nr_sheets = count($excel->sheets);       // gets the number of sheets
$excel_data = '';              // to store the the html tables with data of each sheet

// traverses the number of sheets and sets html table with each sheet data in $excel_data
for($i=0; $i<$nr_sheets; $i++) {
  $excel_data .= '<h4>Sheet '. ($i + 1) .' (<em>'. $excel->boundsheets[$i]['name'] .'</em>)</h4>'. sheetData($excel->sheets[$i]) .'<br/>';  
}
?>
<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8" />
<title>Example PHP Excel Reader</title>
<style type="text/css">
table {
 border-collapse: collapse;
}        
td {
 border: 1px solid black;
 padding: 0 0.5em;
}        
</style>
</head>
<body>

<?php
// displays tables with excel file data
echo $excel_data;
?>    

</body>
</html>

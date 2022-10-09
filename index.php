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

function sheetData($sheet) {
   echo $sheet['cells'][1][1];
   $x = 65536;
   while($x <= 70634) {
    $y = 1;
      $pdo = con_db();
      $sth2 = $pdo->prepare('UPDATE  `attributesvalue`  SET
      `title_en`="'.$sheet['cells'][$x][2].'" where id = "'.$sheet['cells'][$x][1].'" ');
      $sth2->execute();
        echo $sheet['cells'][$x][1].'->';
        echo $sheet['cells'][$x][2].'-> imported<br/>';

    $x++;
  }

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

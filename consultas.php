<?php 
//Script para conexión con base de datos en Mysql
include("db_controller/mysql_script.php");

// Obtenemos parametros
$obj = (object)$_REQUEST;

switch ($obj->action) {
  case 'exportToEXCEL':
  	// 
  	$rows = query("SELECT * FROM personal ORDER BY nombres ASC LIMIT 100");
	
	$filename = "Reporte en EXCEL.xls";

  	header("Pragma: public");
	header("Expires: 0");
	header("Content-type: application/x-msdownload");
	header("Content-Disposition: attachment; filename=$filename");
	header("Pragma: no-cache");
	header("Cache-Control: must-revalidate, post-check=0, pre-check=0");

	$html = '<table style="margin:auto;" border="1">
				<tr>
					<th colspan="4" align="center">NOMBRES MAS FRECUENTES DE MUJERES EN ESPAÑA</th>
				</tr>
				<tr>
					<th>Nro</th>
					<th>Nombres</th>
					<th>Fecuencia</th>
					<th>Serie</th>
				</tr>
		';
	    
    foreach ($rows as $key => $row) {
    	$html.= '<tr>
    				<td>'.( (int)$key+1 ).'</td>
    				<td>'.$row['nombres'].'</td>
    				<td>'.$row['frecuencia'].'</td>
    				<td>'.$row['serie'].'</td>
				</tr>
    			';
    }
			
	$html.='</table>';
	echo $html;
  	break;
  	
}
?>
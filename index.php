<?php 
ini_set('display_errors', 1);
error_reporting(E_ERROR);
$instancia = $_GET["instancia"];
require_once "libs/vendor/autoload.php";
require "PHPMailer/PHPMailer.php";
require "PHPMailer/SMTP.php";
require "PHPMailer/Exception.php";
use PhpOffice\PhpSpreadsheet\IOFactory;
use PHPMailer\PHPMailer\PHPMailer;
use PHPMailer\PHPMailer\SMTP;
use PHPMailer\PHPMailer\Exception;

//Configuración SMTP
$config["smpt_secure"]         = "ssl";
$config["smpt_server"]         = "mail.vicom.mx"; // sets the SMTP server
$config["smpt_port_conection"] = 465;// set the SMTP port for the GMAIL server
$config["smpt_user"]           = "no-reply@vicom.mx"; // SMTP account username
$config["smpt_pwd"]            = "OC8wV^HDRs!G";// SMTP account password
$config["smpt_name_from"]      = "Vicom";
$config["subject_title"]       = "Vicom | Axo Files Validate";
$config["to"]                  = "arturohrdez@gmail.com,arturohrdez@outlook.com";


if(empty($instancia)){
    die("<h3>La instancia es requerida</h3>");
}else{
    //Busca archivo para procesar
    $dir = getcwd()."/".$instancia;
    chdir($dir);
    $files_xlsx = glob("*.xlsx");
    $files_xls  = glob("*.xls");

    //Valida si existe archivos con extensión "xlsx" y "xls"
    $response_file_valid = [];
    $decode_error        = "";
    $path_files          = [];

    if(!empty($files_xlsx) || !empty($files_xls)){
        if(!empty($files_xlsx)){
            foreach ($files_xlsx as $file_item_xlsx) {
                $response_file_valid_xlsx[] = validaArchivo($file_item_xlsx);
            }//end foreach
        }//end if


        if(!empty($files_xls)){
            foreach ($files_xls as $file_item_xls) {
                $response_file_valid_xls[] = validaArchivo($file_item_xls);
            }
        }//end if

        //Muestra errores en archivos procesados "xlsx"
        if(!empty($response_file_valid_xlsx)){
            foreach ($response_file_valid_xlsx as $res_item_xlsx) {
                $err       = $res_item_xlsx["err"];
                $arr_error = $res_item_xlsx["arr_error"];
                $file_name = $res_item_xlsx["file_name"];

                //No hubo error el archivo es válido
                if($err == 0){
                    rename($dir."/".$file_name, $dir."/procesados/".$file_name);
                    $decode_error .= decodeError($err,$arr_error,$file_name);
                    $path_files[] = $dir."/procesados/".$file_name;
                }//end if

                //Error 1 [Faltán columnas obligatorias]
                if($err == 1){
                    rename($dir."/".$file_name, $dir."/erroneos/".$file_name);
                    $decode_error .= decodeError($err,$arr_error,$file_name);
                    $path_files[] = $dir."/erroneos/".$file_name;
                }//end if

                //Error 2 [Falta información en requerida en el contenido del excel]
                if($err == 2){
                    rename($dir."/".$file_name, $dir."/erroneos/".$file_name);
                    $decode_error .= decodeError($err,$arr_error,$file_name);
                    $path_files[] = $dir."/erroneos/".$file_name;
                }//end if

                //Error 404 [No se encontro el archivo]
                if($err == 404){
                    $decode_error .= decodeError($err,$arr_error,$file_name);
                }//end if

            }//end foreach
        }//end if

        //Muestra errores en archivos procesados "xls"
        if(!empty($response_file_valid_xls)){
            foreach ($response_file_valid_xls as $res_item_xls) {
                $err       = $res_item_xls["err"];
                $arr_error = $res_item_xls["arr_error"];
                $file_name = $res_item_xls["file_name"];

                //No hubo error el archivo es válido
                if($err == 0){
                    rename($dir."/".$file_name, $dir."/procesados/".$file_name);
                    $decode_error .= decodeError($err,$arr_error,$file_name);
                    $path_files[] = $dir."/procesados/".$file_name;
                }//end if

                //Error 1 [Faltán columnas obligatorias]
                if($err == 1){
                    rename($dir."/".$file_name, $dir."/erroneos/".$file_name);
                    $decode_error .= decodeError($err,$arr_error,$file_name);
                    $path_files[] = $dir."/erroneos/".$file_name;
                }//end if

                //Error 2 [Falta información en requerida en el contenido del excel]
                if($err == 2){
                    rename($dir."/".$file_name, $dir."/erroneos/".$file_name);
                    $decode_error .= decodeError($err,$arr_error,$file_name);
                    $path_files[] = $dir."/erroneos/".$file_name;
                }//end if

                //Error 404 [No se encontro el archivo]
                if($err == 404){
                    $decode_error .= decodeError($err,$arr_error,$file_name);
                }//end if
            }//end foreach
        }//end if

        //Manda notifiación vía Email
        $response_email = sendEmail($decode_error,$config,$path_files);

    }elseif(empty($files_xlsx) && empty($files_xls)){
        die("<h2>Sin archivo para procesar</h2>");
    }//end if
}//end if

function formatTitle($string){
    $str         = strtolower($string);
    $str         = str_replace(" ","_",$str);
    $originales  = 'ÀÁÂÃÄÅÆÇÈÉÊËÌÍÎÏÐÑÒÓÔÕÖØÙÚÛÜÝÞßàáâãäåæçèéêëìíîïðñòóôõöøùúûýýþÿ';
    $modificadas = 'aaaaaaaceeeeiiiidnoooooouuuuybsaaaaaaaceeeeiiiidnoooooouuuyyby';
    $str         = utf8_decode($str);
    $str         = strtr($str, utf8_decode($originales), $modificadas);
    return utf8_encode($str);
}//end function


function validaArchivo($file_name){
    $columnasRequeridas_titulos   = ["codigo_producto","titulo_producto","marca","sku","ean/upc","fecha_de_lanzamiento","departamento","categoria","talla","color","genero","deporte","edad"];
    $columnasRequeridas_contenido = ["A","B","C","D","E","G","L","M","O","P","Q","R","U"];

    if(file_exists($file_name)){
        $documento    = IOFactory::load($file_name);
        $totalDeHojas = $documento->getSheetCount();
        
        //Valida columnas requeridas
        $requeridosEncontrados = [];
        for ($indiceHoja = 0; $indiceHoja < $totalDeHojas; $indiceHoja++) {
            $hojaActual = $documento->getSheet($indiceHoja);

            $numeroMayorDeFila    = $hojaActual->getHighestRow();   // Numérico
            $letraMayorDeColumna  = $hojaActual->getHighestColumn();// Letra
            $numeroMayorDeColumna = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::columnIndexFromString($letraMayorDeColumna);

            //Valida que las columnas requeridas existan en el excel
            for ($indiceColumna = 1; $indiceColumna <= $numeroMayorDeColumna; $indiceColumna++) {
                $celda    = $hojaActual->getCellByColumnAndRow($indiceColumna, 2);
                $valorRaw = formatTitle($celda->getValue());
                if(in_array($valorRaw,$columnasRequeridas_titulos)){
                    $requeridosEncontrados[] = $valorRaw;
                }//end if
            }//end for
        }//end for

        $totalColRequeridos = count($columnasRequeridas_titulos);
        $totalColEncontrada = count($requeridosEncontrados);
        if($totalColRequeridos != $totalColEncontrada){
            $array_diff = array_diff($columnasRequeridas_titulos,$requeridosEncontrados);
            $arr_error_c = ["err"=>"1","arr_error"=>$array_diff,"file_name"=>$file_name];
            return $arr_error_c;
        }//end if

        //Valida el contenido requerido del archivo
        $array_error = [];
        for ($indiceFila = 3; $indiceFila <= $numeroMayorDeFila; $indiceFila++) {
            for ($indiceColumna = 1; $indiceColumna <= $numeroMayorDeColumna; $indiceColumna++) {
                $celda    = $hojaActual->getCellByColumnAndRow($indiceColumna, $indiceFila);
                $valorRaw = $celda->getFormattedValue();
                $fila     = $celda->getRow();
                $columna  = $celda->getColumn();

                if(empty($valorRaw)){
                    if(in_array($columna,$columnasRequeridas_contenido)){
                        $array_error[] = $columna." - ".$fila;
                    }//end if
                }//end if
            }//end for
        }//end for

        $c_error = count($array_error);
        if($c_error > 0){
            $arr_error_c = ["err"=>"2","arr_error"=>$array_error,"file_name"=>$file_name];
            return $arr_error_c;
        }else{
            $arr_error_c = ["err"=>"0","arr_error"=>null,"file_name"=>$file_name];
            return $arr_error_c;
        }//end if
    }else{
        $arr_error_c = ["err"=>"404","arr_error"=>null,"file_name"=>$file_name];
        return $arr_error_c;
    }//end if
}//end function

function decodeError($err,$arr_error,$file_name){
    $body = "";
    if($err == 0){
        $body .= "<h3> - El archivo : <strong style='color: green;'>".$file_name."</strong> es válido.</h3> <hr>";
        return $body;
    }//end if

    if($err == 1){
        $body .= "<h3> - El archivo : <strong style='color: green;'>".$file_name."</strong> <strong style='color: red;'> NO </strong> es válido, ";
        $body .= "<span> faltan algunas columnas requeridas en el excel que a continuación se listan :</span> </h3>";
        $body .=  "<ul>";
        foreach ($arr_error as $value) {
            $body .= "<li>".$value."</li>";
        }//end foreach
        $body .=  "</ul>";
        $body .=  "<hr>";

        return $body;
    }//end if

    if($err == 2){
        $body .= "<h3> - El archivo : <strong style='color: green;'>".$file_name."</strong> <strong style='color: red;'> NO </strong> es válido, ";
        $body .= "falta información requerida en el archivo, a continuación se listan columnas y filas con información faltante : </h3>";

        $body .= "
        <table>
            <tr>
                <td width='85' align='center'>Columna</td>
                <td width='75' align='center'>Fila</td>
            </tr>
        ";

        foreach ($arr_error as $value) {
            $arr_s = explode("-",$value);
            $body .= "<tr><td align='center'>".$arr_s[0]."</td><td align='center'>".$arr_s[1]."</td></tr>";
        }//end foreach

        $body .= "
        </table>
        <hr>
        ";

        return $body;
    }

    if($err == 404){
        $body .= "<h3> - El archivo con el nombre: <strong style='color: green;'>".$file_name."</strong> no se encontra en la ruta especificada.</h3> <hr>";
        return $body;
    }
}//end function

function sendEmail($HTML=null,$config=null,$path_files=null){
    $phpmailer             = new PHPMailer();
    $phpmailer->IsSMTP();
    $phpmailer->SMTPAuth   = true;// enable SMTP authentication
    $phpmailer->SMTPSecure = $config["smpt_secure"];
    $phpmailer->Host       = $config["smpt_server"]; // sets the SMTP server
    $phpmailer->Port       = $config["smpt_port_conection"];// set the SMTP port for the GMAIL server
    $phpmailer->Username   = $config["smpt_user"]; // SMTP account username
    $phpmailer->Password   = $config["smpt_pwd"];// SMTP account password
    $phpmailer->SMTPDebug  = 0;
    $phpmailer->From       = $config["smpt_user"] ;
    $phpmailer->FromName   = utf8_decode($config["smpt_name_from"]);
    $phpmailer->Subject    = utf8_decode($config["subject_title"]);
    $arr_to = explode(",",$config["to"]);
    foreach ($arr_to as $to_) {
        $phpmailer->AddAddress($to_);
    }//end foreach

    //Attachments
    if(!is_null($path_files)){
        foreach ($path_files as $file) {
            $phpmailer->addAttachment($file);
        }
    }//end if
    $phpmailer->MsgHTML(utf8_decode($HTML));
    if (!$phpmailer->send()) {
        echo 'Mailer Error: ' . $phpmailer->ErrorInfo;
    } else {
        echo 'Report Message sent!';
    }//end if
    die();
}//end function
?>
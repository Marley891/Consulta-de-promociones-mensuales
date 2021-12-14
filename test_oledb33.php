<?php

$conn = new COM("ADODB.Connection") or die ("No se puede crear la instancia de ADODB");;


//$conn->Open("Provider=vfpoledb;Data Source='C:/vsairesp/Sai.DBC';Collating Sequence=Machine;Exclusive=No;ReadOnly=true");
$conn->Open("Provider=vfpoledb;Data Source='//192.168.10.99/CAMSA/Sai.DBC';Collating Sequence=Machine;Exclusive=No;ReadOnly=true");





//insert some data
//$filename="/myfolder/myfile.jpg";
//$ma_sql='insert into archivo (arch_real) values ("'.$filename.'")';
//$conn->Execute($ma_sql);

// lets display records...
$ma_sql="

            Select
                   prom.Cve_prod,
                   prod.Desc_prod,
                   prom.Fecha_venc,
                   Exist
            from
                 promen prom
                 
            left join
                 Producto Prod On  prom.cve_prod = Prod.cve_prod

            left join  ( Select  sum(Existencia) as Exist , cve_prod from Existe  group by cve_prod ) Exis on Exis.Cve_prod = Prom.Cve_prod
                 
             where Fecha_venc >= date()


 ";


$rs = $conn->Execute($ma_sql);


$num_columns = $rs->Fields->Count();
$num_rows = $rs->recordCount();

echo "Numero de Columnas : ". $num_columns . "<br>";
echo "Numero de Renglones : ". $num_rows . "<br>";


for ($i=0; $i < $num_columns; $i++) {
    $fld[$i] = $rs->Fields($i);
}

for ($i=0; $i < $num_columns; $i++) {
        echo $fld[$i]->name . "\t";
    }
    echo "<br>";
    
for ($i=0; $i < $num_columns; $i++) {
        echo $fld[$i]->type . " - ";
    }
    echo "<br>";

for ($i=0; $i < $num_columns; $i++) {
        echo $fld[$i]->DefinedSize . "\t";
    }
    echo "<br>";


$rowcount = 0;
while (!$rs->EOF) {
    for ($i=0; $i < $num_columns; $i++) {
        $fld[$i] = $rs->Fields($i);
        
        echo $fld[$i]->value . " | ";
        //if  (is_object($fld[$i]->value)) echo "yes";
       }
    echo "<br>";
    $rowcount++;            // increments rowcount
    $rs->MoveNext();
}

//while (!$rs->EOF)
//{
    //$row = $rs->fetch();
	//.... print data ....
	//echo $row;
//    $rs->MoveNext();
//}

$rs->Close();
$conn->Close();
$rs = null;
$conn = null;

?>


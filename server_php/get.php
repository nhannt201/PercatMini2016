<?php
$name = $_GET["name"];
$content = $_SERVER['REMOTE_ADDR'];  
$handle = fopen ("user/".$content.".txt","ab");
fwrite ($handle, $name); 
?>
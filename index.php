<?php

declare(strict_types=1);

use App\Docx2Html;

require __DIR__ . '/vendor/autoload.php';


$docx2Html = new Docx2Html();

$d = dir(getcwd().'/docs');

while (($source = $d->read()) !== false){
    
    if(substr($source,0,5) == "(new)")
        continue;

    if(file_exists($source))
        continue;

    if(explode('.',$source)[1] != 'docx')
        continue;

    $sourcePath = getcwd().'/docs/'.$source;

    echo "filename: " . $source . "<br>";
    $docx2Html->addFile($source);

    echo $docx2Html->applyTemplate(
        $sourcePath, 
        $docx2Html->readFile($sourcePath)
    );
    
}
$d->close();

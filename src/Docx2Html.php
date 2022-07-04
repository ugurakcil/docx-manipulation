<?php
namespace App;

use DOMDocument;
use ZipArchive;

class Docx2Html
{
    protected $zip;
    protected $xml;
    protected $phpWord;
    protected $files = [];

    public function __construct()
    {
        $this->zip = new ZipArchive;
        $this->xml = new DOMDocument();
    }

    public function addFile($filename) 
    {
        $this->files[] = $filename;
    }

    public function getFiles()
    {
        return $this->files;
    }

    public function readFile($source)
    {
        if (true !== $this->zip->open($source)) 
            return false;
    
        if (($index = $this->zip->locateName("word/document.xml")) === false)
            return false;
        
        $data = $this->zip->getFromIndex($index);  
        $this->xml->loadXML($data, LIBXML_NOENT | LIBXML_XINCLUDE | LIBXML_NOERROR | LIBXML_NOWARNING);  
        
        $outWithWP = strip_tags($this->xml->saveXML(), '<w:p>'); 
        
        $this->zip->close(); 
        return preg_replace('/<w:p(.*?)>(.*?)<\/w:p>/',"$2\n",$outWithWP);
    }

    public function getContents()
    {
        $fileContents = [];
        foreach($this->files as $source) {
            $fileContents[] = $this->readFile($source);
        }
        return $fileContents;
    }

    public function applyTemplate($filePath, $docxTextFull)
    {
        $this->phpWord = new \PhpOffice\PhpWord\PhpWord();
        
        $docxParagraphList = explode("\n", $docxTextFull);

        $docxParagraphList = array_values(array_filter($docxParagraphList));

        $section = $this->phpWord->addSection();

        $docxText = "";
        $loop = 0;
        foreach($docxParagraphList as $docxParagraphRow) {
            if($loop == 0) {
                $section->addText(htmlspecialchars("<h1>{$docxParagraphRow}</h1>"));
                $docxText .= "<h1>{$docxParagraphRow}</h1>";
            }
            else {
                $section->addText(htmlspecialchars("<p>{$docxParagraphRow}</p><br><br>"));
                $docxText .= "<p>{$docxParagraphRow}</p><br><br>";
            }

            ++$loop;
        }
        
        $objWriter = \PhpOffice\PhpWord\IOFactory::createWriter($this->phpWord, 'Word2007');

        
        $filePathArr = explode("/", $filePath);
        $filePathArr[count($filePathArr) - 1] = "(new) ". $filePathArr[count($filePathArr) - 1];
        
        $objWriter->save(implode('/',$filePathArr));
        return $docxText;
    }

}
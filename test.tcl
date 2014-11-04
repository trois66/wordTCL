package req tdom
console show
update
proc addImgToDoc {docFileName} {
  puts "Démarrage de l'inclusion dans le fichier document.xml"
  set xmlData [dom parse [tDOM::xmlReadFile $docFileName]]
  
  set ns2 {rel urn:schemas-microsoft-com:vml w http://schemas.openxmlformats.org/wordprocessingml/2006/main}
  set rootFile [$xmlData documentElement]
  puts $rootFile
  puts [$rootFile nodeName]
  set lastParagraph [$xmlData selectNodes -namespaces $ns2 {/w:document/w:body}]
  
  
  set myPara [createTheParagraphPart $xmlData]
  #$lastParagraph appendChild $myPara
  
  
    
  puts "myparagraph"
  #puts [$myPara asXML]
  puts [$lastParagraph lastChild]
  #$lastParagraph insertBefore
  #$lastParagraph appendXML  "<w:p w:rsidR='0082274B' w:rsidRDefault='00896E90'/>"
  #puts [$lastParagraph asXML]
  #puts "Move    ----------------------"
  update
  #set move [$xmlData selectNodes -namespaces $ns2 {/w:document/w:body/w:p[last()]}]
  #puts [$move asXML]
  #puts "Move    ----------------------"
  #update
  
  #$xmlData removeChild $move
  #set nextPara [$lastParagraph insertBefore
  $lastParagraph insertBefore $myPara [$lastParagraph lastChild]
  puts "lastPara ===================="
  puts $lastParagraph 
  puts [$lastParagraph asXML]

}

proc createTheParagraphPart {doc} {
  global shapeId
  set myPara [$doc createElement "w:p"]
  $myPara setAttribute w:rsidR "0082274B"
  $myPara setAttribute w:rsidRDefault "00896E90"
  set myRElement [$doc createElement "w:r"]
  set myPict [$doc createElement "w:pict"]
  set myShape [$doc createElement "v:shape"]
  set textShapeId "tomRipShapeId$shapeId"
  $myShape setAttribute id $textShapeId
  $myShape setAttribute type "#_x0000_t75"
  #$myShape setAttribute style "width:3in;height:3in"
  set myImageData [$doc createElement "v:imagedata"]
  set textImageId "tomRipImageId$shapeId"
  $myImageData setAttribute r:id $textImageId
  $myShape appendChild $myImageData
  $myPict appendChild $myShape
  $myRElement appendChild $myPict
  $myPara appendChild $myRElement
  return $myPara
}

proc addImgToDocRel {docFileName} {
  puts "Démarrage de l'inclusion dans le fichier document.xml.rels"
  set xmlData [dom parse [tDOM::xmlReadFile $docFileName]]
  set ns {rel http://schemas.openxmlformats.org/package/2006/relationships}
  set rootFile [$xmlData documentElement]
  puts [$rootFile asXML]
  set lastParagraph [$xmlData selectNodes -namespaces $ns {/rel:Relationships}]
  puts "lastParaFromRel"
  puts [$lastParagraph asXML]
  update
  set myPara [createTheRelationshipsPart $xmlData "toto.gif"]
  $lastParagraph appendChild $myPara
  puts "++++++++++++++++++++++++++++++++++++++++++++++"
  puts [$lastParagraph asXML]
}



proc createTheRelationshipsPart {doc filename} {
  global shapeId
  set myPara [$doc createElementNS "http://schemas.openxmlformats.org/package/2006/relationships" "Relationship"]
  set textShapeId "tomRipImageId$shapeId"
  $myPara setAttribute Id $textShapeId
  $myPara setAttribute Type "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
  $myPara setAttribute Target $filename
  $myPara setAttribute TargetMode "External"
  incr shapeId
  return $myPara
}

    

set shapeId 0
set lesImgs [list "toto.gif" "tata.gif"]

addImgToDoc "./origine/word/document.xml"
addImgToDocRel "./origine/word/_rels/document.xml.rels"
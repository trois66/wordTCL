package req tdom
console show
update
proc addImgToDoc {docFileName} {
  puts "Démarrage de l'inclusion dans le fichier document.xml"
  set xmlData [dom parse [tDOM::xmlReadFile ${docFileName}]]
  
  set ns2 {rel urn:schemas-microsoft-com:vml w http://schemas.openxmlformats.org/wordprocessingml/2006/main}
  set rootFile [$xmlData documentElement]
  set lastParagraph [$xmlData selectNodes -namespaces $ns2 {/w:document/w:body}]
  set myPara [createTheParagraphPart $xmlData]
  #$lastParagraph appendChild $myPara
  $lastParagraph insertBefore $myPara [$lastParagraph lastChild]
  
  # Write the new document image inserted
  set fdOut [open "${docFileName}" w]
  fconfigure $fdOut -encoding utf-8
  puts $fdOut {<?xml version="1.0" encoding="UTF-8" standalone="yes"?>}
  puts $fdOut [$rootFile asXML]
  close $fdOut
  
  $xmlData delete


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

proc addImgToDocRel {docFileName fileName} {
  puts "Démarrage de l'inclusion dans le fichier document.xml.rels"
  set xmlData [dom parse [tDOM::xmlReadFile $docFileName]]
  set ns {rel http://schemas.openxmlformats.org/package/2006/relationships}
  set rootFile [$xmlData documentElement]
  set lastParagraph [$xmlData selectNodes -namespaces $ns {/rel:Relationships}]
  set myPara [createTheRelationshipsPart $xmlData $fileName]
  $lastParagraph appendChild $myPara
  # Write the new document image inserted
  set fdOut [open "${docFileName}" w]
  fconfigure $fdOut -encoding utf-8
  puts $fdOut {<?xml version="1.0" encoding="UTF-8" standalone="yes"?>}
  puts $fdOut [$rootFile asXML]
  close $fdOut
  $xmlData delete
}



proc createTheRelationshipsPart {doc filename} {
  global shapeId
  set myPara [$doc createElementNS "http://schemas.openxmlformats.org/package/2006/relationships" "Relationship"]
  set textShapeId "tomRipImageId$shapeId"
  $myPara setAttribute Id $textShapeId
  $myPara setAttribute Type "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
  $myPara setAttribute Target $filename
  $myPara setAttribute TargetMode "External"
  return $myPara
}
    

set shapeId 0
set lesImgs [list "toto.gif" "titi.gif" "tata.gif"]

proc addImage {listImgPath wordDoc} {
  global shapeId
  foreach fichierAInclure $listImgPath {
    addImgToDoc "$wordDoc/word/document.xml"
    addImgToDocRel "$wordDoc/word/_rels/document.xml.rels" $fichierAInclure
    incr shapeId
  }
}

addImage $lesImgs "./origine"



puts "C'est Fini"
package req TclOO
source "./wordDocument.tcl"

set monDoc [wordDocument new document.docx]
#$monDoc initialisation
puts [$monDoc getTrImgIdLast]
$monDoc addParagraph "1 le contenu de mon paragraph sans titre"
$monDoc addParagraph "2 le contenu de mon paragraph avec titre" "Titre1"
$monDoc addParagraph "3 le paragraph" "heading 1"

$monDoc addImage "./test.gif"
$monDoc addImage "test.gif"
$monDoc addImage "test.gif"
$monDoc addParagraph "3 le paragraph" "Titre1"
$monDoc addParagraph "3 le paragraph" "Titre2"
$monDoc addParagraph "3 le paragraph" "Titre2"
$monDoc addParagraph "3 le paragraph" "Titre3"


#set rels [wordDocumentRels new]
#$rels loadRels document.docx


#$monDoc copyAndMountDocX document.docx
foreach im [$monDoc getContentList] {
puts [[$im renderTDOM] asXML]
}
puts "#################################################################"
$monDoc copyAndMountDocX document.docx
#puts [[$monDoc renderTDOM] asXML]
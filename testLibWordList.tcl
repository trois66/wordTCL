package req TclOO
source "./wordDocument.tcl"

set monDoc [wordDocument new document.docx]
#$monDoc initialisation

set laFameuseListe {}
lappend laFameuseListe {Paragraph "je ne sais pas" "Titre1"}
lappend laFameuseListe {Paragraph "Mais il faudrait savoir"}
lappend laFameuseListe {Image "test.gif"}
puts $laFameuseListe
$monDoc createFromList $laFameuseListe
#$monDoc addParagraph "1 le contenu de mon paragraph sans titre"
#$monDoc addParagraph "2 le contenu de mon paragraph avec titre" "Titre1"
#$monDoc addParagraph "3 le paragraph" "heading 1"

#$monDoc addImage "./test.gif"
#$monDoc addImage "test.gif"
#$monDoc addImage "test.gif"
#$monDoc addParagraph "3 le paragraph" "Titre1"
#$monDoc addParagraph "3 le paragraph" "Titre2"
#$monDoc addParagraph "3 le paragraph" "Titre2"
#$monDoc addParagraph "3 le paragraph" "Titre3"


#set rels [wordDocumentRels new]
#$rels loadRels document.docx

puts "#################################################################"

#$monDoc copyAndMountDocX document.docx
foreach im [$monDoc getContentList] {
puts [[$im renderTDOM] asXML]
}
puts "#################################################################"
$monDoc copyAndMountDocX document.docx
#puts [$monDoc renderTDOM] 
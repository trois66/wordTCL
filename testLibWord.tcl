#package req TclOO

lappend auto_path [pwd]

package req wordTcl


set monDoc [::wordTcl::wordDocument new document.docx]
$monDoc addParagraph "1 le contenu de mon paragraph sans titre"
$monDoc addParagraph "2 le contenu de mon paragraph avec titre" "Titre1"
$monDoc addParagraph "3 le paragraph" "heading 1"

$monDoc addImage "./test.gif"
$monDoc addImage "test.gif"
$monDoc addImage "test.gif"
$monDoc addParagraph "4 le paragraph" "Titre1"
$monDoc addParagraph "5 le paragraph" "Titre2"
$monDoc addParagraph "6 le paragraph" "Titre2"
$monDoc addParagraph "7 le paragraph" "Titre3"


# foreach im [$monDoc getContentList] {
  # puts [[$im renderTDOM] asXML]
# }
puts "#################################################################"
$monDoc copyAndMountDocX documentFromAdd.docx

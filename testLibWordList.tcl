#package req TclOO
lappend auto_path [pwd]

package req wordTcl

set monDoc [wordTcl::wordDocument new document.docx]
#$monDoc initialisation

set laFameuseListe {}
lappend laFameuseListe {Paragraph "On insere un titre" "Titre1"}
lappend laFameuseListe {Paragraph "c'est un paragraph Normal qui suit le titre au format TITRE 1"}
lappend laFameuseListe {Paragraph "deuxieme paragraph suivi d'une image"}
lappend laFameuseListe {Image "test.gif"}
lappend laFameuseListe {Paragraph "ici on met une image qui n'existe pas" "Titre1"}
lappend laFameuseListe {Paragraph "juste pour voir les liens finaux dans le docx"}
lappend laFameuseListe {Image "toto.gif"}
lappend laFameuseListe {Paragraph "Titre de niveau 1" "Titre1"}
lappend laFameuseListe {Paragraph "Titre de niveau 2" "Titre2"}
lappend laFameuseListe {Paragraph "Titre de niveau 3" "Titre3"}
lappend laFameuseListe {Paragraph "Titre de niveau 4" "Titre4"}
lappend laFameuseListe {Paragraph "Titre de niveau 5" "Titre5"}
lappend laFameuseListe {Paragraph "Titre de niveau 6" "Titre6"}

# puts $laFameuseListe
$monDoc createFromList $laFameuseListe


puts "#################################################################"
$monDoc copyAndMountDocX documentResult.docx
#puts [$monDoc renderTDOM]  
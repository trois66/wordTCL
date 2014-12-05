package req TclOO
package require tdom

oo::define wordDocument method addParagraph {content {style ""}} {
        variable documentContent
        set temp [paragraph new]
        $temp setParagraphContent $content
        $temp setStyle $style
        lappend documentContent $temp
        return #$documentContent
    }

oo::class create paragraph {
    constructor {} {
        variable paragraphContent ""
        variable paragraphStyle ""
    }
    method getParagraphContent {} {
        variable paragraphContent
        return $paragraphContent
    }
    method setParagraphContent {content} {
        variable paragraphContent
        set paragraphContent $content
    }
    method getStyle {} {
        variable paragraphStyle
        return $paragraphStyle
    }
    method setStyle {content} {
        variable paragraphStyle
        set paragraphStyle $content
    }
    method render {} {
        variable paragraphStyle
        variable paragraphContent
        set result "<w:p>"
        if {[[self] getStyle] != ""} {
            set result "$result<w:pPr><w:pStyle w:val=\"[[self] getStyle]\"/></w:pPr>"
        }
        set result "$result<w:r><w:t>[[self] getParagraphContent]</w:t></w:r></w:p>"
        return  $result
    }
    # TODO: use tdom lib to generate XML
    method renderTDOM {} {
        set doc [dom createDocument "w:p"]
        set root [$doc documentElement]
        #$root setAttribute version 1.0
        if {[[self] getStyle] != ""} {
            dom createNodeCmd elementNode w:pPr
            dom createNodeCmd elementNode w:pStyle
            dom createNodeCmd textNode t
            $root appendFromScript {
                w:pPr {
                    w:pStyle  w:val [[self] getStyle] {}
                }
            
            }
            #$root appendChild [[$doc createElement w:pPr] appendChild [$doc createElement w:pStyle]]]
        }
        dom createNodeCmd elementNode w:r
        dom createNodeCmd elementNode w:t
        dom createNodeCmd textNode t
        $root appendFromScript {
            w:r {
                w:t  {t [[self] getParagraphContent] }
            }
            
        }
        #$root appendChild [[$doc createElement w:r] appendChild [[$doc createElement w:t] setAttribute w:val [[self] getStyle]]]
        return $root
    }
    destructor {
        variable paragraphContent
        variable paragraphStyle
        puts "Destruction du paragraph"
    }
}


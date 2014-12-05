package req TclOO
package require tdom

#package require Img

#proc imageSize {file {format ""}} {
#    if { [catch {image create photo -file $file -format $format} img] } {
#       if { $format == "" } {
#            set form ""
#          } else {
#            set form "$format "
#          }
#       error "'$file' is not a recognisable ${form}image"
#    }
#    set info [list [image width $img] [image height $img]]
#    image delete $img
#    return $info;
#}


proc gifsize {name} {
    set f [open $name r]
    fconfigure $f -translation binary
    # read GIF signature -- check that this is
    # either GIF87a or GIF89a
    set sig [read $f 6]
    switch $sig {
        "GIF87a" -
        "GIF89a" {
            # do nothing
        }
        default {
            close $f
            error "$f is not a GIF file"
        }
    }
    # read "logical screen size", this is USUALLY the image size too.
    # interpreting the rest of the GIF specification is left as an exercise
    binary scan [read $f 2] s wid
    binary scan [read $f 2] s hgt
    close $f

    return [list $wid $hgt]
 }

oo::define wordDocument method addImage {content {format "external"} {style ""}} {
    variable documentContent
    variable imageId
    variable relsDocument
    set temp [image new]
    $temp setImagePath $content
    $temp setImageFormat $format
    $temp setImageStyle $style
    $temp setImageId [[self] getTrImgIdLast]
    lappend documentContent $temp
    $relsDocument setNewRelsToAdd $temp
    return #$documentContent
}

oo::class create image {
    constructor {} {
        variable imagePath ""
        variable imageFormat ""
        variable imageStyle ""
        variable imageId 0
    }
    method getImagePath {} {
        variable imagePath
        return $imagePath
    }
    method setImagePath {content} {
        variable imagePath
        set imagePath $content
    }
    method getImageId {} {
        variable imageId
        return $imageId
    }
    method setImageId {content} {
        variable imageId
        set imageId $content
    }
    method getImageFormat {} {
        variable imageFormat
        return $imageFormat
    }
    method setImageFormat {content} {
        variable imageFormat
        set imageFormat $content
    }
    method getImageStyle {} {
        variable imageStyle
        return $imageStyle
    }
    method setImageStyle {content} {
        variable imageStyle
        set temp [gifsize [[self] getImagePath]]
        #puts "width:[lindex $temp 0]px;height:[lindex $temp 1]px"
        set imageStyle "width:[lindex $temp 0]px;height:[lindex $temp 1]px"
    }
    # TODO: Import from paragraph
    method render {} {
        variable imagePath ""
        variable imageFormat ""
        variable imageStyle ""
        set result "<w:p>"
        if {[[self] getStyle] != ""} {
            set result "$result<w:pPr><w:pStyle w:val=\"[[self] getStyle]\"/></w:pPr>"
        }
        set result "$result<w:r><w:t>[[self] getParagraphContent]</w:t></w:r></w:p>"
        return  $result
    }
    
    method renderTDOM {} {
        set doc [dom createDocument "w:p"]
        set root [$doc documentElement]
        #$root setAttribute version 1.0
        dom createNodeCmd elementNode w:r
        dom createNodeCmd elementNode w:pict
        dom createNodeCmd elementNode v:shape
        dom createNodeCmd elementNode v:imagedata
        dom createNodeCmd elementNode o:lock
        # <w:p>
        #  <w:r>
        #    <w:pict>
        #      <v:shape id="_x0000_i161" style="width:525.75pt;height:162pt" o:preferrelative="f">
        #        <v:imagedata r:id="extImgId61"/>
        #        <o:lock v:ext="edit" aspectratio="f"/>
        #      </v:shape>
        #    </w:pict>
        #  </w:r>
        #</w:p>
        $root appendFromScript {
            w:r {
                w:pict  {
                    v:shape  id "_x0000_i[[self] getImageId]" style [[self] getImageStyle] o:preferrelative "f" {
                        v:imagedata r:id "trImgId[[self] getImageId]"
                        o:lock v:ext "edit" aspectratio "f"
                    }
                }
            }
            
        }
        #$root appendChild [[$doc createElement w:r] appendChild [[$doc createElement w:t] setAttribute w:val [[self] getStyle]]]
        return $root 
    }
    method renderRels {} {
    #  <Relationship Id="rId402" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="file:///D:\DemoBmRchOp\docx\gif\TSIAFOG3_TIR.gif" TargetMode="External"/>
set doc [dom createDocument "Relationships"]
set myPara [$doc createElementNS "http://schemas.openxmlformats.org/package/2006/relationships" "Relationship"]
  set textShapeId "trImgId[[self] getImageId]"
  $myPara setAttribute Id $textShapeId
  $myPara setAttribute Type "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
  $myPara setAttribute Target "file:///[[self] getImagePath]"
  $myPara setAttribute TargetMode "External"



        #set doc [dom createDocument "Relationship"]
        #set root [$doc documentElement]

        #$root appendFromScript {
       #    Id "rId402" Type "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target "file:///[[self] getImagePath]" TargetMode [[self] getImageFormat]
        #}
        #$root appendChild [[$doc createElement w:r] appendChild [[$doc createElement w:t] setAttribute w:val [[self] getStyle]]]
        return $myPara 
    }
    destructor {
        variable imagePath ""
        variable imageFormat ""
        variable imageStyle ""
        variable imageId 0
        puts "Destruction du paragraph"
    }
}


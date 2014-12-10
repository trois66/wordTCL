package require TclOO;  #if loading for the first time; not required on Tcl 8.6
package require vfs::zip
package require tdom

oo::class create wordDocumentRels {
    constructor {} {
        variable relationShipList
        variable trImgIdLast
        variable newRelsToAdd 
        set newRelsToAdd {}
        #variable docRelContent
        #variable documentXmlContent
        puts "construction de relation"
    }
    
    constructor {filename} {
        variable relationShipList
        variable trImgIdLast
        [self] loadRels $filename
        #variable docRelContent
        #variable documentXmlContent
        puts "construction de relation"
    }
    method  setNewRelsToAdd {elem} {
        variable newRelsToAdd
        lappend newRelsToAdd $elem
    
    }
    method loadRels {filename} {
        variable relationShipList
        variable trImgIdLast
        set mnt_file [vfs::zip::Mount $filename ./temp]
        set doc [dom parse [tDOM::xmlReadFile "./temp/word/_rels/document.xml.rels"]]
        
        set root [$doc documentElement]
        puts "avant: [$root asXML]"
        
        set ns {rel http://schemas.openxmlformats.org/package/2006/relationships}
        set allRels [$doc selectNodes -namespaces $ns {/rel:Relationships/*}]
        set idList {}
        foreach relation $allRels {
            set tempString  [$relation getAttribute Id]
            #puts $tempString
            if {[string match "trImgId*" $tempString]} {
                lappend idList [string replace $tempString 0 2]
            }
        }
        
        if {[llength $idList] == 0} {
            set trImgIdLast 0
        } else {
            set trImgIdLast  [lindex [lsort $idList] end]
        }

        vfs::zip::Unmount $mnt_file ./temp
        

    }
    method copyDocX {} {
        variable relationShipList
        variable trImgIdLast
        variable newRelsToAdd
        if {![info exist newRelsToAdd ]} {
        return}
        
        
        
        
        
        set doc [dom parse [tDOM::xmlReadFile "./temp2/temp/word/_rels/document.xml.rels"]]
        
        
        
        
        set ns {rel http://schemas.openxmlformats.org/package/2006/relationships}
  set rootFile [$doc documentElement]
  set lastParagraph [$doc selectNodes -namespaces $ns {/rel:Relationships}]
  foreach nrel $newRelsToAdd {
    set myPara [$nrel renderRels]
    $lastParagraph appendChild $myPara
}
        
        
        
        
        set root [$doc documentElement]
        puts "_________rels: [$root asXML]"
        
        
        
        
        puts "on ecris fichier\n\n"
        set out [open "./temp2/temp/word/_rels/document.xml.rels" w]
        chan configure $out -encoding utf-8
        puts $out "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
        puts $out [$doc asXML]
        close $out

    }
    method getTrImgIdLast {} {
        variable trImgIdLast
        return [incr trImgIdLast]
    }
    method createTheRelationshipsPart {doc id} {
        set myPara [$doc createElementNS "http://schemas.openxmlformats.org/package/2006/relationships" "Relationship"]
        set textShapeId "trImgId$id"
        $myPara setAttribute Id $textShapeId
        $myPara setAttribute Type "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
        $myPara setAttribute Target $filename
        $myPara setAttribute TargetMode "External"
        return $myPara
    }
    destructor {
         puts "destruction de document"
    }
}

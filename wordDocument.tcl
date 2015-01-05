package require TclOO;  #if loading for the first time; not required on Tcl 8.6
package require vfs::zip
package require tdom

oo::class create wordDocument {
    constructor {} {
        variable documentContent
        variable documentXmlContent
        variable relsDocument
        puts "construction de document"
    }
    constructor {filename} {
        variable documentSourceFileName
        set documentSourceFileName $filename
        variable documentContent
        variable documentXmlContent
        variable relsDocument
        set relsDocument [wordDocumentRels new $filename]
        puts "construction de document avec template '$filename'."
    }
    method initialisation {{filename document.docx}} {
        variable documentContent
        variable documentXmlContent
        variable relsDocument
        set relsDocument [wordDocumentRels new]
        $relsDocument loadRels $filename
        puts "construction de document"
    }
    method render {} {
        variable documentContent
        set result ""
        foreach element $documentContent {
            #puts [$element getParagraphContent]
            set result "[$element render] $result"
        }
        return $result
    }
    method renderTDOM {} {
        variable documentContent
        set result "<root>"
        foreach element $documentContent {
        #puts "TUTU:[$element renderTDOM]\n\n\n"
            set result "$result[$element renderTDOM]"
        }
        set result "$result</root>"
        #puts $result
        set valueToRet [dom parse $result]
        set root [$valueToRet documentElement]
        return $root
    }
    method getContentList {} {
        variable documentContent
        return $documentContent
    }
    method getTrImgIdLast {} {
        variable relsDocument
        return [$relsDocument getTrImgIdLast]
    }
    method  createFromList {laListe} {
        foreach line $laListe {
            set type [lindex $line 0]
            set content [lindex $line 1]
            set arg [lrange $line 2 end]
            #puts "traitement de la ligne:$line\n"
            #puts "type:$type\n"
            #puts "content:$content\n"
            #puts "arg:$arg\n"
            [self] add$type $content $arg
        
        }

    }
    method copyAndMountDocX {filename} {
        variable documentSourceFileName
        variable documentXmlContent
        variable documentContent
        variable relsDocument
        #file copy -force $documentSourceFileName temp_$documentSourceFileName
        set mnt_file [vfs::zip::Mount $documentSourceFileName ./temp]
        #puts "mnt_file:$mnt_file"
        file delete -force -- ./temp2/temp
        
        file copy -force -- temp "./temp2/"
        #file copy -force "temp/_rels/.rels" "./temp2/temp/_rels/rels"
        vfs::zip::Unmount $mnt_file ./temp
       
        #cd ..
        
        set doc [dom parse [tDOM::xmlReadFile "./temp2/temp/word/document.xml"]]
        
        set root [$doc documentElement]
        #puts "avant: [$root asXML]"
        
        
        set ns2 {rel urn:schemas-microsoft-com:vml w http://schemas.openxmlformats.org/wordprocessingml/2006/main}
        set body [$doc selectNodes -namespaces $ns2 {/w:document/w:body}]
        #set myPara [[self] renderTDOM]
       
        #puts "_________________________________________________"
        foreach element $documentContent {
        #puts "TUTU:[$element renderTDOM]\n\n\n"
            $body insertBefore [$element renderTDOM] [$body lastChild]
            #set result "$result[$element renderTDOM]"
        }
        #$body insertBefore $myPara [$body lastChild]
  
  
        #set body [$root selectNodes "/w:body"]
        #puts "avantbody: [$body asXML]"
        #$body appendChild [[self] renderTDOM]
        #puts [$doc asXML]
        
        #puts "on ecris fichier\n\n"
        set out [open "./temp2/temp/word/document.xml" w]
        chan configure $out -encoding utf-8
        puts $out "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
        puts $out [$doc asXML]
        close $out
        $relsDocument copyDocX
        puts "on compresse tout\n\n"
        set listFileF [findFiles temp2/temp/ *]
        # use this file for linux
        if { [lindex $::tcl_platform(os) 0] != "Windows" } {
           lappend listFileF "_rels/.rels"
        } 
        
        zipper::list2zip temp2/temp $listFileF $filename
        
    }
    destructor {
         puts "destruction de document"
    }
}



# findFiles
# basedir - the directory to start looking in
# pattern - A pattern, as defined by the glob command, that the files must match
proc findFiles { basedir pattern } {

    # Fix the directory name, this ensures the directory name is in the
    # native format for the platform and contains a final directory seperator
    
        set allPwd [pwd]
    
    set basedir [string trimright [file join [file normalize $basedir] { }]]
        cd $basedir
    #set basedir [file normalize $basedir]
    set fileList {}

    # Look in the current directory for matching files, -type {f r}
    # means ony readable normal files are looked at, -nocomplain stops
    # an error being thrown if the returned list is empty
    foreach fileName [glob -nocomplain -type {f r} $pattern] {
        lappend fileList $fileName
    }

    # Now look for any sub direcories in the current directory
    foreach dirName [glob -nocomplain -type {d  r} *] {
        # Recusively call the routine on the sub directory and append any
        # new files to the results
        set subDirList [findFiles $dirName $pattern]
        if { [llength $subDirList] > 0 } {
            foreach subDirFile $subDirList {
                lappend fileList $dirName/$subDirFile
            }
        } else {
            lappend fileList $dirName
        }
    }
        cd $allPwd
    return $fileList
 }
 

source "./zipper.tcl"
source "./wordDocumentRels.tcl"
source "./wordParagraph.tcl"
source "./wordImage.tcl"
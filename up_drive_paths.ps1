#----------------------------------------------
# Name:           Up drive path
# Version:        1.0.0.0
# Start date:     09.07.2014
# Release date:   11.07.2014
# Description:    
#
# Author:         George Dicu
# Department:     Cloud, Backup  
#----------------------------------------------

cd\

$nbpath = "C:\Program Files\Veritas\Volmgr\bin"

if (Test-Path $nbpath) {

    cd $nbpath

    $drives_info = .\vmoprcmd
    
    $drive1 = $drives_info[24..28]
    $drive2 = $drives_info[31..35]
    $drive3 = $drives_info[38..42]
    $drive4 = $drives_info[45..49]
    $drive5 = $drives_info[52..56]
    
    for($i=1;$i -le 5;$i++){
        
        $drive = Get-Variable -Name "drive$i" -ValueOnly
        $downpaths = $drive -match "DOWN-TLD"
        
        if ($downpaths.count -ne 0) {
        
            write-host "IBM.ULT3580-TD5.Drive$i has "$downpaths.Count" paths down,"
            
            foreach ($downpath in $downpaths){
             
                write-host "Puttin up "($downpath -replace "\s+",";").split(";")[1..2]""
                
                .\vmoprcmd -upbyname "IBM.ULT3580-TD5.Drive$i" -h ($downpath -replace "\s+",";").split(";")[1] 
                
                Start-Sleep -s 5
                
            }
            write-host "All paths from drive IBM.ULT3580-TD5.Drive$i are up."
        }
        else{
            write-host "IBM.ULT3580-TD5.Drive$i has 0 paths down"
        }        
    }
    
}
else {
    write-host "This script can only run on Netbackup Windows Servers"
}
#----------------------------------------------
# Name:           tape_empty_header_search
# Version:        1.0.0.0
# Start date:     21.11.2014
# Release date:   21.11.2014
# Description:    
#
# Author:         George Dicu
# Department:     Cloud, Backup  
#----------------------------------------------

cd\

$nbpath = "C:\Program Files\Veritas\Volmgr\bin"
$dc =@{}
$tapes = @{}
#$tape = Read-Host

$tapes = Import-Csv "C:\EH tapes test\tapes.csv"

function tape_exists ($tape) {
    if ($dc -match $tape) {
        return $true
    }
    else {
        return $false
    }
}

if (Test-Path $nbpath) {

    cd $nbpath
    
    foreach($tape in $tapes) {
         
        $tape = $tape | Select -ExpandProperty Tapes

        $drives_info = .\vmoprcmd
        
        $drive1 = $drives_info[24..28]
        $drive2 = $drives_info[31..35]
        $drive3 = $drives_info[38..42]
        $drive4 = $drives_info[45..49]
        $drive5 = $drives_info[52..56]
        
        $dc = $drives_info[23],$drives_info[30],$drives_info[37],$drives_info[44],$drives_info[51]

        
        if (tape_exists ($tape)) {
        
            $ft = $dc -match $tape
            Write-Host (Get-Date -format "MM.dd.yyyy HH:mm:ss") " - Tape $tape exist in " ($ft[0].Split(""))[0].split(".")[-1]
            $driveused = Get-Variable -Name ($ft[0].Split(""))[0].split(".")[-1] -ValueOnly
            $p = $driveused -match "active"
            $path = (($P[0] -replace(' ')).split("{"))[1].split("}")[0]
            
            Write-Host (Get-Date -format "MM.dd.yyyy HH:mm:ss") " - Scanning $tape, please wait..."
            .\scsi_command -map -d "{$path}" > "C:\EH tapes test\$tape" 
            
            Write-Host (Get-Date -format "MM.dd.yyyy HH:mm:ss") " - Scanning of tape $tape finnished, unmounting..."
            #'C:\Program Files\Veritas\NetBackup\bin\admincmd\nbrbutil.exe -releaseMedia $tape'
            .\tpunmount $tape -force
        }
        else {
        
            Write-Host (Get-Date -format "MM.dd.yyyy HH:mm:ss") " - Tape $tape doesnt exist in a drive, Requesting..."
            .\tpreq -m $tape -f $tape
            
            Start-Sleep -s 60
            
            #updateing tape cache
            $drives_info = .\vmoprcmd
            
            $drive1 = $drives_info[24..28]
            $drive2 = $drives_info[31..35]
            $drive3 = $drives_info[38..42]
            $drive4 = $drives_info[45..49]
            $drive5 = $drives_info[52..56]
        
            $dc = $drives_info[23],$drives_info[30],$drives_info[37],$drives_info[44],$drives_info[51]
            
            if (tape_exists ($tape)) {
                
                $ft = $dc -match $tape
                Write-Host (Get-Date -format "MM.dd.yyyy HH:mm:ss") " - Tape $tape exist in " ($ft[0].Split(""))[0].split(".")[-1]
                $driveused = Get-Variable -Name ($ft[0].Split(""))[0].split(".")[-1] -ValueOnly
                $p = $driveused -match "active"
                $path = (($P[0] -replace(' ')).split("{"))[1].split("}")[0]
                
                Write-Host (Get-Date -format "MM.dd.yyyy HH:mm:ss") " - Request completed, scanning $tape, please wait..."
                .\scsi_command -map -d "{$path}" > "C:\EH tapes test\$tape"
                
                Write-Host (Get-Date -format "MM.dd.yyyy HH:mm:ss") " - Scanning of tape $tape finnished, unmounting..."
                #'C:\Program Files\Veritas\NetBackup\bin\admincmd\nbrbutil.exe -releaseMedia $tape'
                .\tpunmount $tape -force 
            }
            
        }
        
    }
}
else {
    write-host "This script can only run on Netbackup Windows Servers, $nbpath path incorrect"
}


$objResult = @()

cd\
cd "C:\Program Files\Veritas\NetBackup\bin\admincmd\"

write-host "Check interval XO300 - XO379"
for($i=300;$i -le 379;$i++){

$rows = .\nbemmcmd.exe -listmedia -mediaid "XO0$i"

    $PropertyHash = @{}
    $PropertyHash +=  @{
        
        "Media ID" = (($rows[3] -replace "\s+",";").split(";"))[2]
        "Media Barcode" =  (($rows[10] -replace "\s+",";").split(";"))[1]
        "Status" = (($rows[29] -replace "\s+",";").split(";"))[2]
        "Host" = (($rows[12] -replace "\s+",";").split(";"))[3]
    }
    
    $objResult += New-Object -TypeName PSObject -Property $PropertyHash 
}

write-host "Check interval XO380 - XO409"
for($i=380;$i -le 409;$i++){

$rows = .\nbemmcmd.exe -listmedia -mediaid 0$i"L5"

    $PropertyHash = @{}
    $PropertyHash +=  @{
        
        "Media ID" = (($rows[3] -replace "\s+",";").split(";"))[2]
        "Media Barcode" =  (($rows[10] -replace "\s+",";").split(";"))[1]
        "Status" = (($rows[29] -replace "\s+",";").split(";"))[2]
        "Host" = (($rows[12] -replace "\s+",";").split(";"))[3]
    }
    
    $objResult += New-Object -TypeName PSObject -Property $PropertyHash 
}

write-host "Check interval XO410 - XO499"
for($i=410;$i -le 499;$i++){

$rows = .\nbemmcmd.exe -listmedia -mediaid "XO0$i"

    $PropertyHash = @{}
    $PropertyHash +=  @{
        
        "Media ID" = (($rows[3] -replace "\s+",";").split(";"))[2]
        "Media Barcode" =  (($rows[10] -replace "\s+",";").split(";"))[1]
        "Status" = (($rows[29] -replace "\s+",";").split(";"))[2]
        "Host" = (($rows[12] -replace "\s+",";").split(";"))[3]
    }
    
    $objResult += New-Object -TypeName PSObject -Property $PropertyHash 
}

write-host "Check interval IBM001 - IBM099"
for($i=1;$i -le 99;$i++){

$rows = .\nbemmcmd.exe -listmedia -mediaid @{$true="IBM00$i";$false="IBM0$i"}[$i -le 9]

    $PropertyHash = @{}
    $PropertyHash +=  @{
        
        "Media ID" = (($rows[3] -replace "\s+",";").split(";"))[2]
        "Media Barcode" =  (($rows[10] -replace "\s+",";").split(";"))[1]
        "Status" = (($rows[29] -replace "\s+",";").split(";"))[2]
        "Host" = (($rows[12] -replace "\s+",";").split(";"))[3]
    
    }
    
    $objResult += New-Object -TypeName PSObject -Property $PropertyHash 
}

$fn = ($objResult | ? {$_.Status -eq "FROZEN"}).Length

$media = $objResult | ? {$_.Status -eq "FROZEN"} | Select-Object -ExpandProperty "Media ID"
$hostcontrol = $objResult | ? {$_.Status -eq "FROZEN"} | Select-Object -ExpandProperty "Host"

if($fn -eq 0){
    write-host "zero media frozen."
    exit
}

for($i=0;$i -le ($fn-1);$i++){
    
    Write-Host ($fn-$i)" media to unfreeze"
    Write-Host "unfreezing media "$media[$i]" on host control "$hostcontrol[$i]
    
    .\bpmedia -unfreeze -m $media[$i] -h $hostcontrol[$i]
}


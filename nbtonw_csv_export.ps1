#----------------------------------------------
# Name:           export_to_csv_for_nw_migation
# Version:        1.0.0.0
# Start date:     9.01.2016
# Release date:   
# Description:    
#
# Author:         George Dicu
# Department:     Cloud, Backup  
#----------------------------------------------

cd\

$nbpath = "C:\Program Files\Veritas\NetBackup\bin\admincmd"

#$policies = Read-Host
#$policies = Import-Csv "C:\EH tapes test\tapes.csv"
$container = @()

#function to create the array of dictionaries
function hashtable () {

    $PropertyHash = @{}
    $PropertyHash +=  @{
        "policy" = $args[0]
        "policy_Type" = $args[1]
        "client" = $args[2]
        "saveset" = $args[3]
        "retention" = $args[4]
        "shedule_type" = $args[5]
        "daily_windows" = $args[6]
    }
    $container += New-Object -TypeName PSObject -Property $PropertyHash
}

if (Test-Path $nbpath) {
    
    cd $nbpath
    
    
    #hash where all data will be saved
    $PolicyHash = @{}
    $policies = @()
    
    Write-Host "Please give link to policies csv file exported by export_all_polies.ps1 script"
    #$pcsvpath = Read-Host
    $policies = Import-Csv -Path "C:\temp\allpolicies2.csv"  
    
    foreach($policy in $policies){
    
        #creating $pc array for all trimmed items from $policycontainer and the final container:$PolicyHash
        $pc = @()
        #excel index variable
        $policyname = $policy | select -ExpandProperty PolicyName
        $policycontainer = .\bppllist $policyname -U
        
        if(($policy | select -ExpandProperty Active) -eq "yes"){
        
            Write-Output "going through Policy:$policyname"
            $pc = @()
            
            #Iterating the policy items without 1st two items
            foreach($item in $policycontainer[2..($policycontainer.Count)]){

                #if on item is null eliminate it
                if ($item.Trim() -eq ""){
                    continue
                }           
                
                if($item -match "EXCLUDE DATE " -or $item -match "SPECIFIC DATE "){
                    continue
                }
                #recreating the array for new item in $policycontainer
                $pc += $item.Trim()
            }
        }
        
        #///////////////////////////////////////////////////////////////////////
        #GET CLIENTS
        #///////////////////////////////////////////////////////////////////////
        
        #once we have the policy detail without any spaces or uneccessary details we get the clients:
        $clients = $pc[([array]::IndexOf($pc,($pc -match "HW/OS/Client:")[0]))..(([array]::IndexOf($pc,($pc -match "Include:")[0]))-1)]
        
        #///////////////////////////////////////////////////////////////////////
        #GET POLICY AND POLICY DETAIL
        #///////////////////////////////////////////////////////////////////////
        
        #saveing all items that could be relevant for migration
        $matches = ("Policy Name","Policy Type")
        
        #///////////////////////////////////////////////////////////////////////
        #GET SAVESETS/SELECTIONS
        #///////////////////////////////////////////////////////////////////////
        
        #Include Column/Property has different aspects for different policy types
        #getting index interval from Include Column  to 1st Schedule-1 where resided last Include item
        $allincludes = $pc[([array]::IndexOf($pc,($pc -match "Include:")[0]))..(([array]::IndexOf($pc,($pc -match "Schedule:")[0]))-1)]
        #1st element has Include: before it`s begining, removing it
        $allincludes[0] = $allincludes[0].substring(8).trim()
        foreach ($include in $allincludes) {
            if($include -eq "NEW_STREAM"){
                    continue
            } 
            $allincludes += $include
        } 
        #///////////////////////////////////////////////////////////////////////
        #GET SCHEDULES
        #///////////////////////////////////////////////////////////////////////
        
        #Since every policy has different number of schedules and Columns/Propierties in them are named 
        #same we have to separate them in order to save them gracefully into hash table

        #because $schedules will always be an array, even if we have 1 item we will iterate the array, 
        #this way will lose a if statement after foreach statement whos iteratting the $schedules
        $schedules = $pc -match "^Schedule:"
        $schindex = @()
        foreach($schedule in $schedules){
            $schindex += [array]::IndexOf($pc,$schedule)
        }
        
        #///////////////////////////////////////////////////////////////////////
        #GET SCHEDULES
        #///////////////////////////////////////////////////////////////////////
        
    }
}
else {
    write-host "This script can only run on Netbackup Windows Servers, $nbpath path incorrect"
}
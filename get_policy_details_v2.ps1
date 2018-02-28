#----------------------------------------------
# Name:           get_policy_details
# Version:        1.0.0.0
# Start date:     20.12.2014
# Release date:   20.12.2014
# Description:    
#
# Author:         George Dicu
# Department:     Cloud, Backup  
#----------------------------------------------

cd\

$nbpath = "D:\Program Files\Veritas\NetBackup\bin\admincmd"

#$policies = Read-Host
#$policies = Import-Csv "C:\EH tapes test\tapes.csv"

if (Test-Path $nbpath) {
    
    cd $nbpath
    
    
    #hash where all data will be saved
    $PolicyHash = @{}
    $policies = @()
    
    Write-Host "Please give link to policies csv file with column named : Name"
    $pcsvpath = Read-Host
    $policies = Import-Csv -Path $pcsvpath  
    
    #creating excel file
    $Excel = new-Object -comobject Excel.Application  
    $Excel.visible = $true
    $Workbook = $Excel.Workbooks.Add()
    $pi = 1
    
    foreach($policy in $policies){
    
        #creating $pc array for all trimmed items from $policycontainer and the final container:$PolicyHash
        $pc = @()
        #excel index variable
        $exi=8
        $policyname = $policy | select -ExpandProperty PolicyName
        $policycontainer = .\bppllist $policyname -U
        
        if(($policy | select -ExpandProperty Active) -eq "no"){
            continue
        }
        
        Write-Host "going through Policy:$policyname"
        
        
        
        ($Workbook.Worksheets.Add()).Name = @{$true=$policyname;$false="PolicyNameTooLong_$pi"}[($policyname.Length) -le 31]
        $ActiveWorksheet = $Workbook.Worksheets.Item(1)
        [void] $ActiveWorksheet.Activate()
        
        $pi++
        #Iterating the policy items without 1st two items
        foreach($item in $policycontainer[2..($policycontainer.Count)]){

            #if on item is null eliminate it
            if ($item.Trim() -eq ""){
                continue
            }
            #recreating the array for new item in $policycontainer
            $pc += $item.Trim()
        }

        #creating final container(a hash table) with special Column that need special formating/manipulating
        $PolicyHash +=  @{
            
            #Column that needs special splitting/formatting after matching it
            "Effective date" =  $pc -match "^Effective date:" | % { $_.Split("")[-2].trim() }
            "Effective time" =  $pc -match "^Effective date:" | % { $_.Split("")[-1].trim() }
            "Volume Pool" = ($pc -match "^Volume Pool:")[0].Split(":")[-1].trim()
            "Server Group" = ($pc -match "^Server Group:")[0].Split(":")[-1].trim()
            "Residence" = ($pc -match "^Residence:")[0].Split(":")[-1].trim()
            "Residence is Storage Lifecycle Policy" = ($pc -match "^Residence is Storage Lifecycle Policy:")[0].Split(":")[-1].trim()

        }
        
        Write-Host "   *  going through Policy Details"
        
        # adding 1st data inside excel
        $ActiveWorksheet.Cells.Item(1,1) = "Name"
        $ActiveWorksheet.Cells.Item(1,1).Font.Bold = $true
        $ActiveWorksheet.Cells.Item(1,2) = "Value"
        $ActiveWorksheet.Cells.Item(1,2).Font.Bold = $true
        $ActiveWorksheet.Cells.Item(2,1) = "Effective date"
        $ActiveWorksheet.Cells.Item(2,2) =  $pc -match "^Effective date:" | % { $_.Split("")[-2].trim() }
        $ActiveWorksheet.Cells.Item(3,1) = "Effective time"
        $ActiveWorksheet.Cells.Item(3,2) =  $pc -match "^Effective date:" | % { $_.Split("")[-1].trim() }
        $ActiveWorksheet.Cells.Item(4,1) = "Volume Pool"
        $ActiveWorksheet.Cells.Item(4,2) = ($pc -match "^Volume Pool:")[0].Split(":")[-1].trim()
        $ActiveWorksheet.Cells.Item(5,1) = "Server Group"
        $ActiveWorksheet.Cells.Item(5,2) = ($pc -match "^Server Group:")[0].Split(":")[-1].trim()
        $ActiveWorksheet.Cells.Item(6,1) = "Residence"
        $ActiveWorksheet.Cells.Item(6,2) = ($pc -match "^Residence:")[0].Split(":")[-1].trim()
        $ActiveWorksheet.Cells.Item(7,1) = "Residence is Storage Lifecycle Policy"
        $ActiveWorksheet.Cells.Item(7,2) = ($pc -match "^Residence is Storage Lifecycle Policy:")[0].Split(":")[-1].trim()


        #new array with all columns from $pc array that are same nd unique for all policy types
        $matches = ("Policy Name","Policy Type","Active","File Restore Raw","Mult. Data Streams",
        "Client Encrypt","Checkpoint","Policy Priority","Max Jobs/Policy","Disaster Recovery",
        "Collect BMR info","Keyword","Data Classification","Application Discovery","Discovery Lifetime",
        "ASC Application and attributes","Granular Restore Info","Ignore Client Direct","Client Compress",
        "Enable Metadata Indexing","Index server name","Use Accelerator","Collect TIR info",
        "File Restore Raw","Interval","Optimized Backup","Application Consistent","Block Incremental",
        "Cross Mount Points","Follow NFS Mounts","Exchange DAG Preferred Server","Exchange Source passive db if available")

        #iterrating $Matches array and adding new hash item intoo the table with its value
        foreach ($match in $matches){

            $value = $pc -match ("^"+[regex]::escape($match)+":") | % { $_.Split(":")[-1].trim() }
            
            #save each column/policy properties with its value in hashtable
            if ($value) {
                $PolicyHash[$match] = $value
                $ActiveWorksheet.Cells.Item($exi,1) = $match
                $ActiveWorksheet.Cells.Item($exi,2) = $value
            }
            #if the policy type doesnt have a propertie(column) save its value as "-"
            else {
                $PolicyHash[$match] = "-"
                $ActiveWorksheet.Cells.Item($exi,1) = $match
                $ActiveWorksheet.Cells.Item($exi,2) = "-"
            }
            $exi++
        }
         
        #Include Column/Property has different aspects for different policy types
        #getting index interval from Include Column  to 1st Schedule-1 where resided last Include item
        $allinclude = $pc[([array]::IndexOf($pc,($pc -match "Include:")[0]))..(([array]::IndexOf($pc,($pc -match "Schedule:")[0]))-1)]
        $allincludes = @()
        $allincludes += $allinclude[0].substring(8).trim()
        foreach ($includes in $allinclude[1..($allinclude.Count)]) {
            if($includes -eq "NEW_STREAM"){
                    continue
            } 
            $allincludes += $includes
        }
        #after finding all Include items, saveing them intoo our hastable
        #1st item in Include array starting with "Include: " so we eliminate this and save it outside foreach
        $PolicyHash["Include_0"] = $allinclude[0].substring(8).trim()

        if(($allincludes.Count) -eq 1){
            $ActiveWorksheet.Cells.Item($exi,1) = "Include"
            $ActiveWorksheet.Cells.Item($exi,2) = $allincludes[0]
            $exi++
        }
        else {
            #$i=1
            $range = (($allincludes.Count) + $exi) - 1
            $ActiveWorksheet.Cells.Item($exi,1) = "Includes"
            $ActiveWorksheet.Cells.Item($exi,1).Orientation = 90
            $ActiveWorksheet.Cells.Item($exi,1).VerticalAlignment = -4108
            $ActiveWorksheet.Cells.Item($exi,1).HorizontalAlignment = -4108
            
            [void]($ActiveWorksheet.Range("A${exi}:A${range}")).Select()
            ($ActiveWorksheet.Range("A${exi}:A${range}")).MergeCells = $true 
            
            foreach($includeitem in $allincludes){
                #$PolicyHash["Include_$i"] = $includeitem
                #$i++
                $ActiveWorksheet.Cells.Item($exi,2) = $includeitem
                $exi++
            }
        }
        #saveing the policy clients in saparate array
        $allservers = $pc[([array]::IndexOf($pc,($pc -match "HW/OS/Client:")[0]))..(([array]::IndexOf($pc,($pc -match "Include:")[0]))-1)]

        #after finding all servers items, saveing them intoo our hastable
        #1st item in servers array starting with "HW/OS/Client:" so we eliminate it
        $servers = @()
        #save all items into other array except the 1st one who we add after
        #we elimintate "HW/OS/Client:" from it
        $newallservers = $allservers[1..($allservers.Count)]
        $newallservers += $allservers[0].substring(14).trim()

        $range = (($newallservers.Count) + $exi)
        $ActiveWorksheet.Cells.Item($exi,1) = "Clients"
        $ActiveWorksheet.Cells.Item($exi,1).Orientation = 90
        $ActiveWorksheet.Cells.Item($exi,1).VerticalAlignment = -4108
        $ActiveWorksheet.Cells.Item($exi,1).HorizontalAlignment = -4108

        [void]($ActiveWorksheet.Range("A${exi}:A${range}")).Select()
        ($ActiveWorksheet.Range("A${exi}:A${range}")).MergeCells = $true

        $ActiveWorksheet.Cells.Item($exi,2) = "Client Hostname"
        $ActiveWorksheet.Cells.Item($exi,2).Font.Bold = $true
        $ActiveWorksheet.Cells.Item($exi,3) = "OS"
        $ActiveWorksheet.Cells.Item($exi,3).Font.Bold = $true
        $ActiveWorksheet.Cells.Item($exi,4) = "Hardware"
        $ActiveWorksheet.Cells.Item($exi,4).Font.Bold = $true
        
        Write-Host "   *  going through Clients"
        
        $ci = 1
        foreach($server in $newallservers){
            #$servers +=  ($server -replace "\s+",";").split(";")[2]

            $PolicyHash["Hostname_$ci"] += ($server -replace "\s+",";").split(";")[2]
            $PolicyHash["Client_OS_$ci"] += ($server -replace "\s+",";").split(";")[1]
            $PolicyHash["Client_HW_$ci"] += ($server -replace "\s+",";").split(";")[0]
            $ci++

            $ActiveWorksheet.Cells.Item($exi+1,2) = ($server -replace "\s+",";").split(";")[2]
            $ActiveWorksheet.Cells.Item($exi+1,3) = ($server -replace "\s+",";").split(";")[1]
            $ActiveWorksheet.Cells.Item($exi+1,4) = ($server -replace "\s+",";").split(";")[0]
            $exi++
        }
        $exi++
        #$PolicyHash["Clients"] = $servers

        #Since every policy has different number of schedules and Columns/Propierties in them are named 
        #same we have to separate them in order to save them gracefully into hash table

        #because $schedules will always be an array, even if we have 1 item we will iterate the array, 
        #this way will lose a if statement before foreach statement whos iteratting the $schedules
        $schedules = $pc -match "^Schedule:"
        $indexno = @()
        foreach($schedule in $schedules){
            $indexno += [array]::IndexOf($pc,$schedule)
        }
        $indexno += ($pc.Count)

        $submatches = ("Schedule","Type","Calendar sched","Frequency","Synthetic",
        "Checksum Change Detection","PFI Recovery","Maximum MPX","Retention Level",
        "Number Copies","Fail on Error","Residence","Volume Pool","Server Group",
        "Residence is Storage Lifecycle Policy","Schedule indexing")

        for($i=1;$i -le ($indexno.Count-1);$i++){
             
            $subpc = @()
            #saveing schedule information based on $indexno intervals, in case its the last inval
            #ternary op is watching if the is the last interval in this case we need the exact indexno value, not -1 
            $subpc = $pc[$indexno[$i-1]..(@{$true=$indexno[$i];$false=($indexno[$i])-1}[$i -eq ($indexno.Count-1)])]
            $schedulename = $subpc[0].Split(":")[-1].trim()
            
            Write-Host "   *  going through Schedule:$schedulename"
            
            $includedates = @()
            $dailywindows = @()
            $includedates = $subpc[([array]::IndexOf($subpc,"Included Dates-----------")+1)..([array]::IndexOf($subpc,"Excluded Dates----------")-1)]
            $dailywindows = $subpc[([array]::IndexOf($subpc,"Daily Windows:")+1)..($subpc.Count)]
            
            $range = 0
            $range = (($submatches.Count) + ($includedates.Count) + ($dailywindows.Count) + $exi)
            
            
            [void]($ActiveWorksheet.Range("A${exi}:A${range}")).Select()
            ($ActiveWorksheet.Range("A${exi}:A${range}")).MergeCells = $true
            
            $ActiveWorksheet.Cells.Item($exi,1) = "Schedule $schedulename"
            $ActiveWorksheet.Cells.Item($exi,1).Orientation = 90
            $ActiveWorksheet.Cells.Item($exi,1).VerticalAlignment = -4108
            $ActiveWorksheet.Cells.Item($exi,1).HorizontalAlignment = -4108
            
            $ActiveWorksheet.Cells.Item($exi,2) = "Name"
            $ActiveWorksheet.Cells.Item($exi,2).Font.Bold = $true
            $ActiveWorksheet.Cells.Item($exi,3) = "Value"
            $ActiveWorksheet.Cells.Item($exi,3).Font.Bold = $true
            $exi++
            
            #iterrating $submatches array and adding new hash item intoo the table with its value
            foreach ($submatch in $submatches){
            
                $subvalue = $subpc -match ("^"+[regex]::escape($submatch)+":") | ForEach-Object { $_.Split(":")[-1].trim() }
                
                #save each column/policy properties with its value in hashtable
                if ($subvalue) {
                    $PolicyHash["Schedule_$i $submatch"] = $subvalue
                    $ActiveWorksheet.Cells.Item($exi,2) = $submatch
                    $ActiveWorksheet.Cells.Item($exi,3) = $subvalue
                }
                #if the policy type doesnt have a propertie(column) save its value as "-"
                else {
                    $PolicyHash["Schedule_$i $submatch"] = "-"
                    $ActiveWorksheet.Cells.Item($exi,2) = $submatch
                    $ActiveWorksheet.Cells.Item($exi,3) = "-"
                }
                $exi++
            }
            
            $ActiveWorksheet.Cells.Item($exi,2) = "Include Dates"
            $ActiveWorksheet.Cells.Item($exi,2).Font.Bold = $true
            
            $range2 = $range - ($dailywindows.Count)
            
            [void]($ActiveWorksheet.Range("B${exi}:B${range2}")).Select()
            ($ActiveWorksheet.Range("B${exi}:B${range2}")).MergeCells = $true
            
            ($ActiveWorksheet.Range("B${exi}:B${range2}")).VerticalAlignment = -4108
            ($ActiveWorksheet.Range("B${exi}:B${range2}")).HorizontalAlignment = -4108
            
            foreach ($includedate in $includedates){
                
                $ActiveWorksheet.Cells.Item($exi,3) = $includedate
                $exi++
            }
            
            if($dailywindows){
                $ActiveWorksheet.Cells.Item($exi,2) = "Daily Windows"
                $ActiveWorksheet.Cells.Item($exi,2).Font.Bold = $true
                
                [void]($ActiveWorksheet.Range("B${exi}:B${range}")).Select()
                ($ActiveWorksheet.Range("B${exi}:B${range}")).MergeCells = $true
                
                $ActiveWorksheet.Range("B${exi}:B${range}").VerticalAlignment = -4108
                $ActiveWorksheet.Range("B${exi}:B${range}").HorizontalAlignment = -4108
                
                foreach ($dailywindow in $dailywindows){
                    $ActiveWorksheet.Cells.Item($exi,3) = $dailywindow
                    $exi++
                }
            }
            $PolicyHash["Include Dates of $schedulename"] += $subpc[([array]::IndexOf($subpc,"Included Dates-----------")+1)..([array]::IndexOf($subpc,"Excluded Dates----------")-1)]
            $PolicyHash["Daily Windows of $schedulename"] += $subpc[([array]::IndexOf($subpc,"Daily Windows:")+1)..($subpc.Count)]
            #arranging every sheet after filling with data
            $objRange = $ActiveWorksheet.UsedRange 
            [void] $objRange.EntireColumn.Autofit()
            $objRange.Borders.LineStyle = 1
        }
    }
}
else {
    write-host "This script can only run on Netbackup Windows Servers, $nbpath path incorrect"
}

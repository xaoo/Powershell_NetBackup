#----------------------------------------------
# Name:           export_NB_cfg_to_xml
# Version:        2.0.0.
# Start date:     24.11.2017
# Release date:   24.11.2017
# Description:    
#
# Author:         George Dicu
# Department:     Cloud, Backup  
#----------------------------------------------

cd\

$nbpath = "C:\Program Files\Veritas\NetBackup\bin\admincmd\"
$container = @()

# Assign the CSV and XML Output File Paths
$XML_Path = "C:\Users\DT997753a\Desktop\SED-cfg.xml"

#///////////////////////////////////////////////////////////////////////
# Create the XML Object, File Tags
#///////////////////////////////////////////////////////////////////////
$xmlWriter = New-Object System.XMl.XmlTextWriter($XML_Path,$Null)
$xmlWriter.Formatting = 'Indented'
$xmlWriter.Indentation = 1
$XmlWriter.IndentChar = "`t"
$xmlWriter.WriteStartDocument()
$xmlWriter.WriteComment('Get all Information about a specific NetBackup domain')
$xmlWriter.WriteStartElement('Policies')
$xmlWriter.WriteEndElement()
$xmlWriter.WriteEndDocument()
$xmlWriter.Flush()
$xmlWriter.Close()

#///////////////////////////////////////////////////////////////////////
# BEGIN
#///////////////////////////////////////////////////////////////////////
if (Test-Path $nbpath) {
    
    cd $nbpath
    
    #Policies Container
    $PolicyHas = @{}
    $policies = @()
    
    Write-Host "Please give link to policies csv file exported by export_all_polies.ps1 script"
    #$pcsvpath = Read-Host
    $policies = Import-Csv -Path "C:\Users\dt997753a\Desktop\all-policies.csv" 
    
    $ii=0
    foreach($policy in $policies){
        
        #creating $pc array for all trimmed items from $policycontainer and the final container:$PolicyHash
        $pc = @()
        $policyname = $policy | select -ExpandProperty PolicyName
        $policycontainer = .\bppllist $policyname -U
        
        #XML Create Policy Node
        $xmlDoc = [System.Xml.XmlDocument](Get-Content $XML_Path);
        $xmlpolicy = $xmlDoc.CreateElement("Policy")
        [void]$xmlDoc.SelectSingleNode("//Policies").AppendChild($xmlpolicy)
        $xmlpolicy.SetAttribute("Name", $policyname)
        $xmlDoc.Save($XML_Path)
        
        if(($policy | select -ExpandProperty Active) -eq "no"){
            continue            
        }

        Write-Output "going through Policy:$policyname"($policies.count-$ii)
        $pc = @()
        $ii++
            
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

        #///////////////////////////////////////////////////////////////////////
        #GET CLIENTS
        #///////////////////////////////////////////////////////////////////////
        #SQL can have Clients field defined as none, exluding
        $clnt = $pc -match "Clients:"
        
        if($clnt){
            $clnt = ($clnt -Split(":"))[-1].trim()
            
            $xmlclients = $xmlpolicy.AppendChild($xmlDoc.CreateElement("Clients"));
            $txtnode = $xmlclients.AppendChild($xmlDoc.CreateTextNode("None Defined"));
            $xmlDoc.Save($XML_Path)
        }
        else{          
            
            #once we have the policy detail without any spaces or uneccessary details we get the clients:
            $clients = $pc[([array]::IndexOf($pc,($pc -match "HW/OS/Client:")[0]))..(([array]::IndexOf($pc,($pc -match "Include:")[0]))-1)]
            $clients[0] = $clients[0].substring(14).trim()
            $hostnames = @()
            $architecture = @()
            $os = @()
            
            #XML Create Clients Node
            $xmlclients = $xmlpolicy.AppendChild($xmlDoc.CreateElement("Clients"));
            
            foreach($client in $clients){
                $architecture = ($client -replace "\s+",";").split(";")[0]
                $hostname =  ($client -replace "\s+",";").split(";")[2]
                $os = ($client -replace "\s+",";").split(";")[1]
                
                #XML Create Clients Hostname
                $xmlclient = $xmlclients.AppendChild($xmlDoc.CreateElement("Client"));
                    $xmlclient.SetAttribute("HostName", $hostname)
                        $xmlos = $xmlclient.AppendChild($xmlDoc.CreateElement("OS"));
                            $txtnode = $xmlos.AppendChild($xmlDoc.CreateTextNode($os));
                        $xmlarh = $xmlclient.AppendChild($xmlDoc.CreateElement("Architecture"));
                            $txtnode = $xmlarh.AppendChild($xmlDoc.CreateTextNode($architecture));
                 $xmlDoc.Save($XML_Path)
            }
        }
        #///////////////////////////////////////////////////////////////////////
        #GET POLICY AND POLICY DETAIL
        #///////////////////////////////////////////////////////////////////////
        
        #saving all items that could be relevant for migration
        $matches = ("Active","Effective date","Block Incremental","Mult. Data Streams","Client Encrypt","Checkpoint","Policy Priority",
"Max Jobs/Policy","Disaster Recovery","Collect BMR info","Residence","Volume Pool","Server Group","Keyword","Data Classification","Residence is Storage Lifecycle Policy",
"Application Discovery","Discovery Lifetime","ASC Application and attributes","Granular Restore Info","Ignore Client Direct","Use Accelerator",
"Client List Type","Selection List Type","Oracle Backup Data File Name Format","Oracle Backup Archived Redo Log File Name Format",
"Oracle Backup Control File Name Format","Oracle Backup Fast Recovery Area File Name Format","Oracle Backup Set ID",
"Oracle Backup Data File Arguments","Oracle Backup Archived Redo Log Arguments","Database Backup Share Arguments","Client Compress",
"Follow NFS Mounts","Cross Mount Points","Collect TIR info","Application Defined","Backup network drvs","Optimized Backup",
"Catalog Disaster Recovery Configuration","Email Address","Disk Path","User Name","Pass Word","Critical policy","File Restore Raw")
        
        #policyname and type can be obtain from $policy variable
        $pn = $policy | select -ExpandProperty PolicyName
        $pt = $policy | select -ExpandProperty PolicyType
        
        #XML Create Options Node
            $xmloptions = $xmlpolicy.AppendChild($xmlDoc.CreateElement("Options"));
                $xmlpt = $xmloptions.AppendChild($xmlDoc.CreateElement("Policy_Type"));
                    $xmlnode = $xmlpt.AppendChild($xmlDoc.CreateTextNode($pt));
             $xmlDoc.Save($XML_Path)
             
        #iterrating $Matches array and adding new hash item intoo the table with its value
        foreach ($match in $matches){
    
            $value = $pc -match ("^"+[regex]::escape($match)+":") | ForEach-Object { $_.Split(":")[-1].trim() }
            
            #xml cannot save name with space, so every match need to be renamed
            if(($match -like '*/*')){
                $name_match = $match.replace("/","_")
                $name_match = $name_match.replace(" ","_")                
            }
            else {
                $name_match = $match.replace(" ","_")
            }
            
            #save each column/policy properties with its value in hashtable
            if ($value) {
            
                #XML Create Option Node
                $xmlopt = $xmloptions.AppendChild($xmlDoc.CreateElement($name_match));
                $xmlnode = $xmlopt.AppendChild($xmlDoc.CreateTextNode($value));
                $xmlDoc.Save($XML_Path)
            }
            #if the policy type doesnt have a propertie(column) save its value as "-"
            else {
                #XML Create Option Node
                $xmlopt = $xmloptions.AppendChild($xmlDoc.CreateElement($name_match));
                $xmlnode = $xmlopt.AppendChild($xmlDoc.CreateTextNode("-"));
                $xmlDoc.Save($XML_Path)
            }
        }
        
        #///////////////////////////////////////////////////////////////////////
        #GET SAVESETS/SELECTIONS
        #///////////////////////////////////////////////////////////////////////
        
        #Include Column/Property has different aspects for different policy types
        #getting index interval from Include Column  to 1st Schedule-1 where resided last Include item
        $allincludes = $pc[([array]::IndexOf($pc,($pc -match "Include:")[0]))..(([array]::IndexOf($pc,($pc -match "Schedule:")[0]))-1)]
        
        #1st element has Include: before it`s begining, removing it
        $allincludes[0] = $allincludes[0].substring(8).trim()
        
        #XML Create Option Node
        $xmlincludes = $xmlpolicy.AppendChild($xmlDoc.CreateElement("BackupSelections"));
         
        foreach ($include in $allincludes) {
        
            if($include -eq "NEW_STREAM"){
                    continue
            } 
            $allincludes += $include
            
            $xmlincl = $xmlincludes.AppendChild($xmlDoc.CreateElement("Selection"));
            $xmlnode = $xmlincl.AppendChild($xmlDoc.CreateTextNode($include));
            $xmlDoc.Save($XML_Path)
        }        
        
        #///////////////////////////////////////////////////////////////////////
        #GET ALL SCHEDULES
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
        #add the last element in the $pc array
        $schindex += $pc.count
        
        #Get all common column of all schedules
        $sched_matchs = ("Calendar sched","Checksum Change Detection","Daily Windows","Fail on Error","Maximum MPX","Number Copies","PFI Recovery",
                        "Residence","Residence is Storage Lifecycle Policy","Retention Level","Server Group","Synthetic","Type",
                        "Volume Pool")
        
        # Create Schedules Node
        $xmlscheds = $xmlpolicy.AppendChild($xmlDoc.CreateElement("Schedules"));
                  
        for($i=1;$i -le ($schindex.Count-1);$i++){
    
            $subpc = @()
            $subpc = $pc[$schindex[$i-1]..(@{$true=$schindex[$i];$false=($schindex[$i])-1}[$i -eq ($schindex.Count-1)])]
            $schedname = $subpc[0].Split(":")[-1].trim()
            $schedtype = $subpc[1].Split(":")[-1].trim()
            
            Write-Output "    Going through: $schedname"
            
            $xmlsched = $xmlscheds.AppendChild($xmlDoc.CreateElement("Schedule"));
            $xmlsched.SetAttribute("Name", $schedname)
            $xmlDoc.Save($XML_Path)
             
            #iterrating $submatches array and adding new hash item intoo the table with its value
            foreach ($submatch in $sched_matchs){
            
                $subvalue = $subpc -match ("^"+[regex]::escape($submatch)+":") | ForEach-Object { $_.Split(":")[-1].trim() }
                
                #xml cannot save name with space, so every match need to be renamed
                $name_submatch = $submatch.replace(" ","_")
                
                #save each column/policy properties with its value in hashtable
                if ($subvalue) {
                
                    $xmlsubmatch = $xmlsched.AppendChild($xmlDoc.CreateElement($name_submatch));
                    $xmlnode = $xmlsubmatch.AppendChild($xmlDoc.CreateTextNode($subvalue));
                    $xmlDoc.Save($XML_Path)
                }
                #if the policy type doesnt have a propertie(column) save its value as "-"
                else {
                
                    $xmlsubmatch = $xmlsched.AppendChild($xmlDoc.CreateElement($name_submatch));
                    $xmlnode = $xmlsubmatch.AppendChild($xmlDoc.CreateTextNode("-"));
                    $xmlDoc.Save($XML_Path)
                }
            }            
            
            $dailyWindow = $subpc[([array]::IndexOf($subpc,"Daily Windows:")+1)..($subpc.Count)]
            $xmldaily = $xmlsched.AppendChild($xmlDoc.CreateElement("DailyWindows"));
            
            foreach($daily in $dailyWindow){
                $xmlnode = $xmldaily.AppendChild($xmlDoc.CreateTextNode($daily));
                $xmlDoc.Save($XML_Path)
            }
            #Application Backup does not have Include Dates
            if($schedtype -ne "Application Backup"){
                $includedates = $subpc[([array]::IndexOf($subpc,"Included Dates-----------")+1)..([array]::IndexOf($subpc,"Excluded Dates----------")-1)]
                $xmlincl = $xmlsched.AppendChild($xmlDoc.CreateElement("IncludeDates"));
            }
            #Default-Application-Backup schedule type does not have Include option
            if($schedname -ne "Default-Application-Backup"){
                foreach($incl in $includedates){
                    $xmlnode = $xmlincl.AppendChild($xmlDoc.CreateTextNode($incl));
                    $xmlDoc.Save($XML_Path)
                }
            }
        }
    }
}
else {
    write-host "This script can only run on Netbackup Windows Servers, $nbpath path incorrect"
} 

$xmlDoc.Save($XML_Path)
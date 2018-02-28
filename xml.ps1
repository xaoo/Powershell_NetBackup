# Assign the CSV and XML Output File Paths
     $XML_Path = "C:\Users\a560341\Desktop\Sample.xml"
      
     # Create the XML File Tags
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
      
      
     # Create Policy Node
     $xmlDoc = [System.Xml.XmlDocument](Get-Content $XML_Path);
     $policy = $xmlDoc.CreateElement("Policy")
     $xmlDoc.SelectSingleNode("//Policies").AppendChild($policy)
     $policy.SetAttribute("Name", "Test_Policy")
     $xmlDoc.Save($XML_Path)
      
     $options = $policy.AppendChild($xmlDoc.CreateElement("Options"));
        $option = $options.AppendChild($xmlDoc.CreateElement("RootFolder"));
            $RootFolderTextNode = $option.AppendChild($xmlDoc.CreateTextNode("Root folder Title"));
     
     $selections = $policy.AppendChild($xmlDoc.CreateElement("Selections"));
        $selection = $selections.AppendChild($xmlDoc.CreateElement("Selection"));
            $RootFolderTextNode = $selection.AppendChild($xmlDoc.CreateTextNode("Root folder Title"));

     $schedules = $policy.AppendChild($xmlDoc.CreateElement("Schedules"));
         $schedule = $schedules.AppendChild($xmlDoc.CreateElement("Schedule"));
         $schedule.SetAttribute("Name", "Schedule_1")
         $schedule = $schedules.AppendChild($xmlDoc.CreateElement("Schedule"));
         $schedule.SetAttribute("Name", "Schedule_2")

     $clients = $policy.AppendChild($xmlDoc.CreateElement("Clients"));
        $client = $clients.AppendChild($xmlDoc.CreateElement("HostName"));
            $RootFolderTextNode = $client.AppendChild($xmlDoc.CreateTextNode("adm_srv_1"));
     
     $xmlDoc.Save($XML_Path)
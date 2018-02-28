$exp = @() 
$exp = get-content C:\temp\exp 
$i = 1 
$c =$exp.Count  

foreach($item in $exp){ 
    Write-host "$i. Expiring $item, "($c-$i)" to expire." 
    bpexpdate.exe -d 0 -force -m $item 
    $i++ 
}
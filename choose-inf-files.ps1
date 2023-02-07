# Must be located in the .\Driver folder
$allInfFiles = Get-ChildItem -Path "Driver" *.inf 
$i = 0
$allInfFiles | ForEach-Object {
    $i++
    $inf = $_.Name
    Write-Host("[$i] $inf")
}
[int]$user_input = Read-Host("Select an inf-file: ")
$user_input--

$selected_inf = $allInfFiles[$user_input]

Write-Host $selected_inf


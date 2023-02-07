<#
Execute this script in the same as the Driver-Folder
it searches for all the .inf files 
and matches the HP_Mombi_Driver_Name
which is needed when installing an HP Printer
#>
$files = Get-ChildItem -Path "Driver" *.inf
$pattern = '^HP_Mombi_Driver_Name="(.*)"$'
$regex = [regex]::new($pattern)

$files | ForEach-Object {
  $m = Get-Content $_ | Select-String $regex -AllMatches | Select-Object -First 1
  if (-not [string]::IsNullOrEmpty($m))
  {
    $driverName = $m.Matches.Groups[1].Value
  }
}


Write-Output $driverName

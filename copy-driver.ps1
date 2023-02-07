Param(
  [Parameter(Mandatory=$true)][String] $target_host
)

# Do not change
$driver_path = (Resolve-Path .\).Path
$driver_folder_name = Split-Path -Path (Get-Location) -Leaf
$target_folder = "\\$target_host\c$\$driver_folder_name"

# Copies from current folder and creates a new one on the target
Robocopy.exe $driver_path $target_folder  /E



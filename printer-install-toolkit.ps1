#------------------------------------------------------
# Variable Declaration
#------------------------------------------------------
<#
Static declaration
#>
$global:infFile = $null
$base = (Get-Item .).FullName 
$driver = $base + "\Driver"

#------------------------------------------------------
# Functions
#------------------------------------------------------
# Action 1
function printer_info()
{
  get-printer | Format-Table Name, DriverName, PortName
}
# Action 2
function select_inf_file()
{
  Write-Host "Select one inf-file."
  # Must be located in the .\Driver folder
  $allInfFiles = Get-ChildItem -Path "Driver" *.inf 
  $i = 0
  $allInfFiles | ForEach-Object {
    $i++
    $file = $_.Name
    Write-Host("[$i] $file")
  }
  [int]$user_input = Read-Host("Your choice")
  $user_input--

  $selected_item = $allInfFiles[$user_input]
  $out = $selected_item.Name
  Write-Host "You selected '$out'."
  $global:infFile = $selected_item # ignore this error
}
# Action 3
function install_printer()
{
  $inf  = $global:infFile
  $install_printer = $true
  # Importig variables
  $config = Get-ChildItem -Filter config.ps1
  if($config)
  {
    $x = $config.Name
    Write-Host "Import variables from $x ..."
    . ./printer-config.ps1
  } else
  {
    Write-Host "Please create printer-config.ps1 with:"
    Write-Host "
    `$hostname_printer = ""   # hostname_printer == Get-PrinterPort.Name
    `$shown_printer_name = "" # shown to the user on apps
    `$driver_name = ""        # only visible after install, install manually at first
    "
  }

  # checking if varaibles exist
  if([string]::IsNullOrEmpty($hostname_printer))
  {
    $install_printer = $false
    Write-Host "Please set `$hostname_printer"
  }
  if([string]::IsNullOrEmpty($shown_printer_name))
  {
    $install_printer = $false
    Write-Host "Please set `$shown_printer_name"

  }
  if([string]::IsNullOrEmpty($driver_name))
  {
    $install_printer = $false
    Write-Host "Please set `$driver_name."
  }
  if([string]::IsNullOrEmpty($inf))
  {
    $install_printer = $false
    Write-Host "Please select an inf file..."
  }

  Write-Host ""
  
  $allPrinters = Get-Printer
  # if any printer is already installed
  if(($allPrinters) -and ($install_printer -eq $true) )
  {
    ForEach ($printer in $allPrinters)
    {
      if($printer.PortName -eq $hostname_printer)
      {
        $n = $printer.Name
        $pn = $printer.PortName
        Write-Host("The Printer '$n' with '$pn' as port, is already installed.")
        [string]$choice = Read-Host("Remove the printer (plus port and driver) and reinstall it? [y/n/q]")
        switch ($choice)
        {
          "y"
          {
            # delete the driver to reinstalling it
            remove_printer $n $pn $driver_name
            $install_printer = $true
          }
          "n"
          {
            # continue the installation
            $install_printer = $true
            Write-Host "Ok, continuing with the installation as usual..."
            continue
          }
          Default
          {
            Set-Location($base)
            $install_printer = $false
            exit
          }
        } # END SWITCH
      }# END IF
    }# END FOREACH
  }# END IF

  if ($install_printer -eq $true)
  {
    Set-Location $driver
    Write-Host "pnputil.exe is adding $inf..."
    # you can find the inf-file inside the driver-folder
    pnputil.exe -a $inf
    Write-Host "Adding printer driver..."
    Add-PrinterDriver -Name $driver_name
    # if the port already exist, remove it with
    # Remoe-Printer
    # Remove-PrinterPort
    # Remove-PrinterDriver
    # restart the spooler service as needed
    Write-Host "Adding printer port..."
    Add-PrinterPort -Name $hostname_printer -PrinterHostAddress $hostname_printer                    
    Write-Host "Adding printer itself..."
    Add-Printer -Name $shown_printer_name -DriverName $driver_name -PortName $hostname_printer
    Get-Printer | Format-Table Name, DriverName, PortName
  } else
  {
    Write-Host "Installation couldn't be started..."
  }
  Set-Location $base
    
} 
# END printer install

# Action 4
function remove_printer ($printer_name, $printer_port, $driver_name)
{
  Write-Host "For this funtion to work, you need elevated (admin) rights!"

  if(([string]::IsNullOrEmpty($printer_name)) -or ([string]::IsNullOrEmpty($printer_port)) -or ([string]::IsNullOrEmpty($driver_name)))
  {
    $object = Get-Printer
    $i = 0
    $object | ForEach-Object {
      $i++
      $a = $_.Name
      $b = $_.PortName
      $c = $_.DriverName
      Write-Host("[$i] Name: $a Port: $b Driver: $c")
    }
    [int]$user_input = Read-Host("Your choice")
    $user_input--

    $choice = $object[$user_input]
    $printer_name = $choice.Name
    $printer_port = $choice.PortName
    $driver_name = $choice.DriverName
  }

  Write-Host "The Script will delete Printer: $printer_name with Port: $printer_port using Driver: $driver_name ."
  [string]$succed = Read-Host("Continue?[y/n/q]")

  if($succed -eq "y")
  {
    Remove-Printer -Name $printer_name
    # Restart "Druckerwarteschlange"
    $PrintSpooler = Get-Service -Name Spooler
    if ($PrintSpooler.Status -eq 'running')
    {
      # ONLY POSSIBLE WITH elevated rights!
      Stop-Service -InputObject $PrintSpooler
    }
    Start-Service -Name Spooler
    # Get-Service -Name Spooler
    Remove-PrinterPort -Name $printer_port
    Remove-PrinterDriver -Name $driver_name

  } else
  {
    Write-Host "Deletion aborted, returning to menu..."
  }

}

# Action 5
function remove_printer_port()
{
  $object = Get-PrinterPort
  Write-Host $choice
  $i = 0
  $object | ForEach-Object {
    $i++
    $curr = $_.Name
    Write-Host("[$i] $curr")
  }
  [int]$user_input = Read-Host("Your choice")
  if($user_input -eq 0 )
  {
    return
  }
  $user_input--
  $selected_item = $object[$user_input]
  # Write-Host "You selected '$selected_item'."
  # Restart "Druckerwarteschlange"
  $PrintSpooler = Get-Service -Name Spooler
  if ($PrintSpooler.Status -eq 'running')
  {
    # ONLY POSSIBLE WITH elevated rights!
    Stop-Service -InputObject $PrintSpooler
  }
  Start-Service -Name Spooler
  # Removing the printer port
  #
  Remove-PrinterPort -Name $selected_item.Name

}

# Action 6
function remove_printer_driver()
{
  $object = Get-PrinterDriver
  $i = 0
  $object | ForEach-Object {
    $i++
    $curr = $_.Name
    Write-Host("[$i] $curr")
  }
  [int]$user_input = Read-Host("Your choice")
  if($user_input -eq 0 )
  {
    return
  }
  $user_input--
  $selected_item = $object[$user_input]
  # Restart "Druckerwarteschlange"
  $PrintSpooler = Get-Service -Name Spooler
  if ($PrintSpooler.Status -eq 'running')
  {
    # ONLY POSSIBLE WITH elevated rights!
    Stop-Service -InputObject $PrintSpooler
  }
  Start-Service -Name Spooler
  # Removing the driver
  Remove-PrinterDriver -Name $selected_item.Name

}

# Action 7
function copy_to_remote()
{
  [string]$target_host = Read-Host("To which computer do you want to copy?")
  
  if(Test-Connection -TargetName $target_host -Quiet)
  {
    # Do not change
    $driver_path = (Resolve-Path .\).Path
    $driver_folder_name = Split-Path -Path (Get-Location) -Leaf
    $target_folder = "\\$target_host\c$\$driver_folder_name"

    # Copies from current folder and creates a new one on the target
    Robocopy.exe $driver_path $target_folder  /E
  } else
  {
    Write-Host "Faulty connection to $target_host ..."
  }
}

function choose_item($object)
{
  $i = 0
  $object | ForEach-Object {
    $i++
    $curr = $_
    Write-Host("[$i] $curr")
  }
  [int]$user_input = Read-Host("Your choice")
  $user_input--

  $selected_item = $object[$user_input]
  Write-Host "You selected '$selected_item'."
  return $selected_item

}

function user_menu()
{
  Write-Host "------------------------------------------------------------------------------"
  Write-Host "What do you want to do?"
  $actions = @(
    "Quit this script."                                     # 0
    "Show all printers.",                                   # 1
    "Select an inf-file from .\Driver .",                   # 2
    "Install Printer locally with set variables",           # 3
    "Remove installed printer (plus port and driver)."      # 4
    "Remove single printer port.",                          # 5
    "Remove single printer driver.",                        # 6
    "Copy this skript to remote location."                  # 7
  )
  $i = 0
  $actions | ForEach-Object {
    Write-Host("[$i] $_")
    $i++
  }
  [int]$choice = Read-Host("Your choice: ") # Please insert integer
  Write-Host "------------------------------------------------------------------------------"

  # Numbers set in this statement related to index in $action
  switch($choice)
  {
    0 
    {
      Write-Host "Quitting..."
      exit
    }
    1
    {
      Write-Host "Currently installed printers..."
      printer_info
      user_menu
    }
    2
    {
      Write-Host "Selecting an inf from .\Driver..."
      select_inf_file 
      user_menu
    }
    3
    {
      Write-Host "Starting printer installation..."
      install_printer
      user_menu
    }
    4
    {
      Write-Host "Removing installed printer..."
      remove_printer
      user_menu
    }
    5
    {
      Write-Host "Removing installed printer port..."
      remove_printer_port
      user_menu
    }
    6
    {
      Write-Host "Removing installed printer driver..."
      remove_printer_driver
      user_menu
    }
    7
    {
      Write-Host "Copying installation folder to remote machine..."
      copy_to_remote
      user_menu
    }
    Default
    {
      Write-Host "Quitting..."
      exit
    }
  }  # END Switch
  
}

function verify_content()
{
  $faults = 0
  # Is there a printer-config.ps1 with all variables
  $config = Test-Path -Path printer-config.ps1 -Pathtype Leaf
  # Is there a Driver directory?
  $driver = Test-Path -Path ".\Driver" -PathType Container

  # verifying...
  switch($true)
  {
      ($config -eq $false)
    {
      Write-Host "Please create printer-config.ps1 with:"
      Write-Host "
        `$hostname_printer = ""   # hostname_printer == Get-PrinterPort.Name
        `$shown_printer_name = "" # shown to the user on apps
        `$driver_name = ""        # only visible after install, install manually at first
        "
      Write-Host " "
      $faults ++
    }
      ($driver -eq $false)
    {
      Write-Host "Your .\Driver directory is missing."
      Write-Host "Without drivers, no printer can be installed."
      Write-Host " "
      $faults ++
    }
  }

  if($faults -gt 0)
  {
    Write-Host "Dependency check failed, exiting..."
    exit
  }

}

<#
Starting point of the script
MAIN
#>
Clear-Host
Write-Host "RUN THIS SCRIPT WITH ELEVATED RIGHTS!"

verify_content
user_menu



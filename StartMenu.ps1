# Specify source and destination of shortcut folders below.
# ****************************************************************************
# * Note: In the source location, create two folders named 32-Bit and 64-Bit *
# *       to host your Start Menu shortcuts (as shortcut paths may vary).    *
# ****************************************************************************
$Source = "\\ABC-SVR-01\Applications$\StartMenu"
$Destination="C:\ProgramData\FolderName\Windows\Start Menu"

# Inform user if no Source or Destination has been provided.
If (($Source -eq "") -or ($Destination -eq ""))
{
    Write-Host "Missing Source and/or Destination, please check and try again."
    Exit
}

# If there is a trailing backslash for $Source and/or $Destination, remove the backslash.
If ($Source.Substring($Source.Length-1) -eq "\")
{
    $Source = $Source.Substring(0,$Source.Length-1)
}
If ($Destination.Substring($Destination.Length-1) -eq "\")
{
    $Destination = $Destination.Substring(0,$Destination.Length-1)
}

# Update the Source path to include the OS Architecture type to the end of the path.
# If no OS Architecture is returned, then use the 32-Bit shortcut folder.
$Bit = ((Get-WMIObject Win32_OperatingSystem).OSArchitecture)
If ($Bit -eq "") {$Bit = "32-Bit"}
$Source = $Source + "\" + $Bit

# Inform user if Source does not exist.
If (!(Test-Path($Source)))
{
    Write-Host "The Source location does not exist, please check and try again."
    Exit
}

# Get the latest last modified date on all shortcuts in Source and Destination paths.
$NetworkStamp = Get-ChildItem "$Source\*" -Recurse -Force -include @("*.lnk","*.xml","*.appref-ms")`
                | Where {!$_.PsIsContainer} | select Name,DirctoryName, LastWriteTime `
                | Sort LastWriteTime -descending | select -first 1
$LocalStamp = Get-ChildItem "$Destination\*" -Recurse -Force -include @("*.lnk","*.xml","*.appref-ms") -ErrorAction SilentlyContinue `
              | Where {!$_.PsIsContainer} | select Name,DirctoryName, LastWriteTime `
              | Sort LastWriteTime -descending | select -first 1

# If there are any differences in the last modified date:
# 1) Delete files in Destination (if it exists).
# 2) Copy files from Source to Destination.
If ($NetworkStamp.LastWriteTime -ne $LocalStamp.LastWriteTime)
{
    If (Test-Path($Destination))
    {
        Remove-Item "$Destination" -Recurse -Force
    }
    New-Item -ItemType directory -Path "$Destination"
    Copy-Item "$Source\*" -Destination "$Destination" -Recurse
}

# Create a list of folders copied to the destination folder.
$ListOfFolders = Get-ChildItem $Destination -Recurse -Force | Where-Object {$_.PSIsContainer} `
                 | Sort-Object FullName -Descending | % { $_.FullName }

# Create a shell object, so we can check if the shortcuts are valid.
$sh = New-Object -COM WScript.Shell

# Check if Array is Null (Prevents ForEach null bug in PowerShell 2).
If ($ListOfFolders -ne $null)
{
    
    # Do the following for every folder we scanned.
    FOREACH ($Folder in $ListOfFolders)
    {
        
        # Set the hidden flag to True for the folder.
        # Note: This will hide the folder later on if applications do not exist.
        $AllItemsHidden = $true
        
        # Get a list of all shortcuts in the current folder.
        $ListOfShortcuts = Get-ChildItem "$Folder\*" -Force -include *.lnk `
                           | % { $_.FullName }
        
        # Check if Array is Null (Prevents ForEach null bug in PowerShell 2).
        If ($ListOfShortcuts -ne $null)
        {
            
            # Do the following for every shortcut we scanned.
            FOREACH ($Shortcut in $ListOfShortcuts)
            {
                
                # Find out if the shortcut is hidden.
                $Item = Get-Item $Shortcut -Force
                $CheckHidden = (Get-ItemProperty $Item).Attributes.ToString() -match "Hidden"
                
                # Get the target path of the shortcut (the path of the application).
                # If it refers to C drive (C:), check if it exists, and ensure the shortcut
                # is visible to the users, otherwise hide the shortcut.
                $PathOfApplication = $sh.CreateShortcut($Shortcut).TargetPath
                If ($PathOfApplication.Length -ge 2)
                {
                    If ($PathOfApplication.Substring(0,2) -eq "C:")
                    {
                        If (Test-Path($PathOfApplication))
                        {
                            $AllItemsHidden = $false
                            If ($CheckHidden)
                            {
                                $Item.Attributes = $Item.Attributes -bxor [system.IO.FileAttributes]::Hidden
                            }
                        }
                        ElseIf ((!(Test-Path($PathOfApplication))) -and (!($CheckHidden)))
                        {
                            $Item.Attributes = $Item.Attributes -bxor [system.IO.FileAttributes]::Hidden
                        }
                    }
                }
                If (($Item.Directory.Name -eq "Programs") -or ($Item.DirectoryName.ToLower().Contains("\Programs\".ToLower())))
                {
                        $CheckHidden = (Get-ItemProperty $Item).Attributes.ToString() -match "Hidden"
                        If (!($CheckHidden))
                        {
                            $Item.Attributes = $Item.Attributes -bxor [system.IO.FileAttributes]::Hidden
                        }
                }
            }
    
            # List all files in the folder (excluding desktop.ini and thumbs.db, which aren't required).
            $ListOfSubFilesAndFolders = Get-ChildItem "$Folder\*" -Force -exclude "desktop.ini","thumbs.db" `
                                        | % { $_.FullName }
            
            # Check if Array is Null (Prevents ForEach null bug in PowerShell 2).
            If ($ListOfSubFilesAndFolders -ne $null)
            {
                
                # Do the following for every file or subfolder we scanned.
                FOREACH ($SubFilesOrFolder in $ListOfSubFilesAndFolders)
                {
                    
                    # Find out if the file is hidden. If it is not, then it is a required resource.
                    # inform the program not to hide the folder.
                    $Item = Get-Item $SubFilesOrFolder -Force
                    $CheckHidden = $Item.Attributes.ToString() -match "Hidden"
                    If ($CheckHidden -eq $false)
                    {
                        $AllItemsHidden = $false
                    }
                }
            }

            # If all of the files are hidden in the folder, then hide the unrequired folder.
            # If any file is unhidden in the folder, then unhide the required folder.
            $Item = Get-Item $Folder -Force
            $CheckHidden = (Get-ItemProperty $Item).Attributes.ToString() -match "Hidden"
            If (($AllItemsHidden -and !($CheckHidden)) -or (!($AllItemsHidden) -and $CheckHidden))
            {
                $Item.Attributes = $Item.Attributes -bxor [system.IO.FileAttributes]::Hidden
            }
        }

        If (($Item.Directory.Name -ne $null) -or ($Item.DirectoryName -ne $null))
        {
            If (($Item.Directory.Name -eq "Programs") -or ($Item.DirectoryName.ToLower().Contains("\Programs\".ToLower())))
            {
                    $CheckHidden = (Get-ItemProperty $Item).Attributes.ToString() -match "Hidden"
                    If (!($CheckHidden))
                    {
                        $Item.Attributes = $Item.Attributes -bxor [system.IO.FileAttributes]::Hidden
                    }
            }
        }
    }
}
# ##############################################################################
# This PowerShell script creates a top-level folder in the Windows Explorer 
# navigation panel pointing to a specified location, similar to how OneDrive
# and Dropbox do it.
# You will need to generate a unique GUID. If you provide no GUID, the script
# will generate one. You need to note the GUID in case you want to alter  or 
# delete the registry entries.
# This script needs to be run with admin rights.
#
# Usage:
# > .\topLevelIcon.ps1 -Add -Clsid <GUID> -Name <Name> -Path <Path> 
#   [-IconPath <iconPath>] [-SortOrder <SortOrder>] [-InfoTip <InfoTip>]
#
# > .\topLevelIcon.ps1 -Remove -Clsid <GUID> 
#
# Created by Thomas Novotny
# Last updated 2024-12-30
#
# ##############################################################################

param(
    # Flag to add or remove
    [switch]$Add,
    [switch]$Remove,

    # Required arguments
    [string]$Clsid,
    [string]$Name,
    [string]$Path,
    [string]$IconPath,

    # Optional arguments
    [int]$SortOrder,
    [string]$InfoTip
)

# Display usage if neither Add nor Remove is passed
if ((-not $Add -and -not $Remove) -or ($Add -and $Remove)) {
    Write-Host "Usage: .\topLevelIcon.ps1 -Add -Clsid <GUID> -Name <Name> -Path <Path> [-IconPath <iconPath>] [-SortOrder <SortOrder>] [-InfoTip <InfoTip>]"
    Write-Host "       .\topLevelIcon.ps1 -Remove -Clsid <GUID>"
    exit 1
}

# Check if user has admin rights
if (-not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Error "This script must be run as an administrator. Please restart the script with elevated privileges."
    exit 1
}

function Add-Entry {
    if (-not $Clsid) {
        $Clsid = "{$([guid]::NewGuid().ToString().ToUpper())}"
        Write-Host "Generated GUID: $Clsid"
    }
    if ($Clsid.Length -ne 38) {
        Write-Error "Clsid must be 38 characters, with curly braces at start and end"
        exit 1
    }
    if (-not $Name) {
        Write-Error "Missing -Name argument"
        exit 1
    }
    if (-not $Path) {
        Write-Error "Missing -Path argument"
        exit 1
    }
    # Check if the path exists
    if (-not (Test-Path -Path $Path)) {
        Write-Error "The path '$Path' does not exist. Exiting the script."
        exit 1
    }
    $ClsidPath = "HKCU:\Software\Classes\CLSID\$Clsid"
    $IconPathExpanded = [System.Environment]::ExpandEnvironmentVariables($iconPath)

    # Check if the registry key exists
    if (Test-Path $ClsidPath) {
        Write-Error "The registry key '$ClsidPath' already exists. Remove entry first"
        exit 1
    }

    # Check if icon exists
    if ($IconPathExpanded) {
        # Check if the path exists
        if (-not (Test-Path -Path "$IconPathExpanded")) {
            Write-Error "The icon '$iconPath' does not exist. Exiting the script."
            exit 1
        }
    }

    Write-Output "Creating Keys"

    # Add the CLSID key
    Write-Output "$ClsidPath `"$Name`""
    New-Item -Path "$ClsidPath" -Force | Out-Null
    New-ItemProperty -Path "$ClsidPath" -Name "(Default)" -Value "$Name" | Out-Null

    # Set the folder path
    Write-Output "$ClsidPath\FolderPath `"$Path`"" 
    New-ItemProperty -Path "$ClsidPath" -Name "FolderPath" -PropertyType String -Value "$Path" | Out-Null

    # Set tooltip
    Write-Output "$ClsidPath\InfoTip `"$infoTip`"" 
    if ($InfoTip) {
        New-ItemProperty -Path "$ClsidPath" -Name "InfoTip" -PropertyType String -Value "$infoTip" | Out-Null
    }

    # Set sort order index
    if ($SortOrder) {
        Write-Output "$ClsidPath\SortOrderIndex $sortOrder" 
        New-ItemProperty -Path "$ClsidPath" -Name "SortOrderIndex" -PropertyType Dword -Value $sortOrder | Out-Null
    }

    # Indicate that this is a custom created item
    Write-Output "$ClsidPath\IsCustom 1" 
    New-ItemProperty -Path "$ClsidPath" -Name "IsCustom" -PropertyType DWord -Value 1 | Out-Null

    # Set Folder to be pinned to navigation pane
    Write-Output "$ClsidPath\System.IsPinnedToNamespaceTree 1" 
    New-ItemProperty -Path "$ClsidPath" -Name "System.IsPinnedToNamespaceTree" -PropertyType DWord -Value 1 | Out-Null

    # Set to instance of Folder Shortcut
    New-Item -Path "$ClsidPath\Instance" -Force | Out-Null
    Write-Output "$ClsidPath\Instance\CLSID {0AFACED1-E828-11D1-9187-B532F1E9575D}" 
    New-ItemProperty -Path "$ClsidPath\Instance" -Name "CLSID" -PropertyType String -Value "{0AFACED1-E828-11D1-9187-B532F1E9575D}" | Out-Null

    # Add the InitPropertyBag key and set the Attributes value
    New-Item -Path "$ClsidPath\Instance\InitPropertyBag" -Force | Out-Null
    Write-Output "$ClsidPath\Instance\InitPropertyBag\Attributes 17" 
    New-ItemProperty -Path "$ClsidPath\Instance\InitPropertyBag" -Name "Attributes" -PropertyType DWord -Value 17 | Out-Null
    Write-Output "$ClsidPath\Instance\InitPropertyBag\Target `"$Path`"" 
    New-ItemProperty -Path "$ClsidPath\Instance\InitPropertyBag" -Name "Target" -PropertyType ExpandString -Value "$Path" | Out-Null

    # Set the custom icon
    if ($IconPath) {
        New-Item -Path "$ClsidPath\DefaultIcon" -Force | Out-Null
        Write-Output "$ClsidPath\DefaultIcon `"$iconPath`"" 
        New-ItemProperty -Path "$ClsidPath\DefaultIcon" -Name "(Default)" -PropertyType ExpandString -Value "$iconPath" | Out-Null
    }

    # Link to shell32.dll (required for navigation pane integration)
    New-Item -Path "$ClsidPath\InProcServer32" -Force | Out-Null
    Write-Output "$ClsidPath\InProcServer32 `"C:\Windows\System32\shell32.dll`"" 
    New-ItemProperty -Path "$ClsidPath\InProcServer32" -Name "(Default)" -PropertyType ExpandString -Value "C:\Windows\System32\shell32.dll" | Out-Null
    Write-Output "$ClsidPath\InProcServer32\ThreadingModel `"Both`"" 
    New-ItemProperty -Path "$ClsidPath\InProcServer32" -Name "ThreadingModel" -PropertyType String -Value "Both" | Out-Null

    # ShellFolder Attributes
    New-Item -Path "$ClsidPath\ShellFolder" -Force | Out-Null
    Write-Output "$ClsidPath\ShellFolder\Attributes 4034921293" 
    New-ItemProperty -Path "$ClsidPath\ShellFolder" -Name "Attributes" -PropertyType DWord -Value 4034921293 | Out-Null
    Write-Output "$ClsidPath\ShellFolder\FolderValueFlags 40" 
    New-ItemProperty -Path "$ClsidPath\ShellFolder" -Name "FolderValueFlags" -PropertyType DWord -Value 40 | Out-Null

    # Add to the Navigation Pane (NameSpace key)
    Write-Output "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Desktop\NameSpace\$Clsid `"$Name`"" 
    New-Item -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Desktop\NameSpace\$Clsid" -Force | Out-Null
    New-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Desktop\NameSpace\$Clsid" -Name "(Default)" -Value "$Name" | Out-Null

    # Hide from Desktop
    $HideDesktopIconsPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\HideDesktopIcons\NewStartPanel"
    if (-not (Test-Path -Path $HideDesktopIconsPath)) {
        New-Item -Path $HideDesktopIconsPath -Force | Out-Null
    }
    Write-Output "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\HideDesktopIcons\NewStartPanel\$Clsid 1"
    New-ItemProperty -Path $HideDesktopIconsPath -Name $Clsid -PropertyType DWord -Value 1 | Out-Null
}

function Remove-Entry {
    if (-not $Clsid) {
        Write-Error "-Clsid is a required argument"
        exit 1;
    }
    if (-not $Clsid.Length -eq 38) {
        Write-Error "Clsid must be 38 characters, with curly braces at start and end"
        exit 1
    }
    $ClsidPath = "HKCU:\Software\Classes\CLSID\$Clsid"
    Write-Host "Deleting keys"
    if (Test-Path -Path $ClsidPath) {
        Write-Host "$clsidPath"
        Remove-Item -Path $ClsidPath -Recurse
        Write-Host "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Desktop\NameSpace\$Clsid"
        Remove-Item -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Desktop\NameSpace\$Clsid" -Recurse
        Write-Output "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\HideDesktopIcons\NewStartPanel\$Clsid"
        Remove-Item -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\HideDesktopIcons\NewStartPanel\$Clsid"
    }
    else {
        Write-Host "The registry key '$ClsidPath' does not exist."
        exit 1
    }
}

if ($Add) {
    Add-Entry
    Write-Host "Shortcut `"$Name`" created, restart Explorer to see changes" -ForegroundColor Green
}
elseif ($Remove) {
    Remove-Entry
    Write-Host "Shortcut deleted, restart Explorer to see changes" -ForegroundColor Green
}
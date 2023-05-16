# gpo-import.ps1
# Author: Moritz Reufsteck
# Date: May 16, 2023
# Description: This script imports a selected GPO from C:\temp.

if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Warning "You need to run this script as an Administrator"
}

if (-not (Get-Module -ListAvailable -Name GroupPolicy)) {
    Write-Warning "GroupPolicy module is not installed"
}

Import-Module GroupPolicy

if (-not (Get-Module -Name GroupPolicy)) {
    Write-Warning "GroupPolicy module is not loaded"
}

$gpos = Get-ChildItem -Path C:\temp -Directory

$table = $gpos | ForEach-Object -Begin {$i=1} -Process {
    [PSCustomObject]@{
        Number = $i++
        Name = $_.Name
    }
}
$table | Format-Table -AutoSize

$selection = Read-Host -Prompt "Select a folder by number"

while (-not [int]::TryParse($selection, [ref]$null) -or [int]$selection -lt 1 -or [int]$selection -gt $gpos.Count) {
    Write-Warning "Invalid input. Please enter a valid folder number."
    $selection = Read-Host -Prompt "Select a folder by number"
}

$selectedFolder = $gpos[[int]$selection - 1]
# Write-Host "Selected folder: $($selectedFolder.Name)"

# Check if a valid GUID folder exists inside the selected folder
$guidFolder = Get-ChildItem -Path $selectedFolder.FullName -Directory | Where-Object { $_.Name -match '^{(.+)}$' }

if ($guidFolder) {
    $guid = $guidFolder.Name -replace '^{(.+)}$', '$1'
    # Write-Host "The GUID in the selected folder is: $guid"
} else {
    Write-Warning "No valid GUID folder found in the selected folder."
    # type any key to exit
    Read-Host -Prompt "Press any key to exit"
    exit 1
}

$import = Read-Host -Prompt "Import GPO from C:\temp\$($selectedFolder.Name)? (y/n)"
$importPath = "C:\temp\$($selectedFolder.Name)"
while ($import -ne "y" -and $import -ne "n") {
    Write-Warning "Invalid input"
    $import = Read-Host -Prompt "Import GPO from C:\temp\$($selectedFolder.Name)? (y/n)"
}
if ($import -eq "n") {
    # press any key to exit
    Read-Host -Prompt "Press any key to exit"
    exit 1
}

# Write-Host "Importing GPO $($selectedFolder.Name)..."
# Write-Host "Select the existing GPO to import the backup into"
$gpos = Get-GPO -All | Select-Object -Property Id, DisplayName
$gpos = $gpos | ForEach-Object -Begin { $i = 1 } -Process { $_ | Add-Member -MemberType NoteProperty -Name Select -Value $i -PassThru; $i++ }
$gpos | Format-Table -AutoSize
$gpo = Read-Host -Prompt "Select the existing GPO to import the backup into by number"

$gpo = [int]$gpo
while (-not ($gpo -as [int]) -or [int]$gpo -gt $gpos.Count -or [int]$gpo -lt 1) {
    Write-Warning "Invalid input"
    $gpo = Read-Host -Prompt "Select GPO by number"
}

$gpo = $gpos[$gpo - 1]
Write-Warning "##############################################"
Write-Warning "Selected GPO to import $($selectedFolder.Name) into : $($gpo.DisplayName)"
Write-Warning "This will eventually overwrite existing settings."
Write-Warning "##############################################"

# ask for final confirmation

$final_confirm = Read-Host -Prompt "Are you sure you want to import the backup into $($gpo.DisplayName)? (y/n)"
while ($final_confirm -ne "y" -and $final_confirm -ne "n") {
    Write-Warning "Invalid input"
    $final_confirm = Read-Host -Prompt "Are you sure you want to import the backup into $($gpo.DisplayName)? (y/n)"
    if ($final_confirm -eq "n") {
        # press any key to exit
        Read-Host -Prompt "Press any key to exit"
        exit 1
    }
}

Import-GPO -BackupId $guid -TargetName $gpo.DisplayName  -Path $importPath

Write-Warning "##############################################"
Write-Warning "Making dummy change to force GPO to be applied"
Write-Warning "##############################################"
Set-GPRegistryValue -Name $gpo.DisplayName -key "HKLM\Software\Dummy" -ValueName "DummyValue" -Type String -Value "DummyValue"
Remove-GPRegistryValue -Name $gpo.DisplayName -key "HKLM\Software\Dummy" -ValueName "DummyValue"
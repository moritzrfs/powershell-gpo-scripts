# gp-export.ps1
# Author: Moritz Reufsteck
# Date: May 16, 2023
# Description: This script exports a selected GPO to C:\temp.

if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Warning "You need to run this script as an Administrator"
    exit 1
}

if (-not (Get-Module -ListAvailable -Name GroupPolicy)) {
    Write-Warning "GroupPolicy module is not installed"
    exit 1
}

Import-Module GroupPolicy

if (-not (Get-Module -Name GroupPolicy)) {
    Write-Warning "GroupPolicy module is not loaded"
    exit 1
}

$gpos = Get-GPO -All | Select-Object -Property Id, DisplayName
$gpos = $gpos | ForEach-Object -Begin { $i = 1 } -Process { $_ | Add-Member -MemberType NoteProperty -Name Select -Value $i -PassThru; $i++ }
$gpos | Format-Table -AutoSize
$gpo = Read-Host -Prompt "Select GPO by number"

$gpo = [int]$gpo
while (-not ($gpo -as [int]) -or [int]$gpo -gt $gpos.Count -or [int]$gpo -lt 1) {
    Write-Warning "Invalid input"
    $gpo = Read-Host -Prompt "Select GPO by number"
}

$gpo = $gpos[$gpo - 1]
Write-Host "Selected GPO: $($gpo.DisplayName)"

$export = Read-Host -Prompt "Export GPO to C:\temp? (y/n)"
# while export not y or n
while ($export -ne "y" -and $export -ne "n") {
    Write-Warning "Invalid input"
    $export = Read-Host -Prompt "Export GPO to C:\temp? (y/n)"
    if ($export -eq "n") {
        exit 1
    }
}

if (-not (Test-Path -Path C:\temp\$($gpo.DisplayName))) {
    New-Item -Path C:\temp\$($gpo.DisplayName) -ItemType Directory
}
# backup the gpo by searching for its id
Backup-GPO -Name $gpo.DisplayName -Path C:\temp\$($gpo.DisplayName)
$exportFilePath = "C:\temp\$($gpo.DisplayName)\export.txt"
$exportContent = @"
GPO Name: $($gpo.DisplayName)
GPO ID: $($gpo.Id)
Creation Date: $(Get-Date)
-----------------------------
Exported by gpo-export.ps1
"@

$exportContent | Set-Content -Path $exportFilePath

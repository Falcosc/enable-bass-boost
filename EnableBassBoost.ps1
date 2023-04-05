<#
.SYNOPSIS
    automatically adds and enables Bass Boost to any playback device
.DESCRIPTION
    Imports registry keys to add enhancement features and enables bass boost.
    It restarts audio service to apply imported registry settings.
    Bass Boost works by reducing the gain of higher frequencies to avoid clipping which means you have to increase volume.
    If you want to restore your physical Audio Device properties, you can find your old values in %temp%\BassBoostBackup.reg
.LINK
    https://github.com/Falcosc/enable-bass-boost
    https://github.com/Falcosc/enable-loudness-equalisation
.LINK
.PARAMETER playbackDeviceName
    Searches for Audio Device Names starting with this String
.PARAMETER maxDeviceCount
    Limits the amount of devices to be configured
.PARAMETER frequency
    Bass Boost frequency
.PARAMETER decibel
    Amount of volume reduction applied to frequencies above Bass Boost frequency
.PARAMETER defaultStereoAudioLevel
    overwrites audio volume level if bass boost settings were missing
.PARAMETER defaultStereoAudioLevel
    overwrites audio volume level if bass boost settings were missing
.PARAMETER noFakeProperties
    apply fake speaker properties only temporarly to get bass boost activated and remove them after service restart.
    some drivers reset your configuration if fake properties are detected, this switch avoids this behaivor.
    If you use this switch you don't see bass boost in UI and each device change will reload settings which means you have to apply it after each reload
.EXAMPLE
    PS> .\EnableBassBoost.ps1 -playbackDeviceName BE279
    only enable bass boost with default settings
.EXAMPLE
    PS> .\EnableBassBoost.ps1 -playbackDeviceName BE279 -enableLoudness -defaultStereoAudioLevel 60
    help with quiet audio and apply 60% volume level
#>

Param(
   [Parameter(Mandatory,HelpMessage='Which Playback Device Name should be configured?')]
   [ValidateLength(3,50)]
   [string]$playbackDeviceName,
   
   [ValidateRange(1, 10)]
   [int]$maxDeviceCount=2,
   
   [ValidateRange(50, 600)]
   [int]$frequency=150,
   
   [ValidateSet(3,6,9,12,15,18,21,24)]
   [int]$decibel=12,
   
   [ValidateSet(20,40,60,80)]
   [int]$defaultStereoAudioLevel=0,
   
   [switch]$noFakeProperties=$false,
   
   [switch]$enableLoudness=$false
)

Add-Type -AssemblyName System.Windows.Forms
function exitWithErrorMsg ([String] $msg){
    [void][System.Windows.Forms.MessageBox]::Show($msg, $PSCommandPath,
    [System.Windows.Forms.MessageBoxButtons]::OK,
    [System.Windows.Forms.MessageBoxIcon]::Error)
    Write-Error $msg
    exit 1
}
function importReg ([String] $file){
    $startprocessParams = @{
        FilePath     = "$Env:SystemRoot\REGEDIT.exe"
        ArgumentList = '/s', $file
        Verb         = 'RunAs'
        PassThru     = $true
        Wait         = $true
    }
    $proc = Start-Process @startprocessParams
    If($? -eq $false -or $proc.ExitCode -ne 0) {
        exitWithErrorMsg "Failed to import $file"
    }
}

$ErrorActionPreference = "Stop"
$PSDefaultParameterValues['Out-File:Encoding'] = 'utf8'
$regFile = "$env:temp\BassBoostTMP.reg"
$regFileBackup = "$env:temp\BassBoostBackup.reg"
$bassBoostFlagKey = "{1864a4e0-efc1-45e6-a675-5786cbf3b9f0},4"
$freqKey = "{61e8acb9-f04f-4f40-a65f-8f49fab3ba10},4"
$dbKey = "{ae7f0b2a-96fc-493a-9247-a019f1f701e1},3"
$freqBytesString = [System.BitConverter]::ToString([System.BitConverter]::GetBytes($frequency), 0, 2) -replace '-',','
$dbSelection = ([int]($decibel/3)-1)
$dbSelectionString = $dbSelection.ToString().PadLeft(2,'0')
$fxPropertiesImport = @"
"{d04e05a6-594b-4fb6-a80d-01af5eed7d1d},1"="{62dc1a93-ae24-464c-a43e-452f824c4250}" ;PreMixEffectClsid activates effects
"{d04e05a6-594b-4fb6-a80d-01af5eed7d1d},2"="{637c490d-eee3-4c0a-973f-371958802da2}" ;PostMixEffectClsid activates effects
"{d04e05a6-594b-4fb6-a80d-01af5eed7d1d},3"="{5860E1C5-F95C-4a7a-8EC8-8AEF24F379A1}" ;UserInterfaceClsid shows it in ui
"{1864a4e0-efc1-45e6-a675-5786cbf3b9f0},4"=hex:03,00,00,00,01,00,00,00,02,00,00,00 ;enable bass boost
"{61e8acb9-f04f-4f40-a65f-8f49fab3ba10},4"=hex:03,00,00,00,01,00,00,00,$freqBytesString,00,00 ;frequenz 
"{ae7f0b2a-96fc-493a-9247-a019f1f701e1},3"=hex:03,00,00,00,01,00,00,00,$dbSelectionString,00,00,00 ;db 3,6,9,12,15,18,21,24
"@
$speakerPropertyKey = "{1da5d803-d492-4edd-8c23-e0c0ffee7f0e}"
$propertiesImport = @'
"{1da5d803-d492-4edd-8c23-e0c0ffee7f0e},0"=dword:00000001 ;FORM_FACTOR is needed if you form factor does not support bass boost
"{1da5d803-d492-4edd-8c23-e0c0ffee7f0e},6"=dword:00000000 ;FULL_RANGE_SPEAKERS is needed if the option was set to 3
'@

if($enableLoudness) {
    $fxPropertiesImport += "`r`n" + '"{fc52a749-4be9-4510-896e-966ba6525980},3"=hex:0b,00,00,00,01,00,00,00,ff,ff,00,00'
}

if($noFakeProperties -and $defaultStereoAudioLevel -ne 0){
    exitWithErrorMsg "You should not combine -noFakeProperties with -defaultStereoAudioLevel,
because without fake properties it would overwrite your volume setting on each execution instead of only if boost was missing"
}

switch ( $defaultStereoAudioLevel ) {
    20 { $propertiesImport += "`r`n" + '"{9855c4cd-df8c-449c-a181-8191b68bd06c},0"=hex:41,00,00,00,01,00,00,00,E0,D5,C2,C1,E0,D5,C2,C1' }
    40 { $propertiesImport += "`r`n" + '"{9855c4cd-df8c-449c-a181-8191b68bd06c},0"=hex:41,00,00,00,01,00,00,00,14,30,5E,C1,14,30,5E,C1' }
    60 { $propertiesImport += "`r`n" + '"{9855c4cd-df8c-449c-a181-8191b68bd06c},0"=hex:41,00,00,00,01,00,00,00,7D,E2,F7,C0,7D,E2,F7,C0' }
    80 { $propertiesImport += "`r`n" + '"{9855c4cd-df8c-449c-a181-8191b68bd06c},0"=hex:41,00,00,00,01,00,00,00,8C,A3,58,C0,8C,A3,58,C0' }
}

$devices = reg query HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\MMDevices\Audio\Render /s /f $playbackDeviceName /d
if(!$?) {
    exitWithErrorMsg "Could not find any device named $playbackDeviceName"
}

$renderer = $devices | Select-String "Render"
$activeRenderer = $renderer | ? { (Get-ItemPropertyValue -Path Registry::$($_ -replace '\\Properties','') -Name DeviceState) -eq 1}
if($activeRenderer.length -lt 1) {
    exitWithErrorMsg "There are $($renderer.length) devices with Name $playbackDeviceName, but non of them is active"
}
if($activeRenderer.length -gt $maxDeviceCount) {
    $devices
    exitWithErrorMsg "Execution aborted, because more then $maxDeviceCount Active Devices found by Name $playbackDeviceName"
}

$missingBassBoost = $false
"Windows Registry Editor Version 5.00" | Set-Content $regFile, $regFileBackup
$renderer | ForEach-Object{
    $propKeyPath = $_ 
    $fxKeyPath = $_ -replace 'Properties','FxProperties'
    $properties = Get-ItemProperty -Path Registry::$propKeyPath
    if (($properties.($speakerPropertyKey+",0") -ne 1) -or ($properties.($speakerPropertyKey+",6") -ne 0)){
        "[" + $propKeyPath + "]" | Add-Content $regFile, $regFileBackup
        $propertiesImport >> $regFile
        "`"$speakerPropertyKey,0`"=dword:" + $properties.($speakerPropertyKey+",0") >> $regFileBackup
        "`"$speakerPropertyKey,6`"=dword:" + $properties.($speakerPropertyKey+",6") >> $regFileBackup
        $missingBassBoost = $true
    }
    $fxProperties = Get-ItemProperty -Path Registry::$fxKeyPath
    if (($fxProperties -eq $null) -or 
    ($fxProperties.$bassBoostFlagKey -eq $null) -or ($fxProperties.$bassBoostFlagKey[8] -ne 2) -or 
    ($fxProperties.$dbKey -eq $null) -or ($fxProperties.$dbKey[8] -ne $dbSelection) -or
    ($fxProperties.$freqKey -eq $null) -or ([bitconverter]::ToInt16($fxProperties.$freqKey,8) -ne $frequency)) { 
        "[" + $fxKeyPath + "]" >> $regFile
        $fxPropertiesImport >> $regFile
        $missingBassBoost = $true
    }
}

if (!$missingBassBoost) {
    "Bass Boost Settings don't need to be applied"
    Start-Sleep -Seconds 5
    exit 0
}

$currentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
if (-not $currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    $arguments = "-File `"$($myInvocation.MyCommand.Definition)`""
    foreach($key in $MyInvocation.BoundParameters.keys) {
        if(($MyInvocation.BoundParameters[$key].GetType()) -eq [switch]) {
            if($MyInvocation.BoundParameters[$key]) {
                $arguments += " -$key"
            }
        } else {
            $arguments += " -$key " + $MyInvocation.BoundParameters[$key]
        }
    }
    Start-Process powershell -Verb runAs -ArgumentList $arguments
    exit
}
"import bass boost activation into registry"
importReg $regFile

"Restart Audio to apply registry settings"
Restart-Service audiosrv -Force

if($noFakeProperties){
    "revert fake device properties"
    importReg $regFileBackup
}


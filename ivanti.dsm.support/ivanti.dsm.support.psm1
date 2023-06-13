<#
.SYNOPSIS
    Collecting DSM Log files
.DESCRIPTION
    this script collects all DSM Logs, which are not older than the specified date, 
    and creates either a ZIP File or a Folder at the users desktop. 
.NOTES
    Additional Notes, eg
    File Name  : Export-DSMLogs.ps1
    Author     : Markus Zierer, markus.zierer@zierer-consulting.de
    Appears in -full 
.EXAMPLE
    Get-DSMLogs
    Creates a ZIP File at the User Desktop with DSM Logs of the local
    machine from within the last day

    Get-DSMLogs -Computername <Servername> -DaysBack 7
    
    Connects to the DSMLogs$ Share of the specified computer and downloads
    Logfiles from within the last 7 Days
.COMPONENT
#>

#region Main
function Get-DSMLogs {
    param(
        [string]$TargetPath = ("$env:UserProfile","Desktop" -join "\"),
        [string]$Computername = $env:COMPUTERNAME,
        [string]$DaysBack = "1",
        [switch]$Local,
        [string]$Path = "C:\Program Files (x86)\Common Files\enteo\NiLogs",
        [switch]$Debug
    )
    # Process debug mode settings
    If ($Debug){
        $DebugPreference = "Continue"
        Write-Debug -Message "Debug Mode enabled"
    }else {$DebugPreference = "SilentlyContinue"}

    # Variables
    $TempPath = "$env:LOCALAPPDATA\Temp\PrepareDSMLogs"
    $NiLogKey = 'HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\NetSupport\NetInstall\LogFileSettings'
    $NiLogValue = 'ResolvedLogFilePath'

    Write-Progress -Activity "Preparing DSM Log Files" -Status "Reading DSM Info from Registry"
    if ($Computername -eq $env:COMPUTERNAME) {
        Read-DSMInfo
    }else {
        $NiLogDir = Join-Path -Path "$Computername" -ChildPath "dsmlogs$"
        $NiLogDir = "\\" + $NiLogDir
    }
    Write-Progress -Activity "Preparing DSM Log Files" -Status "Copy files to temporary location"
    Copy-DSMLogsToTempLocation
    Write-Progress -Activity "Preparing DSM Log Files" -Status "Create DSM Log archive"
    New-DSMLogArchive
    Write-Progress -Activity "Preparing DSM Log Files" -Status "Cleanup temporary files" -PercentComplete 90
    Remove-TempFiles
    Read-Host "ZIP File created, press ENTER to continue"
}

function Remove-DSMLogs {
    param(
        [string]$Computername = $env:COMPUTERNAME,
        [string]$DaysBack = "1",
        [switch]$Debug
    )
    # Process debug mode settings
    If ($Debug){
        $DebugPreference = "Continue"
        Write-Debug -Message "Debug Mode enabled"
    }else {$DebugPreference = "SilentlyContinue"}

    # Variables
    $NiLogKey = 'HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\NetSupport\NetInstall\LogFileSettings'
    $NiLogValue = 'ResolvedLogFilePath'

    Read-DSMInfo
    Remove-OutdatedFiles
}
#endregion

#region Functions
function Read-DSMInfo {
    Write-Debug -Message "start function Read-DSMInfo"
    If ($Local){
        # Use NiLog Path handed over by param
        $global:NiLogDir = $Path
        
    } else{
        # Read NiLog Dir from Registry
        $global:RegKey = "Registry::$NiLogKey"
        $global:NiLogDir = Get-Itemproperty -Path $Regkey -Name $NiLogValue | Select-Object -ExpandProperty $NiLogValue
    }
    Write-Debug -Message "end function Read-DSMInfo"
}
function Copy-DSMLogsToTempLocation {
    Write-Progress -Id 1 -Activity "Preparing DSM Log Files" -Status "robocopy is running"
    Write-Debug -Message "start function Copy-DSMLogsToTempLocation"
    if ($Debug){
        robocopy $NiLogDir $TempPath /mir /maxage:$DaysBack
    }else {
        robocopy $NiLogDir $TempPath /mir /maxage:$DaysBack | Out-Null
    }
    Write-Progress -Id 1 -Activity "Preparing DSM Log Files" -Status "robocopy is running" -Completed
    Write-Debug -Message "end function Copy-DSMLogsToTempLocation"
}
function Remove-OutdatedFiles {
    Write-Debug -Message "start function Remove-OutdatedFiles"
    $CurrentDate = Get-Date
    $DateToDelete = $CurrentDate.AddDays(-$DaysBack)
    Get-ChildItem $NiLogDir -Recurse | Where-Object { $_.LastWriteTime -lt $DatetoDelete } | Remove-Item -Recurse -Force
    Write-Debug -Message "end function Remove-OutdatedFiles"
}
function New-DSMLogArchive {
    Write-Debug -Message "start function Create-DSMLogArchive"
    # Copy Logs to Users Desktop
    If($PSVersionTable.PSVersion.Major -ge "5"){
        $ZIPPath = "$TargetPath\$Computername.zip"
        Compress-Archive -Path $TempPath -DestinationPath $ZIPPath -CompressionLevel Optimal -Force
    }
    Else{
        $CopyTargetPath = "$TargetPath","$env:COMPUTERNAME" -join "\"
        Copy-Item $TempPath $CopyTargetPath -recurse 
    }
    Write-Debug -Message "end function Create-DSMLogArchive"
}
function Remove-TempFiles {
    Write-Progress -Id 1 -Activity "Prepare DSM Logs" -Status "Cleanup Temp Folder"
    Write-Debug -Message "start function Remove-TempFiles"
    # Cleanup Temp Folder
    Get-ChildItem $TempPath | Remove-Item -Recurse -Force
    Remove-Item $TempPath -Force
    Write-Debug -Message "end function Remove-TempFiles"
}
#endregion

Export-ModuleMember -Function Get-DSMLogs
Export-ModuleMember -Function Remove-DSMLogs

#for testing only
#Remove-DSMLogs -DaysBack 12 -Debug
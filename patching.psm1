Using namespace System.Collections.Generic
Add-Type -AssemblyName PresentationFramework
[System.Net.WebRequest]::DefaultWebProxy.Credentials = [System.Net.CredentialCache]::DefaultCredentials
[System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls11 -bor [System.Net.SecurityProtocolType]::Tls12

#region Global Objects
#region Variables
[String]$uCMDBServerOS = '[CVS Data]:CVS OS'
[String]$uCMDBServerFqdn = '[CVS Data]:CVS FQDN'
[String]$uCMDBServerLabel = '[CVS Data]:Display Label'
[String]$uCMDBServerIPAddress = '[CVS Data]:CVS IP'
[String]$uCMDBServerPlatform = '[CVS Data]:CVS Platform'
[String]$uCMDBDecomStatus = '[CVS Data]:CVS Decom Status'
[String]$uCMDBAppInfoOS = 'Operating System - Version'
[String]$uCMDBAppInfoPlatform = 'Platform'
[String]$NotAvailable = 'N/A'
[String]$NotSet = ''
[String]$fn_ServerName = 'Server Name'
[String]$fn_Name = 'Name'
[String]$fn_IPAddress = 'IP Address'
[String]$fn_OperatingSystem = 'Operating System'
[String]$fn_AppName = 'Application Name'
[String]$fn_AppEnvironment = 'Application Environment'
[String]$fn_PrimarySOP = 'Primary SOP'
[String]$fn_SpecialSOP = 'Special SOP'
[String]$fn_SOP = 'SOP'
[String]$fn_ClusterName = 'Cluster Name'
[String]$fn_DRTier = 'DR Tier'
[String]$fn_IsTier1 = 'Is Tier 1'
[String]$fn_PrimaryShakeoutEmail = 'Primary Shakeout Email'
[String]$fn_PrimaryShakeoutPhone = 'Primary Shakeout Phone'
[String]$fn_SecondaryShakeoutEmail = 'Secondary Shakeout Email'
[String]$fn_SecondaryShakeoutPhone = 'Secondary Shakeout Phone'
[String]$fn_ShakeoutCompletionStatus = 'Shakeout Completion Status'
[String]$fn_PatchDay = 'Patch Day'
[String]$fn_PatchDate = 'Patch Date'
[String]$fn_PatchTime = 'Patch Time'
[String]$fn_PatchWeek = 'PatchWeek'
[String]$fn_PatchWindow = 'Patch Window'
[String]$fn_TimeZone = "Time Zone"
[String]$fn_StdTimeZone = "Std Time Zone"
[String]$fn_ITManagerFirst = 'IT Manager First'
[String]$fn_ITManagerLast = 'IT Manager Last'
[String]$fn_ITManagerEmail = 'IT Manager Email'
[String]$fn_ITDirectorFirst = 'IT Director First'
[String]$fn_ITDirectorLast = 'IT Director Last'
[String]$fn_ITDirectorEmail = 'IT Director Email'
[String]$fn_BusDirectorFirst = 'Bus Director First'
[String]$fn_BusDirectorLast = 'Bus Director Last'
[String]$fn_BusDirectorEmail = 'Bus Director Email'
[String]$fn_VPPOCFirst = 'VP POC First'
[String]$fn_VPPOCLast = 'VP POC Last'
[String]$fn_VPPOCEmail = 'VP POC Email'
[String]$fn_ITVPFirst = 'IT VP First'
[String]$fn_ITVPLast = 'IT VP Last'
[String]$fn_ITVPEmail = 'IT VP Email'
[String]$fn_PatchScheduleStatus = 'Patch Schedule Status'
[String]$fn_Comments1 = 'Comments 1'
[String]$fn_Comments2 = 'Comments 2'
[String]$fn_RecordStatus = 'Record Status'
[String]$fn_Decommissioned = 'Decommissioned'
[String]$fn_Retarget = 'Retarget'
[String]$fn_Exclusion = 'Exclusion'
[String]$NotSelected = 'NotSelected'
[String]$fn_SPItemID = 'SPItemID'
[String]$uCMDBAppObjects = $null
[String]$uCMDBServerObjects = $null
[String]$PreviousMonthMasterList = $null
[String]$ServerPatchList = $null
[String]$ExcelFilePath = $null
[String]$CSVFilePath = $null
[array]$arrResolvedEmailObjects = @()
$EmailFileName = "EmailAddressRetrieval $(Get-Date -Format MM-dd-yyyy-hh.mmtt).csv"
$EmailFolderName = $(Get-Date -Format MM-dd-yyyy)
$EmailLoggingFolder = "c:\Patching\EmailAddressLogs"
$EmailMappingsFolder = "c:\Patching\EmailAddressMappings"
$EmailFolderPath = "$EmailLoggingFolder\$EmailFolderName"
$EmailLogFilePath = "$EmailFolderPath\$EmailFileName"
$NewServerInput = $null
$CarryOver = $null
$WorkingDirectory = $null


# Importing email address correction mappings.
#$EmailAddressCorrections = Import-Csv -Path "$EmailMappingsFolder\EmailAddressCorrections.csv"
$EmailAddressAdditions = Import-Csv -Path "$EmailMappingsFolder\EmailAddressAdditions.csv"
$EmailAddressMappings = Import-Csv -Path "$EmailMappingsFolder\EmailAddressMappings.csv"

#region Production Fields
[Array]$ProductionFields = @{ Name = $fn_SPItemID; Expression = { $PSItem.FieldValues.ID } },
@{ Name = $fn_ServerName; Expression = { $PSItem.FieldValues.Title } },
@{ Name = $fn_IPAddress; Expression = { $PSItem.FieldValues.IP_x0020_Address } },
@{ Name = $fn_OperatingSystem; Expression = { $PSItem.FieldValues.Operating_x0020_System } },
@{ Name = $fn_AppName; Expression = { $PSItem.FieldValues.Application_x0020_Name } },
@{ Name = $fn_SOP; Expression = { $PSItem.FieldValues.SOP } },
@{ Name = $fn_PrimarySOP; Expression = { $PSItem.FieldValues.Primary_x0020_SOP } },
@{ Name = $fn_SpecialSOP; Expression = { $PSItem.FieldValues.Special_x0020_SOP } },
@{ Name = $fn_AppEnvironment; Expression = { $PSItem.FieldValues.Application_x0020_Environment } },
@{ Name = $fn_PatchWeek; Expression = { $PSItem.FieldValues.PatchWeek } },
@{ Name = $fn_PatchDay; Expression = { $PSItem.FieldValues.Patch_x0020_Day } },
@{ Name = $fn_PatchWindow; Expression = { $PSItem.FieldValues.Patch_x0020_Window } },
@{ Name = $fn_PatchDate; Expression = { $PSItem.FieldValues.Patch_x0020_Date } },
@{ Name = $fn_ITManagerFirst; Expression = { $PSItem.FieldValues.IT_x0020_Manager_x0020_First } },
@{ Name = $fn_ITManagerLast; Expression = { $PSItem.FieldValues.IT_x0020_Manager_x0020_Last } },
@{ Name = $fn_ITManagerEmail; Expression = {
        If ($RetrieveEmailAddresses) { (Get-EmailAddress -firstName $PSItem.FieldValues.IT_x0020_Manager_x0020_First -lastName $PSItem.FieldValues.IT_x0020_Manager_x0020_Last) }
        Else { $PSItem.FieldValues.IT_x0020_Manager_x0020_Email }
    }
},
@{ Name = $fn_ITDirectorFirst; Expression = { $PSItem.FieldValues.IT_x0020_Director_x0020_First } },
@{ Name = $fn_ITDirectorLast; Expression = { $PSItem.FieldValues.IT_x0020_Director_x0020_Last } },
@{ Name = $fn_ITDirectorEmail; Expression = {
        If ($RetrieveEmailAddresses) { (Get-EmailAddress -firstName $PSItem.FieldValues.IT_x0020_Director_x0020_First -lastName $PSItem.FieldValues.IT_x0020_Director_x0020_Last) }
        Else { $PSItem.FieldValues.IT_x0020_Director_x0020_Email }
    }
},
@{ Name = $fn_BusDirectorFirst; Expression = { $PSItem.FieldValues.Bus_x0020_Director_x0020_First } },
@{ Name = $fn_BusDirectorLast; Expression = { $PSItem.FieldValues.Bus_x0020_Director_x0020_Last } },
@{ Name = $fn_BusDirectorEmail; Expression = {
        If ($RetrieveEmailAddresses) { (Get-EmailAddress -firstName $PSItem.FieldValues.Bus_x0020_Director_x0020_First -lastName $PSItem.FieldValues.Bus_x0020_Director_x0020_Last) }
        Else { $PSItem.FieldValues.Bus_x0020_Director_x0020_Email }
    }
},
@{ Name = $fn_VPPOCFirst; Expression = { $PSItem.FieldValues.VP_x0020_POC_x0020_First } },
@{ Name = $fn_VPPOCLast; Expression = { $PSItem.FieldValues.VP_x0020_POC_x0020_Last } },
@{ Name = $fn_VPPOCEmail; Expression = {
        If ($RetrieveEmailAddresses) { (Get-EmailAddress -firstName $PSItem.FieldValues.VP_x0020_POC_x0020_First -lastName $PSItem.FieldValues.VP_x0020_POC_x0020_Last) }
        Else { $PSItem.FieldValues.VP_x0020_POC_x0020_Email }
    }
},
@{ Name = $fn_ITVPFirst; Expression = { $PSItem.FieldValues.IT_x0020_VP_x0020_First } },
@{ Name = $fn_ITVPLast; Expression = { $PSItem.FieldValues.IT_x0020_VP_x0020_Last } },
@{ Name = $fn_ITVPEmail; Expression = {
        If ($RetrieveEmailAddresses) { (Get-EmailAddress -firstName $PSItem.FieldValues.IT_x0020_VP_x0020_First -lastName $PSItem.FieldValues.IT_x0020_VP_x0020_Last) }
        Else { $PSItem.FieldValues.IT_x0020_VP_x0020_Email }
    }
},
@{ Name = $fn_PatchScheduleStatus; Expression = { $PSItem.FieldValues.Patch_x0020_Schedule_x0020_Statu } },
@{ Name = $fn_Comments1; Expression = { $PSItem.FieldValues.Comments_x002d_1 } },
@{ Name = $fn_Comments2; Expression = { $PSItem.FieldValues.Comments_x002d_2 } },
@{ Name = $fn_Decommissioned; Expression = { $PSItem.FieldValues.Decommissioned } },
@{ Name = $fn_ModerationStatus; Expression = { $PSItem.FieldValues._ModerationStatus } },
@{ Name = $fn_Retarget; Expression = { $PSItem.FieldValues.Retarget } }
#endregion Production Fields
#region Non-Production Fields
[Array]$NonProdFields = @{ Name = $fn_SPItemID; Expression = { $PSItem.FieldValues.ID } },
@{ Name = $fn_ServerName; Expression = { $PSItem.FieldValues.Title } },
@{ Name = $fn_IPAddress; Expression = { $PSItem.FieldValues.IP_x0020_Address } },
@{ Name = $fn_OperatingSystem; Expression = { $PSItem.FieldValues.Operating_x0020_System } },
@{ Name = $fn_AppName; Expression = { $PSItem.FieldValues.Application_x0020_Name } },
@{ Name = $fn_SOP; Expression = { $NotSet } },
@{ Name = $fn_PrimarySOP; Expression = { $PSItem.FieldValues.Primary_x0020_SOP } },
@{ Name = $fn_SpecialSOP; Expression = { $PSItem.FieldValues.Special_x0020_SOP } },
@{ Name = $fn_DrTier; Expression = { $PSItem.FieldValues.DR_x0020_Tier } },
@{ Name = $fn_AppEnvironment; Expression = { $PSItem.FieldValues.Application_x0020_Environment } },
@{ Name = $fn_PatchWeek; Expression = { $PSItem.FieldValues.PatchWeek } },
@{ Name = $fn_PatchDay; Expression = { $PSItem.FieldValues.Patch_x0020_Day } },
@{ Name = $fn_PatchWindow; Expression = { $PSItem.FieldValues.Patch_x0020_Window } },
@{ Name = $fn_PatchDate; Expression = { $PSItem.FieldValues.Patch_x0020_Date } },
@{ Name = $fn_ITManagerFirst; Expression = { $PSItem.FieldValues.IT_x0020_Manager_x0020_First } },
@{ Name = $fn_ITManagerLast; Expression = { $PSItem.FieldValues.IT_x0020_Manager_x0020_Last } },
@{ Name = $fn_ITManagerEmail; Expression = {
        If ($RetrieveEmailAddresses) { (Get-EmailAddress -firstName $PSItem.FieldValues.IT_x0020_Manager_x0020_First -lastName $PSItem.FieldValues.IT_x0020_Manager_x0020_Last) }
        Else { $PSItem.FieldValues.IT_x0020_Manager_x0020_Email }
    }
},
@{ Name = $fn_ITDirectorFirst; Expression = { $PSItem.FieldValues.IT_x0020_Director_x0020_First } },
@{ Name = $fn_ITDirectorLast; Expression = { $PSItem.FieldValues.IT_x0020_Director_x0020_Last } },
@{ Name = $fn_ITDirectorEmail; Expression = {
        If ($RetrieveEmailAddresses) { (Get-EmailAddress -firstName $PSItem.FieldValues.IT_x0020_Director_x0020_First -lastName $PSItem.FieldValues.IT_x0020_Director_x0020_Last) }
        Else { $PSItem.FieldValues.IT_x0020_Director_x0020_Email }
    }
},
@{ Name = $fn_BusDirectorFirst; Expression = { $PSItem.FieldValues.Bus_x0020_Director_x0020_First } },
@{ Name = $fn_BusDirectorLast; Expression = { $PSItem.FieldValues.Bus_x0020_Director_x0020_Last } },
@{ Name = $fn_BusDirectorEmail; Expression = {
        If ($RetrieveEmailAddresses) { (Get-EmailAddress -firstName $PSItem.FieldValues.Bus_x0020_Director_x0020_First -lastName $PSItem.FieldValues.Bus_x0020_Director_x0020_Last) }
        Else { $PSItem.FieldValues.Bus_x0020_Director_x0020_Email }
    }
},
@{ Name = $fn_VPPOCFirst; Expression = { $PSItem.FieldValues.VP_x0020_POC_x0020_First } },
@{ Name = $fn_VPPOCLast; Expression = { $PSItem.FieldValues.VP_x0020_POC_x0020_Last } },
@{ Name = $fn_VPPOCEmail; Expression = {
        If ($RetrieveEmailAddresses) { (Get-EmailAddress -firstName $PSItem.FieldValues.VP_x0020_POC_x0020_First -lastName $PSItem.FieldValues.VP_x0020_POC_x0020_Last) }
        Else { $PSItem.FieldValues.VP_x0020_POC_x0020_Email }
    }
},
@{ Name = $fn_ITVPFirst; Expression = { $PSItem.FieldValues.IT_x0020_VP_x0020_First } },
@{ Name = $fn_ITVPLast; Expression = { $PSItem.FieldValues.IT_x0020_VP_x0020_Last } },
@{ Name = $fn_ITVPEmail; Expression = {
        If ($RetrieveEmailAddresses) { (Get-EmailAddress -firstName $PSItem.FieldValues.IT_x0020_VP_x0020_First -lastName $PSItem.FieldValues.IT_x0020_VP_x0020_Last) }
        Else { $PSItem.FieldValues.IT_x0020_VP_x0020_Email }
    }
},
@{ Name = $fn_PatchScheduleStatus; Expression = { $PSItem.FieldValues.Patch_x0020_Schedule_x0020_Statu } },
@{ Name = $fn_Comments1; Expression = { $PSItem.FieldValues.Comments_x002d_1 } },
@{ Name = $fn_Comments2; Expression = { '$NotSet' } },
@{ Name = $fn_Decommissioned; Expression = { $PSItem.FieldValues.Decommissioned } },
@{ Name = $fn_Retarget; Expression = { $PSItem.FieldValues.Retarget } }
#endregion Non-Production Fields
#region Shakeout Fields
[Array]$ShakeoutFields = @{ Name = $fn_SPItemID; Expression = { $PSItem.FieldValues.ID } },
@{ Name = $fn_AppName; Expression = { $PSItem.FieldValues.Title } },
@{ Name = $fn_AppEnvironment; Expression = { $PSItem.FieldValues.Application_x0020_Environment } },
@{ Name = $fn_SOP; Expression = { $PSItem.FieldValues.SOP } },
@{ Name = $fn_PrimarySOP; Expression = { $PSItem.FieldValues.Primary_x0020_SOP } },
@{ Name = $fn_SpecialSOP; Expression = { $PSItem.FieldValues.Special_x0020_SOP } },
@{ Name = $fn_IsTier1; Expression = { $PSItem.FieldValues.Is_x0020_Tier_x0020_1 } },
@{ Name = $fn_PrimaryShakeoutEmail; Expression = { $PSItem.FieldValues.Primary_x0020_Shakeout_x0020_Ema } },
@{ Name = $fn_PrimaryShakeoutPhone; Expression = { $PSItem.FieldValues.Primary_x0020_Shakeout_x0020_Pho } },
@{ Name = $fn_SecondaryShakeoutEmail; Expression = { $PSItem.FieldValues.Secondary_x0020_Shakeout_x0020_E } },
@{ Name = $fn_SecondaryShakeoutPhone; Expression = { $PSItem.FieldValues.Secondary_x0020_Shakeout_x0020_P } },
@{ Name = $fn_ShakeoutCompletionStatus; Expression = { $PSItem.FieldValues.Shakeout_x0020_Completion_x0020_ } },
@{ Name = $fn_PatchWeek; Expression = { $PSItem.FieldValues.PatchWeek } },
@{ Name = $fn_PatchDay; Expression = { $PSItem.FieldValues.Patch_x0020_Day } },
@{ Name = $fn_PatchWindow; Expression = { $PSItem.FieldValues.Patch_x0020_Window } },
@{ Name = $fn_PatchDate; Expression = { $PSItem.FieldValues.Patch_x0020_Date } },
@{ Name = $fn_ITManagerFirst; Expression = { $PSItem.FieldValues.IT_x0020_Manager_x0020_First } },
@{ Name = $fn_ITManagerLast; Expression = { $PSItem.FieldValues.IT_x0020_Manager_x0020_Last } },
@{ Name = $fn_ITManagerEmail; Expression = {
        If ($RetrieveEmailAddresses) { (Get-EmailAddress -firstName $PSItem.FieldValues.IT_x0020_Manager_x0020_First -lastName $PSItem.FieldValues.IT_x0020_Manager_x0020_Last) }
        Else { $PSItem.FieldValues.IT_x0020_Manager_x0020_Email }
    }
},
@{ Name = $fn_ITDirectorFirst; Expression = { $PSItem.FieldValues.IT_x0020_Director_x0020_First } },
@{ Name = $fn_ITDirectorLast; Expression = { $PSItem.FieldValues.IT_x0020_Director_x0020_Last } },
@{ Name = $fn_ITDirectorEmail; Expression = {
        If ($RetrieveEmailAddresses) { (Get-EmailAddress -firstName $PSItem.FieldValues.IT_x0020_Director_x0020_First -lastName $PSItem.FieldValues.IT_x0020_Director_x0020_Last) }
        Else { $PSItem.FieldValues.IT_x0020_Director_x0020_Email }
    }
},
@{ Name = $fn_BusDirectorFirst; Expression = { $PSItem.FieldValues.Bus_x0020_Director_x0020_First } },
@{ Name = $fn_BusDirectorLast; Expression = { $PSItem.FieldValues.Bus_x0020_Director_x0020_Last } },
@{ Name = $fn_BusDirectorEmail; Expression = {
        If ($RetrieveEmailAddresses) { (Get-EmailAddress -firstName $PSItem.FieldValues.Bus_x0020_Director_x0020_First -lastName $PSItem.FieldValues.Bus_x0020_Director_x0020_Last) }
        Else { $PSItem.FieldValues.Bus_x0020_Director_x0020_Email }
    }
},
@{ Name = $fn_VPPOCFirst; Expression = { $PSItem.FieldValues.VP_x0020_POC_x0020_First } },
@{ Name = $fn_VPPOCLast; Expression = { $PSItem.FieldValues.VP_x0020_POC_x0020_Last } },
@{ Name = $fn_VPPOCEmail; Expression = {
        If ($RetrieveEmailAddresses) { (Get-EmailAddress -firstName $PSItem.FieldValues.VP_x0020_POC_x0020_First -lastName $PSItem.FieldValues.VP_x0020_POC_x0020_Last) }
        Else { $PSItem.FieldValues.VP_x0020_POC_x0020_Email }
    }
},
@{ Name = $fn_ITVPFirst; Expression = { $PSItem.FieldValues.IT_x0020_VP_x0020_First } },
@{ Name = $fn_ITVPLast; Expression = { $PSItem.FieldValues.IT_x0020_VP_x0020_Last } },
@{ Name = $fn_ITVPEmail; Expression = {
        If ($RetrieveEmailAddresses) { (Get-EmailAddress -firstName $PSItem.FieldValues.IT_x0020_VP_x0020_First -lastName $PSItem.FieldValues.IT_x0020_VP_x0020_Last) }
        Else { $PSItem.FieldValues.IT_x0020_VP_x0020_Email }
    }
}
#endregion Shakeout Fields

#endregion Variables

# Importing email address correction mappings.
$EmailAddressAdditions = Import-Csv -Path "$EmailMappingsFolder\EmailAddressAdditions.csv"
$EmailAddressMappings = Import-Csv -Path "$EmailMappingsFolder\EmailAddressMappings.csv"

# Settup Outlook.Application ComObject for the entire module to reuse.
$ol = New-Object -ComObject Outlook.Application

# Module based session variables #####################################################################################################################################################

# Setting up email address cache to eliminate redundant queries...
$global:arrResolvedEmailObjects = @()

# Create the logging folder if it doesn't exist...
if (!(Test-Path $EmailFolderPath)) { New-Item -ItemType Directory -Force -Path $EmailFolderPath }

#region Functions
Function Get-UCMDBFiles {
    $LibraryPath = "https://collab.corp.cvscaremark.com/sites/IT/ASD/Shared Documents" # Location of uCMDB files
    $Network = New-Object -ComObject WScript.Network
    $Network.MapNetworkDrive('R:', $LibraryPath) # Map a network drive. 
    $Files = @("R:\ucmdb application info*", "R:\ucmdb Server Report*")  # Only the files we need
    $host.ui.WriteVerboseLine("Downloading uCMDB files to local directory")
    Start-BitsTransfer $Files  # Uses BITS to copy the files
    $Network.RemoveNetworkDrive("R:")
}
Function Read-HostContinue {
    param (
        [Parameter(Position = 0)]
        [String]$PromptTitle = '',
        [Parameter(Position = 1)]
        [string]$PromptQuestion = 'Continue?',
        [Parameter(Position = 2)]
        [string]$YesDescription = 'Do this.',
        [Parameter(Position = 3)]
        [string]$NoDescription = 'Do not do this.',
        [Parameter(Position = 4)]
        [switch]$DefaultToNo,
        [Parameter(Position = 5)]
        [switch]$Force
    )
    if ($Force) {
        (-not $DefaultToNo)
        return
    }
    $yes = New-Object System.Management.Automation.Host.ChoiceDescription "&Yes", $YesDescription
    $no = New-Object System.Management.Automation.Host.ChoiceDescription "&No", $NoDescription

    if ($DefaultToNo) {
        $ConsolePrompt = [System.Management.Automation.Host.ChoiceDescription[]]($no, $yes)
    }
    else {
        $ConsolePrompt = [System.Management.Automation.Host.ChoiceDescription[]]($yes, $no)
    }
    if (($host.ui.PromptForChoice($PromptTitle, $PromptQuestion , $ConsolePrompt, 0)) -eq 0) {
        $true
    }
    else {
        $false
    }
}
Function Get-DaysToAdvance {
    <#
.SYNOPSIS
Get the number of days between Patch Tuesdays

.DESCRIPTION
Get the number of days between Patch Tuesdays so we know where to set the dates for the new schedule.  Should always be 28 or 35.

.PARAMETER Month
A string containing the full month name.  i.e. January, February.  This parameter is calculated by default and not required.
If running during the current month but after Patch Tuesday, you will be prompted to use the following month.  This shouldn't occur often,
but sometimes we will generate the next month's schedule before the 1st of the month.

.EXAMPLE
$Month, $Year, $Days = Get-DaysToAdvance
In April 2021, before Patch Tuesday, this will return: $Month = "April" and $Days = 35.

.NOTES
1.0 | First release
1.1 | Add '$Month' to Return statement
1.2 | Add '$Year' to Return statement
#>
    Param (
        $PatchTuesday
    )

    $Month = ([CultureInfo]::InvariantCulture).DateTimeFormat.GetMonthName($PatchTuesday.Month)
    $Base = Get-Date "$($PatchTuesday.Month)/$($PatchTuesday.Year)" -Day 12
    $Year = $Base.Year
    $Previous = $Base.AddMonths(-1)
    $PreviousPT = ($Previous.AddDays(2 - [int]$Previous.DayOfWeek))
    $DaysDiff = ($PatchTuesday - $PreviousPT).Days

    Return ($Month, $Year, $DaysDiff)
}
Function Get-PatchTuesday {
    Param(
        [string]$Month = ([CultureInfo]::InvariantCulture).DateTimeFormat.GetMonthName((Get-Date).Month)
    )

    $Base = (Get-Date $Month/1 -Day 12)
    $PatchTuesday = ($Base.AddDays(2 - [int]$Base.DayOfWeek))
    if ($PatchTuesday -lt (Get-Date)) {
        $Days = ((Get-Date) - $PatchTuesday).Days
        $UseNextMonth = Read-HostContinue -PromptTitle "Calculated Patch Tuesday - $($PatchTuesday.ToLongDateString()) is $($Days) ago." -PromptQuestion "Use next month?"
    }
    if ($UseNextMonth) {
        $Base = $Base.AddMonths(1)
        $PatchTuesday = ($Base.AddDays(2 - [int]$Base.DayOfWeek))
    }
    Return($PatchTuesday)
}
Function Get-FileName {
    Param (
        [Parameter(Mandatory = $false)]
        [String]$InitialDirectory = 'c:\Patching',
        [Parameter(Mandatory = $false)]
        [String]$Title = 'Select file to import...',
        [Parameter(Mandatory = $false)]
        $FilePreference = 'XLS'
    )

    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null
    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $OpenFileDialog.Title = $Title
    $OpenFileDialog.InitialDirectory = $InitialDirectory

    Switch ($FilePreference) {
        'CSV' { $OpenFileDialog.filter = "Comma Separated (*.csv)|*.csv|Text files (*.txt)|*.txt|Spreadsheets (*.xls?)|*.xls?|All files (*.*)|*.*" }
        'TXT' { $OpenFileDialog.filter = "Text files (*.txt)|*.txt|Comma Separated (*.csv)|*.csv|Spreadsheets (*.xls?)|*.xls?|All files (*.*)|*.*" }
        'XLS' { $OpenFileDialog.filter = "Spreadsheets (*.xls?)|*.xls?|Comma Separated (*.csv)|*.csv|Text files (*.txt)|*.txt|All files (*.*)|*.*" }
    }

    $OpenFileDialog.ShowDialog() | Out-Null
    $OpenFileDialog.filename
}
Function Convert-ExceltoCSV {
    # Converts input xls/xlsx files to csv.  For UCMDB, the extra lines at the beginning are removed.
    Param (
        $ExcelFilePath, #Input file path
        $CSVFilePath    #Output file path
    )
    $ExcelWB = New-Object -ComObject Excel.Application
    $ExcelWB.displayAlerts = $false
    $Workbook = $ExcelWB.Workbooks.Open($ExcelFilePath)
    $Workbook.SaveAs($CSVFilePath, 6)
    $Workbook.Close($false)
    $ExcelWB.Quit()
    $csvData = Get-Content $CSVFilePath
    If ($CSVFilePath -match "Server Report") {
        $host.ui.WriteVerboseLine("Removing headers from uCMDB Server Report file")
        $ServerHeaders = $csvData -match "\[CVS Data\]" | Select-Object -First 1
        $ToSkip = $csvData.IndexOf($ServerHeaders)
        # (Get-Content $csvFilePath | Select-Object -Skip $ToSkip) | Set-Content $csvFilePath
        $csvData | Select-Object -Skip $ToSkip | Set-Content $csvFilePath
    }
    ElseIf ($CSVFilePath -match "Application Info") {
        $host.ui.WriteVerboseLine("Removing headers from uCMDB Application Info Report file")
        $AppHeaders = $csvData -match "Business Unit" | Select-Object -First 1
        $ToSkip = $csvData.IndexOf($AppHeaders)
        # (Get-Content $csvFilePath | Select-Object -Skip $ToSkip) | Set-Content $csvFilePath
        $csvData | Select-Object -Skip $ToSkip | Set-Content $csvFilePath
    }
}
Function Get-PatchDay {
    Param ([string]$Day)
    $Day = $Day.split('/')[0].trim()
    If ([int]$Day -in 9 .. 15) {
        $Day = ([int]$Day - 7).ToString()
    }
    ElseIf ([int]$Day -in 16 .. 21) {
        $Day = ([int]$Day - 14).ToString()
    }
    Return ($Day)
}
Function Build-Record {
    Param (
        $SPRec,
        $InvRec,
        $TZHash
    )
    $PDay = $SPRec.'Patch Days'
    [pscustomobject] @{
        'Server Name'                = $SPRec.DomainServer
        'IP Address'                 = $InvRec.$fn_IPAddress
        'Operating System'           = $SPRec.OS
        'Application Name'           = $SPRec.Application
        'Application Environment'    = If ([System.String]::IsNullOrWhiteSpace($SPRec.'App Env')) {
            ""
        }
        Else {
            $($SPRec.'app env'.split(",", [System.StringSplitOptions]::RemoveEmptyEntries).Trim().ToUpper() | Sort-Object -Unique) -join (', ')
        }
        'Primary SOP'                = $InvRec.$fn_PrimarySOP
        'Special SOP'                = $invRec.$fn_SpecialSOP
        'SOP'                        = $InvRec.$fn_SOP
        'Cluster Name'               = $InvRec.$fn_ClusterName
        'DR Tier'                    = $InvRec.$fn_DRTier
        'Is Tier 1'                  = $InvRec.$fn_IsTier1
        'Primary Shakeout Email'     = $InvRec.$fn_PrimaryShakeoutEmail
        'Primary Shakeout Phone'     = $InvRec.$fn_PrimaryShakeoutPhone
        'Secondary Shakeout Email'   = $InvRec.$fn_SecondaryShakeoutEmail
        'Secondary Shakeout Phone'   = $InvRec.$fn_SecondaryShakeoutPhone
        'Shakeout Completion Status' = $InvRec.$fn_ShakeoutCompletionStatus
        'Patch Day'                  = $SPRec.'Patch Days'
        'Patch Date'                 = $($PatchTuesday.AddDays($PDay).ToShortDateString())
        'Patch Time'                 = $SPRec.'Patch Times'
        'Patch Window'               = $InvRec.$fn_PatchWindow
        'IT Manager First'           = $InvRec.$fn_ITManagerFirst
        'IT Manager Last'            = $InvRec.$fn_ITManagerLast
        'IT Manager Email'           = $InvRec.$fn_ITManagerEmail
        'IT Director First'          = $InvRec.$fn_ITDirectorFirst
        'IT Director Last'           = $InvRec.$fn_ITDirectorLast
        'IT Director Email'          = $InvRec.$fn_ITDirectorEmail
        'Bus Director First'         = $InvRec.$fn_BusDirectorFirst
        'Bus Director Last'          = $InvRec.$fn_BusDirectorLast
        'Bus Director Email'         = $InvRec.$fn_BusDirectorEmail
        'VP POC First'               = $InvRec.$fn_VPPOCFirst
        'VP POC Last'                = $InvRec.$fn_VPPOCLast
        'VP POC Email'               = $InvRec.$fn_VPPOCEmail
        'IT VP First'                = $InvRec.$fn_ITVPFirst
        'IT VP Last'                 = $InvRec.$fn_ITVPLast
        'IT VP Email'                = $InvRec.$fn_ITVPEmail
        'Comments 1'                 = $InvRec.$fn_Comments1
        'Comments 2'                 = $InvRec.$fn_Comments2
        'Retarget'                   = $InvRec.$fn_Retarget
        'Record Status'              = $InvRec.$fn_RecordStatus
        'Time Zone'                  = $SPRec.'Server Time Zone'
        'Std Time Zone'              = $TZHash[$($SPRec.'Server Time Zone'.split('(')[0].trim())]
    }
}
Function Show-Progress {
    <#
        .SYNOPSIS
            Show progress as items pass through a section of the pipline
        .DESCRIPTION
            This function allows you to show progress from the pipeline.
            Its intentionally written for efficiency/speed so some compromises on readability were made
        .PARAMETER InputObject
            The items on the pipeline being processed
        .PARAMETER Activity
            The activity being measured
        .PARAMETER UpdatePercentage
            The percentage of time to update the progress bar.
            Write-Progress is a slow cmdlet so this is used for performance reasons with larger data sets
        .EXAMPLE
            # This runs through the numbers 1 through 1000 and displays a progress bar based on how many have gone through the pipeline
            1..1000 | Show-Progress
        .EXAMPLE
            # Showing progress with a specific activity message and only updating at 10%, 20% etc
            1..1000 | Show-Progress -Activity "doin stuff" -UpdatePercentage 10
        #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory, Position = 0, ValueFromPipeline)]
        [PSObject[]]
        $InputObject,

        [string]
        $Activity = "Processing items",

        [ValidateRange(1, 100)]
        [int]
        $UpdatePercentage
    )

    $count = 0
    [int]$totalItems = $Input.count
    $progressWritten = @()

    # use a dot sourced scriptblock to loop despite its lower redability as its extremely fast
    $Input | . {
        process {
            # pass thru the input
            $_

            $count++

            [int]$percentComplete = ($Count / $totalItems * 100)

            $writeProgressSplat = @{
                Activity        = $Activity
                PercentComplete = $percentComplete
                Status          = "Working - $percentComplete%"
            }

            if ($percentComplete -notin $progressWritten -and ($UpdatePercentage -eq 0 -or $percentComplete % $UpdatePercentage -eq 0)) {
                $progressWritten += $percentComplete
                Write-Progress @writeProgressSplat
            }
        }
    }
}
function Create-MonthlySharePointLists {
    <#
.SYNOPSIS
The Create-MonthlySharePointLists function generates the three SharePoint lists 
(Non-Prod, Prod, and Shakeout) to be imported into SharePoint.

.DESCRIPTION
The Create-MonthlySharePointLists function will require the location of last month's 
Master Inventory Patching Schedule.

.PARAMETER MasterServerListPath
The MasterServerListPath parameter is asking for the full filesystem path to last month's 
Master Inventory Patching Schedule.

.PARAMETER Month
The Month parameter is optional: but it's a good idea to specify the month of the new patching schedule.

.PARAMETER Year
The Year parameter is option: but again it's a good idea to enter the year of the next patching cycle.

.EXAMPLE
PS C:\> Create-MonthlySharePointLists -Month October -Year 2018

.NOTES
Once you call this function it will open a 'FileOpenDialog', browse to the files required and click Open.
#>
    Param(

        [Parameter(Mandatory = $false, HelpMessage = "Full path to Master Server List to be used as source for monthly SharePoint lists...")]
        [String]$MasterServerListPath = $(Get-FileName -Title "Select Master Server List to be used as source for monthly SharePoint lists..." -InitialDirectory "C:\Patching\$($Year)\$($Month)\Schedule" -FilePreferrence 'CSV' ),

        [Parameter(Mandatory = $false, HelpMessage = "Enter the target month name i.e. 'July' where you want the results to go... Leave blank for current month...")]
        [ValidateSet('January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December')]
        [String]$Month = (Get-Culture).DateTimeFormat.GetMonthName((Get-Date -f MM)),

        [Parameter(Mandatory = $false, HelpMessage = "Enter the target year i.e. '2018' where you want the results to go... Leave blank for current year...")]
        [String]$Year = (Get-Date -f yyyy)

    )

    $WorkingDirectory = "c:\Patching\$($Year)\$($Month)\Schedule"
    Write-Host "Switching location to working directory $WorkingDirectory..."
    Set-Location -Path $WorkingDirectory

    Write-Host "Importing $MasterServerListPath..." -ForegroundColor Green
    $MasterServerList = Import-Csv -Path $MasterServerListPath | Where-Object { $PSItem.$fn_PatchDay -match "\d" }
    Write-Host "Master Server Inventory List '$MasterServerListPath' imported successfully..." -ForegroundColor Green

    # Creating Production Patching Schedule...
    $ProductionSchedulePath = "$WorkingDirectory\$Month $Year - Production Servers Schedule.csv"
    Write-Host "Creating $Month $Year - Production Patching Schedule at: $ProductionSchedulePath..."

    $MasterServerList |
    Where-Object { ([int]$PSItem.$fn_PatchDay -ne 1) } |
    Sort-Object -Property $fn_ServerName, $fn_IPAddress, $fn_AppName, $fn_AppEnvironment, $fn_PatchDay -Unique |
    Select-Object -Property $fn_ServerName, $fn_IPAddress, $fn_OperatingSystem, $fn_AppName, $fn_AppEnvironment, $fn_PrimarySOP, $fn_SpecialSOP, $fn_SOP, $fn_IsTier1, $fn_PatchDay, $fn_PatchDate, $fn_PatchTime, $fn_TimeZone, $fn_ITManagerFirst, $fn_ITManagerLast, $fn_ITManagerEmail, $fn_ITDirectorFirst, `
        $fn_ITDirectorLast, $fn_ITDirectorEmail, $fn_BusDirectorFirst, $fn_BusDirectorLast, $fn_BusDirectorEmail, $fn_VPPOCFirst, $fn_VPPOCLast, $fn_VPPOCEmail, $fn_ITVPFirst, $fn_ITVPLast, `
        $fn_ITVPEmail, $fn_PatchScheduleStatus, $fn_Comments1, $fn_Comments2, $fn_Decommissioned, $fn_Retarget | Export-Csv -Path $ProductionSchedulePath -NoTypeInformation

    Convert-ToExcelTable -SourceCsv $ProductionSchedulePath


    # Creating Non-Production Patching Schedule...
    $NonProductionSchedulePath = "$WorkingDirectory\$Month $Year - NON Production Servers Schedule.csv"
    Write-Host "Creating $Month $Year - Non-Production Patching Schedule at: $NonProductionSchedulePath..."

    $MasterServerList | Where-Object { [int]$PSItem.$fn_PatchDay -eq 1 } |
    Sort-Object -Property $fn_ServerName, $fn_IPAddress, $fn_AppName, $fn_AppEnvironment, $fn_PatchDay -Unique |
    Select-Object -Property $fn_ServerName, $fn_IPAddress, $fn_OperatingSystem, $fn_AppName, $fn_AppEnvironment, $fn_PrimarySOP, $fn_SpecialSOP, $fn_SOP, $fn_IsTier1, $fn_PatchDay, $fn_PatchDate, $fn_PatchTime, $fn_TimeZone, $fn_ITManagerFirst, $fn_ITManagerLast, $fn_ITManagerEmail, $fn_ITDirectorFirst, `
        $fn_ITDirectorLast, $fn_ITDirectorEmail, $fn_BusDirectorFirst, $fn_BusDirectorLast, $fn_BusDirectorEmail, $fn_VPPOCFirst, $fn_VPPOCLast, $fn_VPPOCEmail, $fn_ITVPFirst, $fn_ITVPLast, `
        $fn_ITVPEmail, $fn_PatchScheduleStatus, $fn_Comments1, $fn_Comments2, $fn_Decommissioned, $fn_Retarget | Export-Csv -Path $NonProductionSchedulePath -NoTypeInformation

    Convert-ToExcelTable -SourceCsv $NonProductionSchedulePath

    # Remove 'Non-Production Wednesday', and 'Do Not Patch' from Shakeout list
    # Creating Shakeout Resources and Application Schedule
    $ShakeoutSchedulePath = "$WorkingDirectory\$Month $Year - Shakeout Resources and Application Schedule.csv"
    Write-Host "Creating $Month $Year - Shakeout Resources and Application Schedules list at: $ShakeoutSchedulePath..."

    $MasterServerList | Where-Object {
        ( [int]$PSItem.$fn_PatchDay -ne 1 ) -and
        (-not ([system.string]::IsNullOrWhiteSpace($PSItem.$fn_AppName))) -and
        ($PSItem.$fn_AppEnvironment -in @("Production", "DR"))
    } | Sort-Object -Property $fn_AppName, $fn_AppEnvironment, $fn_IsTier1, $fn_PatchDay, $fn_PatchWindow -Unique |
    Select-Object -Property $fn_AppName, $fn_AppEnvironment, $fn_SOP, $fn_IsTier1, $fn_PrimaryShakeoutEmail, $fn_PrimaryShakeoutPhone, $fn_SecondaryShakeoutEmail, $fn_SecondaryShakeoutPhone, `
        $fn_ShakeoutCompletionStatus, $fn_PatchDay, $fn_PatchDate, $fn_PatchTime, $fn_TimeZone, $fn_ITManagerFirst, $fn_ITManagerLast, $fn_ITManagerEmail, $fn_ITDirectorFirst, $fn_ITDirectorLast, `
        $fn_ITDirectorEmail, $fn_BusDirectorFirst, $fn_BusDirectorLast, $fn_BusDirectorEmail, $fn_VPPOCFirst, $fn_VPPOCLast, $fn_VPPOCEmail, $fn_ITVPFirst, $fn_ITVPLast, `
        $fn_ITVPEmail, $fn_Comments1 | Export-Csv -Path $ShakeoutSchedulePath -NoTypeInformation

    Convert-ToExcelTable -SourceCsv $ShakeoutSchedulePath
}
function Get-EmailAddressMapping {
    Param(
        [Parameter(Mandatory = $true)][String]$firstName,
        [Parameter(Mandatory = $true)][String]$lastName
    )

    $MappingResults = $EmailAddressMappings | Where-Object( { $PSItem.FirstName -eq $firstName -and $PSItem.LastName -eq $lastName })
    if ($MappingResults) {
        # We return what we know is correct based on feedback gathered in the $EmailAddressMappings.csv
        Write-Verbose "Email Address mapping found..."
        Return $MappingResults.EmailAddress
    }
    else {
        # We return $null if a firstname, lastname match was found.
        Return $null
    }
}
function Get-EmailAddress {
    [CmdLetBinding()]
    Param(

        [Parameter(Mandatory = $true)][AllowEmptyString()][string]$firstName,
        [Parameter(Mandatory = $true)][AllowEmptyString()][string]$lastName
    )

    # Saving a copy of the params as we modify them later on
    $cachedFirstName = $firstName
    $cachedLastName = $lastName

    # Before we search the GAL, we check to see if we've processed this firstName, $initial, $lastName already.
    $cachedEmailObject = $global:arrResolvedEmailObjects.Where( { ($PSItem.FirstName -eq $cachedFirstName) -and ($PSItem.LastName -eq $cachedLastName) })

    if ($cachedEmailObject[0].FirstName -eq $cachedFirstName -and $cachedEmailObject[0].LastName -eq $cachedLastName) {

        $cachedEmailObject[0].Results = "FOUND USING CACHE"
        $cachedEmailObject[0] | Export-Csv -Path $EmailLogFilePath -NoTypeInformation -Append

        return $cachedEmailObject[0].EmailAddress

    }
    else {

        if ([System.String]::IsNullOrWhiteSpace($firstName) -or [System.String]::IsNullOrWhiteSpace($lastName) -or ("N/A" -match @("$firstname|$lastname"))) {

            $newEmailObject = [PSCustomObject]@{
                FirstName    = $cachedFirstName
                LastName     = $cachedLastName
                EmailAddress = 'NotFound'
                Results      = "Manager Info is NULL"
            }

            $global:arrResolvedEmailObjects += $newEmailObject
            $newEmailObject | Export-Csv -Path $EmailLogFilePath -NoTypeInformation -Append

            return 'NotFound'
        }

        $Commands = @()
        $galEmailAddress = ""

        # Grabbing initial, and trimming firstname before using
        $initial = $firstName.Split(" ", 2)[1]
        $firstName = $firstName.Split(" ", 2)[0]

        # Based on what I've found, these are the known issues with the lastname...
        # We need to remove anything after a "(" or " (" string.
        if ($lastName.Contains("(")) {
            $lastName = $lastName.Split(" (", 2)[0]
            $lastName = $lastName.Split("(", 2)[0]
        }

        # We need to remove the string ' III' from any last name.
        if ($lastName.Contains("III")) {
            $lastName = $lastName.Replace(" III", "")
        }

        # We need to remove anything after a " - " string..
        if ($lastName.Contains("-")) {
            $lastName = $lastName.Split(" - ", 2)[0]
        }

        # We need to escape the "'" character.
        if ($lastName.Contains("'")) {
            $lastName = $lastName.Replace("'", "''")
        }

        #region Begin EmailAddressMappings ##################################################################################################################################################
        # If we are here; the cache hit was a miss, we now search the EmailAddressMappings file for a match.

        $mappingEmailAddress = Get-EmailAddressMapping -firstName $cachedFirstName -lastName $cachedLastName

        if ($mappingEmailAddress) {

            # Caching results...
            $newEmailObject = [PSCustomObject]@{
                FirstName    = $cachedFirstName
                LastName     = $cachedLastName
                EmailAddress = $mappingEmailAddress
                Results      = "FOUND USING MAPPING FILE"
            }

            $global:arrResolvedEmailObjects += $newEmailObject
            $newEmailObject | Export-Csv -Path $EmailLogFilePath -NoTypeInformation -Append

            return $mappingEmailAddress

        }
        #endregion Begin EmailAddressMappings

        #region Begin GAL ##################################################################################################################################################################
        # If we are here; the cache hit and mappings file was a miss, we now search the GAL using Outlook for a match.

        # Building $commands array: first we exclude the middle initial to find email address
        $Commands += "& Search-GAL -searchString '$lastName, $firstName'"
        # $Commands += "& Search-GAL -searchString '$firstName $lastName'"

        # Building $commands array: second we include the middle initial (if it exists) to find email address
        if ($initial) {
            $Commands += "& Search-GAL -searchString '$lastName, $firstName $initial'"
            # $Commands += "& Search-GAL -searchString '$firstName $initial $lastName'"
        }


        # We iterate through the $commands array executing and testing return for a strict 'firstname, lastname' match.
        foreach ($Command in $Commands) {
            $galEmailAddress = Invoke-Expression -Command $Command
            # Checking for NULL
            if ($galEmailAddress) {
                # Making sure it can be split into recipient/domain parts
                <# if ($galEmailAddress.Contains("@")) {

$emailPrefix = $galEmailAddress.Split("@", 2)[0]

# if there's a lastname we need to verify email address first and last names match.
if ($emailPrefix.Split(".", 2).Count -eq 2) {

$emailFirstName = $emailPrefix.Split(".", 2)[0]
$emailLastName = $emailPrefix.Split(".", 2)[1]

if (($emailFirstName -eq $firstName) -and ($emailLastName -eq $lastName)) { #>

                # Caching results...
                $newEmailObject = [PSCustomObject]@{
                    FirstName    = $cachedFirstName
                    LastName     = $cachedLastName
                    EmailAddress = $galEmailAddress
                    Results      = "FOUND USING GAL"
                }

                $global:arrResolvedEmailObjects += $newEmailObject
                $newEmailObject | Export-Csv -Path $EmailLogFilePath -NoTypeInformation -Append

                <# return $galEmailAddress

}
}
} #>
                return $galEmailAddress
            }
        }
        #endregion End GAL ####################################################################################################################################################################

        #region Begin AD GC ################################################################################################################################################################
        # If we are here; cache, mappings, and GAL were all a miss, trying AD GC

        # We try AD GC with initial first, then without...
        if ($initial) {

            $adEmailAddress = (Get-ADUser -Filter { (GivenName -eq $firstName) -and (Initials -eq $initial) -and (SurName -eq $lastName) -and (SamAccountName -notlike 'ADM*') -and (SamAccountName -notlike 'A_*') -and (mail -like "*") } `
                    -Properties mail -Server corp.cvscaremark.com:3268).mail

            # Checking if the return is NOT NULL...
            if ($adEmailAddress) {
                # Checking if the return is an array
                if ($adEmailAddress.Count -ge 2) {
                    # Looping through the array of email addresses
                    foreach ($adEmailAdr in $adEmailAddress) {
                        # Checking for NULL array elements
                        if ($adEmailAdr -ne $null) {
                            # Verifying this is an email address.
                            if ($adEmailAdr.Contains("@")) {
                                # Grabbing the recipient portion of the email address
                                $adEmailPrefix = $adEmailAdr.Split("@", 2)[0]

                                # if there's a lastname we need to verify email address first and last names match.
                                if ($adEmailPrefix.Split(".", 2).Count -eq 2) {

                                    $adEmailFirstName = $adEmailPrefix.Split(".", 2)[0]
                                    $adEmailLastName = $adEmailPrefix.Split(".", 2)[1]

                                    if (($adEmailFirstName -eq $firstName) -and ($adEmailLastName -eq $lastName)) {

                                        # Caching results...
                                        $newEmailObject = [PSCustomObject]@{
                                            FirstName    = $cachedFirstName
                                            LastName     = $cachedLastName
                                            EmailAddress = $adEmailAdr
                                            Results      = "FOUND WITH INITIAL USING AD-GC"
                                        }

                                        $global:arrResolvedEmailObjects += $newEmailObject
                                        $newEmailObject | Export-Csv -Path $EmailLogFilePath -NoTypeInformation -Append

                                        return $adEmailAdr

                                    }
                                }
                            }
                        }
                    }
                }
                else {

                    # Caching results...
                    $newEmailObject = [PSCustomObject]@{
                        FirstName    = $cachedFirstName
                        LastName     = $cachedLastName
                        EmailAddress = $adEmailAddress
                        Results      = "FOUND USING AD-GC"
                    }

                    $global:arrResolvedEmailObjects += $newEmailObject
                    $newEmailObject | Export-Csv -Path $EmailLogFilePath -NoTypeInformation -Append

                    return $adEmailAddress

                }
            }

        }

        # If we are here; we now try to resolve address without using initial against AD GC
        $adEmailAddress = (Get-ADUser -Filter { (GivenName -eq $firstName) -and (SurName -eq $lastName) -and (SamAccountName -notlike 'ADM*') -and (SamAccountName -notlike 'A_*') -and (mail -like "*") } `
                -Properties mail -Server corp.cvscaremark.com:3268).mail

        # Checking if the return is NOT NULL...
        if ($adEmailAddress) {
            # Checking if the return is an array
            if ($adEmailAddress.Count -ge 2) {
                # Looping through the array of email addresses
                foreach ($adEmailAdr in $adEmailAddress) {
                    # Checking for NULL array elements
                    if ($adEmailAdr -ne $null) {
                        # Verifying this is an email address.
                        if ($adEmailAdr.Contains("@")) {
                            # Grabbing the recipient portion of the email address
                            $adEmailPrefix = $adEmailAdr.Split("@", 2)[0]

                            # if there's a lastname we need to verify email address first and last names match.
                            if ($adEmailPrefix.Split(".", 2).Count -eq 2) {

                                $adEmailFirstName = $adEmailPrefix.Split(".", 2)[0]
                                $adEmailLastName = $adEmailPrefix.Split(".", 2)[1]

                                if (($adEmailFirstName -eq $firstName) -and ($adEmailLastName -eq $lastName)) {

                                    # Caching results...
                                    $newEmailObject = [PSCustomObject]@{
                                        FirstName    = $cachedFirstName
                                        LastName     = $cachedLastName
                                        EmailAddress = $adEmailAdr
                                        Results      = "FOUND USING AD-GC"
                                    }

                                    $global:arrResolvedEmailObjects += $newEmailObject
                                    $newEmailObject | Export-Csv -Path $EmailLogFilePath -NoTypeInformation -Append

                                    return $adEmailAdr

                                }
                            }
                        }
                    }
                }
            }
            else {

                # Caching results...
                $newEmailObject = [PSCustomObject]@{
                    FirstName    = $cachedFirstName
                    LastName     = $cachedLastName
                    EmailAddress = $adEmailAddress
                    Results      = "FOUND USING AD-GC"
                }

                $global:arrResolvedEmailObjects += $newEmailObject
                $newEmailObject | Export-Csv -Path $EmailLogFilePath -NoTypeInformation -Append

                return $adEmailAddress

            }
        }
        else {

            # Caching results...
            $newEmailObject = [PSCustomObject]@{
                FirstName    = $cachedFirstName
                LastName     = $cachedLastName
                EmailAddress = 'NotFound'
                Results      = "NOT FOUND"
            }

            $global:arrResolvedEmailObjects += $newEmailObject
            $newEmailObject | Export-Csv -Path $EmailLogFilePath -NoTypeInformation -Append

            return 'NotFound'

        }
        #endregion End AD GC ##################################################################################################################################################################

    }
}
function Search-GAL {
    [CmdLetBinding()]
    Param (
        [Parameter(Mandatory = $true)][string]$searchString
    )

    Begin {
        if (!$ol) {
            $ol = New-Object -ComObject Outlook.Application
        }
    }

    Process {
        $item = $ol.Session.GetGlobalAddressList().AddressEntries.Item($searchString)
        $emailAddress = $item.GetExchangeUser().PrimarySmtpAddress

        return $emailAddress
    }

    End { }
}
function Get-AdditionalEmailAddress {
    <#
.SYNOPSIS
Get addtional email address(es) to include with IT Manager's address

.DESCRIPTION
Uses a lookup against the EmailAddressAdditions.csv file in the EmailAddressMappings folder

.PARAMETER EmailAddress
Input Email address for the IT Manager

.OUTPUTS
Secondary email addresses separated by ';' to be appended to the IT Manager's email address
#>
    [cmdletbinding()]
    Param($EmailAddress)
    Begin {
        $AdditionalAddresses = Import-Csv C:\Patching\EmailAddressMappings\EmailAddressAdditions.csv
    }

    Process {
        $Addresses = ($AdditionalAddresses | Where-Object { $PSItem.ITManagerEmail.trim() -eq $EmailAddress.trim() }).AdditionalEmail
        $AdditionalAddress = $Addresses -join ("; ")
    }
    End {
        if ($AdditionalAddress) {
            Return("; " + $AdditionalAddress)
        }
    }
}
function Get-TmpTableRows {
    Param($ShakeOutUrl, $PatchAppOwner)
    $TableRow = @"
<li><a href="$ShakeOutUrl"><span style='mso-bookmark:_MailOriginal'>$PatchAppOwner</a></li>
"@
    Return($TableRow)
}
function New-ShakeoutNotification {
    <#
.SYNOPSIS
A brief description of the New-ShakeoutNotification function.

.DESCRIPTION
A detailed description of the New-ShakeoutNotification function.

.PARAMETER Escalate
A description of the Escalate parameter.

.PARAMETER Type
A description of the Type parameter.

.PARAMETER Month
A description of the Month parameter.

.PARAMETER Year
A description of the Year parameter.

.EXAMPLE
PS C:\> New-ShakeoutNotification

.NOTES
Additional information about the function.
#>
    [CmdletBinding()]
    param
    (
        [switch]$Escalate,
        [ValidateSet('Resource', 'Completion')]
        [String]$Type,
        [String]$Month,
        [String]$Year
    )

    $ShakeoutInfo = Get-SharePointOnlineListItem -Month $Month -Year $Year -PatchScheduleType ShakeoutSchedule | Where-Object {
        -not ([string]::IsNullOrWhiteSpace($_.$fn_AppName)) -and ([string]::IsNullOrWhiteSpace($_.$fn_IsTier1))
    }
    $ProdShakeoutListname = Format-UrlListNameValue -ListName "$Month $Year - Shakeout Resources and Application Schedule".Substring(0, 51).Trim()
    $PatchAppOwnerFieldName = Format-UrlFilterFieldName -FieldNameToFormat $fn_ITManagerEmail
    $NSORecords = @()
    $Filter2 = "FilterField2=Primary%5Fx0020%5FShakeout%5Fx0020%5FEma&FilterValue2=&FilterType2=Text&FilterField3=Is%5Fx0020%5FTier%5Fx0020%5F1&FilterValue3=&FilterType3=Text"
    $Filter3 = "FilterField2=Shakeout%5Fx0020%5FCompletion%5Fx0020%5F&FilterValue2=&FilterType2=Text"

    if (!$ol) {
        $ol = New-Object -ComObject Outlook.Application
    }


    if ($Type -eq 'Resource') {
        [array]$MissingInfo = $ShakeoutInfo | Where-Object { 
([string]::IsNullOrWhiteSpace($_.$fn_PrimaryShakeoutEmail)) -and
(-not ([system.string]::IsNullOrWhiteSpace($_.$fn_ITManagerEmail)))
        }
        $BodyNote = "Below is a list of IT Manager/AppOwners who have yet to provide shakeout resources for some or all of their applications.  Please follow the link for your name and provide Primary and Secondary resources."
    }
    elseif ($Type -eq 'Completion') {
        [array]$MissingInfo = $ShakeoutInfo | Where-Object {
(-not [string]::IsNullOrWhiteSpace($_.$fn_PrimaryShakeoutEmail)) -and
([string]::IsNullOrWhiteSpace($_.$fn_ShakeoutCompletionStatus)) -and
([string]::IsNullOrWhiteSpace($_.$fn_IsTier1)) -and
([datetime]$_.$fn_PatchDate -lt (Get-Date).AddDays(-3).Date) -and
(-not ([system.string]::IsNullOrWhiteSpace($_.$fn_ITManagerEmail)))
        }
        $BodyNote = "Below is a list of IT Manager/AppOwners who have yet to provide shakeout completion status for some or all of their applications.  Please follow the link for your name and provide the completion status for applications that have had patching completed."
    }

    $MissingInfo = $MissingInfo | Sort-Object $fn_ITManageremail -Unique
    $MissingCount = $($MissingInfo.count)

    $NSORecords = $MissingInfo | ForEach-Object {
        $Email = $($_.$fn_ITManageremail).Trim()
        $Name = "$($_.$fn_ITManagerLast), $($_.$fn_ITManagerFirst)"

        if ($Type -eq "Resource") {
            $URL = "https://aetnao365.sharepoint.com/Sites/WinCompliance/Patching/Lists/$ProdShakeoutListName/AllItems.aspx?FilterField1=$PatchAppOwnerFieldName&FilterValue1=$($Email)&$($Filter2)"
        }
        elseif ($Type -eq "Completion") {
            $URL = "https://aetnao365.sharepoint.com/Sites/WinCompliance/Patching/Lists/$ProdShakeoutListName/AllItems.aspx?FilterField1=$PatchAppOwnerFieldName&FilterValue1=$($Email)&$($Filter3)"
        }
        [PSCustomObject]@{
            Name  = $Name
            Email = $Email
            URL   = $URL
        }
    }

    $NSORecords = $NSORecords | Sort-Object -Property Name -Unique
    If ($NSORecords.count -eq 0) {
        Write-Output "There are no IT Managers to contact at this time."
        Start-Sleep -Seconds 2
        break
    }
    $MissingCount = $NSORecords.count
    $Addresses = (($MissingInfo.$fn_ITManagerEmail | Where-Object { $_ -match "@" }).split(@(";", ",", "/", " "), [System.StringSplitOptions]::RemoveEmptyEntries) -match "@").Trim()
    # $Addresses += ($MissingInfo.$fn_PrimaryShakeoutEmail -match "@").split(@(";",",","/"," "),[System.StringSplitOptions]::RemoveEmptyEntries) -match "@"
    # $Addresses += ($MissingInfo.$fn_SecondaryShakeoutEmail -match "@").split(@(";",",","/"," "),[System.StringSplitOptions]::RemoveEmptyEntries) -match "@"
    $AddressList = ($Addresses = $Addresses | Sort-Object -Unique) -join ("; ")

    $Directors = (($MissingInfo.$fn_ITDirectorEmail | Where-Object { $_ -match "@" }).split(@(";", ",", "/", " "), [System.StringSplitOptions]::RemoveEmptyEntries) -match "@").Trim()
    $DirectorList = ($Directors | Sort-Object -Unique) -Join ("; ")

    $TableRows = @()

    Foreach ($Notification in $NSORecords) {
        $PatchAppOwner = $Notification.name
        $ShakeoutURL = $Notification.URL
        $TableRows += Get-TmpTableRows -ShakeoutURL $ShakeoutURL -PatchAppOwner $PatchAppOwner
    }

    $MissingShakeout = @"
<!DOCTYPE html>
<html lang="en" xmlns="http://www.w3.org/1999/xhtml" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office">
<head>
<title>Shakeout Requirement Notification</title>
</head>
<body style="margin:0;padding:0;min-width:100%;background-color:#ffffff;">
<p><span style='font-size:12.0pt;font-family:"Tahoma",sans-serif'>Application Owners,</span></p>
<p><span style='font-size:10.0pt;font-family:"Tahoma",sans-serif'>This is a courtesy notification that Shakeout
Resource information has not yet been provided/is missing for $MissingCount application owners listed below.
Please click on your name and provide Shakeout Resource(s) for your applications.  Provide primary and secondary
resource contact information.  Upon completion of shakeout, please provide the Shakeout Completion Status.
If you are receiving this, but your applications are Tier 1, you are receiving it because your application(s) is/are
not identified as Tier 1 in uCMDB.
&nbsp
</span></p>
<style>
table,th,td
</style>
<table align="center" width="50%" border="1" cellpadding="0" cellspacing="0" style="min-width: 50%;" role="presentation">
<tr>
<th align="center">Shakeout Info.  (Click name)</th>
</tr>
$tablerows
</table>
&nbsp
&nbsp
&nbsp
Thank you,
$nbsp
$nbsp
Server Security Patch Deployments Team
</body>
</html>
"@

    $mail = $ol.CreateItem(0)
    $mail.HTMLBody = $MissingShakeout
    if ($Type -eq "Resource") {
        $mail.Subject = "***PLEASE READ*** $MissingCount Application shakeout resources missing for $($month) $($Year) patch cycle.  Please provide ASAP"
    }
    elseif ($Type -eq "Completion") {
        $mail.Subject = "***PLEASE READ*** Shakeout completion status not provided.  Please complete ASAP"
    }
    $mail.SentOnBehalfOfName = 'Server Security Patch Deployments Team<ServerSecurityPatchDeployments@Aetna.com>'
    $mail.To = $AddressList
    $mail.CC = if ($Escalate) { $DirectorList + "; ServerSecurityPatchDeployments@AETNA.com" }else { "ServerSecurityPatchDeployments@AETNA.com" }
    $notificationRecipients = $mail.Recipients
    $resolved = $notificationRecipients.ResolveAll()
    $mail.Save()
}
function Get-TmpTableRows {
    Param($ShakeOutUrl, $PatchAppOwner)
    $TableRow = @"
    <li><a href="$ShakeOutUrl"><span style='mso-bookmark:_MailOriginal'>$PatchAppOwner</a></li>
"@
    Return($TableRow)
}
function Get-SharePointOnlineListItem {
    <#
.SYNOPSIS
Get SharePoint Online List Items

.DESCRIPTION
Get SharePoint list items from SharePoint Online using PnP/CSOM.

.PARAMETER PatchScheduleType
A description of the PatchScheduleType parameter.

.PARAMETER Month
Month of schedule

.PARAMETER Year
A description of the Year parameter.

.PARAMETER MaxRows
A description of the MaxRows parameter.

.EXAMPLE
PS C:\> Get-SharePointOnlineListItem

.NOTES
Additional information about the function.
#>
    [CmdletBinding()]
    Param
    (
        [ValidateSet('ProductionSchedule', 'NonProductionSchedule', 'ShakeoutSchedule')]
        [String]$PatchScheduleType,
        [ValidateSet('January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December')]
        [String]$Month,
        [String]$Year
    )
    if (-not (Get-Module -Name PnP.PowerShell -ListAvailable)) {
        Install-Module -Name PnP.PowerShell
        Import-Module -Name PnP.Powershell
        Write-Output "PnP.PowerShell module has been installed and imported."
    }
    else {
        Import-Module -Name PnP.Powershell
        Write-Output "PnP.PowerShell module has been imported."
    }

    Connect-PnPOnline -Url https://aetnao365.sharepoint.com/Sites/WinCompliance/Patching -UseWebLogin


    Switch ($PatchScheduleType) {
        'ShakeoutSchedule' {
            # The name of the list
            $ListTitle = "$Month $Year - Shakeout Resources and Application Schedule"
            $ListFields = $ShakeoutFields
        }
        'ProductionSchedule' {
            # The name of the list
            $ListTitle = "$Month $Year - Production Servers Schedule"
            $ListFields = $ProductionFields
        }
        'NonProductionSchedule' {
            # The name of the list
            $ListTitle = "$Month $Year - NON Production Servers Schedule"
            $ListFields = $NonProdFields
        }
    }

    $global:List = Get-PnPList -Identity $ListTitle -Includes ItemCount

    Get-PnPListItem -List $List -PageSize 5000 | Select-Object $ListFields
}
function Update-SharePointOnlineListItem {
    Param (
        [ValidateSet('ProductionSchedule', 'NonProductionSchedule', 'ShakeoutSchedule')]
        [String]$PatchScheduleType,
        [ValidateSet('January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December')]
        [String]$Month,
        [String]$Year,
        [int]$MaxRows = 5000,
        [hashtable]$Updates,
        $List,
        [int]$ItemID
    )

    $Item = Set-PnPListItem -List $List -Identity $ItemID -Values $Updates -Verbose
}
function Remove-SharePointOnlineListItem {
    Param (
        $List,
        [int]$ItemID
    )

    Remove-PnPListItem -List $List -Identity $ItemID -Verbose
}
function Format-UrlListNameValue {
    [OutputType([string])]
    Param(
        [Parameter(Mandatory = $true)][String]$ListName
    )

    $ListName = $ListName.Replace('-', '')
    $ListName = $ListName.Replace(' ', '%20')

    return $ListName

}
function Format-UrlFilterFieldName {
    [OutputType([string])]
    Param(
        [Parameter(Mandatory = $true)][String]$FieldNameToFormat,
        [Parameter(Mandatory = $false, HelpMessage = 'If we are servicing SharePoint 2016...')][Switch]$Is2016
    )

    if ($Is2016) {

        # Replacing spaces in FieldName with UTF-8 compliant characters.
        $FieldNameToFormat = $FieldNameToFormat.Replace(' ', '%255Fx0020%255F')

    }
    else {

        # Replacing spaces in FieldName with UTF-8 compliant characters.
        $FieldNameToFormat = $FieldNameToFormat.Replace(' ', '%5Fx0020%5F')

    }
    return $FieldNameToFormat

}
function New-Schedule {
    Param(
        [Parameter(Mandatory = $false, HelpMessage = "Enter the target month name i.e. 'July' where you want the results to go... Leave blank for current month...")]
        [ValidateSet('January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December')]
        [String]$Month = (Get-Culture).DateTimeFormat.GetMonthName((Get-Date -f MM))
    )

    If (-not $Month) {
        $Month = ([CultureInfo]::InvariantCulture).DateTimeFormat.GetMonthName((Get-Date).Month)
    }

    # $PatchTuesday = Get-PatchTuesday -Month $Month
    $PatchTuesday = Get-PatchTuesday

    $Month, $Year, $Days = Get-DaysToAdvance -PatchTuesday $PatchTuesday
    # $Year = $PatchTuesday.Year

    $WorkingDirectory = "C:\Patching\$($Year)\$($Month)\Schedule"
    Write-Output "Working directory is $WorkingDirectory."
    If (-not ([System.IO.Directory]::Exists($WorkingDirectory))) {
        [void][System.IO.Directory]::CreateDirectory($WorkingDirectory)
        $TemplateDir = "C:\Patching\MonthlyTemplate"
        Write-Output "Creating base folders under C:\Patching\$($Year)\$($Month)\Schedule"
        [void](RoboCopy $TemplateDir $WorkingDirectory /MIR)
    }

    $ShakeoutList = Get-FileName -InitialDirectory C:\Patching -Title "Select Permanent Shakeout List"
    $Ext = [io.path]::GetExtension($ShakeoutList)
    if ($Ext -match "xls") {
        $CSVFilePath = $ShakeoutList.Replace($Ext, '.csv')
        if (-not [io.file]::Exists($csvFilePath) -or
            ([io.file]::Exists($csvFilePath) -and 
            [datetime](Get-Item $csvFilePath).LastWriteTime -lt [datetime](Get-Item $ShakeoutList).LastWriteTime)) {
            Convert-ExcelToCSV -ExcelFilePath $ShakeoutList -CSVFilePath $csvFilePath
        }
    }

    $SOHash = [System.Collections.hashtable]::new()
    Import-Csv $CsvFilePath | ForEach-Object {
        $SOHash.Add($_.Application.ToUpper(), $_)
    }

    $uCMDBDir = "C:\Patching\$($Year)\$($Month)\uCMDB"
    if (-not ([System.IO.Directory]::Exists($uCMDBDir))) {
        $host.ui.WriteVerboseLine("Create output folder")
        [void][System.IO.Directory]::CreateDirectory($uCMDBDir)  # If folder does not exist, it is created.  Including full path.
    }

    $MoreInput = Read-HostContinue -PromptTitle "Additional Input?" -PromptQuestion "Are there additional servers to import to the process? ('Y/N')"
    if ($MoreInput) {
        $NewServerInputPath = $(Get-FileName -Title "Select additional server input file..." -InitialDirectory "C:\Patching\Import" -FilePreference 'CSV' )
        $NewServerInput = Import-Csv $($NewServerInputPath)
    }

    $GetUCMDBFiles = Read-HostContinue -PromptTitle "Get uCMDB files" -PromptQuestion "Download and convert uCMDB files?"
    If ($GetUCMDBFiles) {
        Set-Location $uCMDBDir
        Remove-Item *.* -Force
        Get-UCMDBFiles
        ForEach ($ExcelFilePath In (Get-ChildItem | Select-Object -ExpandProperty FullName)) {
            $Ext = [io.path]::GetExtension($ExcelfilePath)
            $CSVFilePath = $ExcelFilePath.Replace($Ext, '.csv')
            $host.ui.WriteVerboseLine("Converting $($ExcelFilePath) to CSV")
            Convert-ExcelToCSV -ExcelFilePath $ExcelFilePath -CSVFilePath $CSVFilePath
            $host.ui.WriteVerboseLine("Removing original file")
            Remove-Item $ExcelFilePath
        }
    }

    [array]$uCMDBAppObjects = Import-Csv -Path $(Get-FileName -Title 'Select source uCMDB Application Information Report xlsx...' -InitialDirectory "C:\Patching\$($Year)\$($Month)\UCMDB" -FilePreference 'CSV')
    Write-Host "Found " -NoNewline; Write-Host $uCMDBAppObjects.Count -NoNewline -ForegroundColor Green; Write-Host " records in Application Info Report..."

    [array]$uCMDBServerObjects = Import-Csv -Path $(Get-FileName -Title 'Select source uCMDB Server Information Report xlsx...' -InitialDirectory "C:\Patching\$($Year)\$($Month)\uCMDB" -FilePreference 'CSV')
    Write-Host "Found " -NoNewline; Write-Host $uCMDBServerObjects.Count -NoNewline -ForegroundColor Green; Write-Host " records in Server Info Report..."

    $PrevMonth = ([cultureinfo]::InvariantCulture).DateTimeFormat.GetMonthName(($PatchTuesday.AddMonths(-1)).Month)
    $PrevYear = $PatchTuesday.AddMonths(-1).Year
    [String]$PreviousMonthMasterListPath = $(Get-FileName -Title "Select previous month's Master Server List..." -InitialDirectory "C:\Patching\$($PrevYear)\$($PrevMonth)\Schedule" -FilePreference 'CSV' )
    [array]$PreviousMonthMasterList = Import-Csv -Path $PreviousMonthMasterListPath

    [String]$SPFilePath = $(Get-FileName -Title "Select Server Patch Export" -InitialDirectory "$env:userprofile\Downloads" -FilePreference 'XLS')
    If ([io.path]::GetExtension($SPFilePath) -match "xls") {
        Write-Output "Converting Server Patch extract to .CSV"

        $ExcelFilePath = $SPFilePath
        $Ext = [io.path]::GetExtension($ExcelFilePath)
        $CSVFilePath = $ExcelFilePath.Replace($Ext, '.csv')
        Convert-ExceltoCSV -ExcelFilePath $ExcelFilePath -CSVFilePath $CSVFilePath
        $SPFilePath = $CSVFilePath
    }
    Else {
        $CSVFilePath = $SPFilePath
    }
    [array]$SPFile = Import-Csv $CSVFilePath
    Write-Output "Removing 'amp;' from '&amp;' in application names"
    $SPFile | Where-Object { $_.Application -match "&" } | ForEach-Object { $_.Application = $_.Application.Replace('amp;', '') }
    $FileName = [io.path]::GetFileName($CSVFilePath)
    Write-Output "Select location to save file"
    Start-Sleep -Seconds 2
    $PickFolder = New-Object -TypeName System.Windows.Forms.SaveFileDialog
    $PickFolder.Filename = $FileName
    $PickFolder.InitialDirectory = "$WorkingDirectory\ServerPatchInfo"

    $PickFolder.ShowDialog() | Out-Null
    $SaveDir = $([io.path]::GetDirectoryName($PickFolder.Filename))
    $SPFile | Export-Csv -Path "$($SaveDir)\$($FileName)" -NoTypeInformation
    $UnScheduled = $SPFile | Where-Object { $_.'Patch Days' -eq 'Unscheduled' }
    If ($UnScheduled.count -gt 0) {
        $UnScheduled | Export-Csv -Path "$($SaveDir)\UnScheduled.csv" -NoTypeInformation
        Write-Output "Unscheduled items written to $($SaveDir)\Unscheduled.csv"
        $host.ui.writeErrorLine("There are $($Unscheduled.count) unscheduled server(s) in ServerPatch")
        If (-not (Read-HostContinue -PromptQuestion "Do you wish to continue")) {
            Exit
        }
    }
    $x = 0
    $SPFiltered = $SPFile | Where-Object { $_.'Patch Days' -match "\d" }
    $SPFiltered | Where-Object { $_.'Patch Days' -eq "CV1" } | ForEach-Object { $_.'Patch Days' = 1 }
    $Total = $SPFiltered.count
    $interval = [int]($Total / 1000)

    $SPRecs = $SPFiltered | ForEach-Object {
        $x++
        if ($x % $Interval -eq 0) {
            Write-Progress -Activity "Building split ServerPatch File" -Status "$($x) of $($Total)" -PercentComplete ($x / $Total * 100) -Id 0
        }
        $Apps = $_.Application.Split(",", [System.StringSplitOptions]::RemoveEmptyEntries).Trim().ToUpper() | Sort-Object -Unique
        ForEach ($App In $Apps) {
            [pscustomobject] @{
                DomainServer       = $_.DomainServer
                Application        = $App.Replace("N/A", "")
                'App Env'          = If ($_.'App Env' -eq "N/A") { "" }else { $_.'App Env'.split(",", [System.StringSplitOptions]::RemoveEmptyEntries).Trim().ToUpper() -Join (',') | Sort-Object -Unique }
                'S/N'              = $_.'S/N'
                'OS'               = $_.OS
                'Patch Days'       = $_.'Patch Days'.split("/", [System.StringSplitOptions]::RemoveEmptyEntries)[0].trim()
                'Patch Times'      = $_.'Patch Times'
                'Server Time Zone' = $_.'Server Time Zone'
            }
        }
    }
    $SPRecs | Export-Csv -Path "$($SaveDir)\ServerPatch_AppsSplit.csv" -NoTypeInformation

    Write-Host "Switching location to working directory $WorkingDirectory..."
    Set-Location -Path $WorkingDirectory

    Write-Host "Script Starttime: $(Get-Date -Format G)" -ForegroundColor Yellow

    # Initializing Counters for Write-Progress
    [int]$x = 0
    [int]$y = 0

    # Filtering the uCMCB App Info Report, we only want Windows Server objects...
    Write-Host "Filtering uCMDB Application Info Report on Windows Servers only..."

    $psAppInfoWindowsServers = $uCMDBAppObjects | 
    Where-Object { ($PSItem.$uCMDBAppInfoOS -notlike '*2003*' -and `
                $PSItem.$uCMDBAppInfoOS -notlike '*2000*' -and `
                $PSItem.$uCMDBAppInfoOS -notlike 'Windows 7*' -and `
                $PSItem.$uCMDBAppInfoOS -notlike 'Windows 8*') -and `
        ($PSItem.$uCMDBAppInfoOS -like 'Windows*' -or `
                $PSItem.$uCMDBAppInfoOS -like '*2008*' -or `
                $PSItem.$uCMDBAppInfoOS -like '*2012*') }

    # Formatting the uCMDB App Info report data, adding required field names...
    Write-Host "Found " -NoNewline; Write-Host $psAppInfoWindowsServers.Count -NoNewline -ForegroundColor Green; Write-Host " Windows Servers in uCMDB Application Info Report..."

    # We now need to add the servers where the uCMDB App Info OS field is blank, but the platform field shows them as Windows.
    Write-Host "Searching through uCMDB Application Info Report for Platform identified Windows Servers, please wait..."
    $psAppInfoPlatformWindowsServers = $uCMDBAppObjects | Where-Object( { $PSItem.$uCMDBAppInfoPlatform -like 'Windows' -and $PSItem.$uCMDBAppInfoOS -eq '' })
    Write-Host "Found " -NoNewline; Write-Host $psAppInfoPlatformWindowsServers.Count -NoNewline -ForegroundColor Green; Write-Host " 'Platform' identified Windows Servers in uCMDB Application Info Report..."

    Write-Host "Adding the " -NoNewline; Write-Host $psAppInfoPlatformWindowsServers.Count -NoNewline -ForegroundColor Green; Write-Host " 'Platform' identified Windows Servers to existing array of Windows Servers found in uCMDB Application Info Report..."
    $psAppInfoWindowsServers += $psAppInfoPlatformWindowsServers
    Write-Host "Updated count is now  " -NoNewline; Write-Host $psAppInfoWindowsServers.Count -NoNewline -ForegroundColor Green; Write-Host " of Windows Servers from the uCMDB Application Info Report..."

    #region Selecting unique Windows Servers from uCMDB Application Info Report
    Write-Host "Selecting unique Windows Servers from uCMDB Application Info Report, please wait..."
    $tempAppInfoSortedServers = $psAppInfoWindowsServers | Sort-Object -Property $fn_ServerName, $fn_IPAddress, $fn_AppName, $fn_AppEnvironment, $fn_DRTier -Unique
    Write-Host "Found " -NoNewline; Write-Host $tempAppInfoSortedServers.Count -NoNewline -ForegroundColor Green; Write-Host " unique Windows Servers in uCMDB Application Info Report..."

    Write-Host "Formatting uCMDB Application Info Report Data and retrieving email addresses, please wait..."

    #Setting Counter for Write-Progress
    $psAppInfoWindowsServerCount = $tempAppInfoSortedServers.Count
    $Interval = [int]($psAppInfoWindowsServerCount / 100)
    $psAppInfoWindowsServersUnique = $tempAppInfoSortedServers | ForEach-Object {
        $x++
        if ($x % $interval -eq 0) {
            Write-Progress -Id 0 -Activity "Formatting uCMDB Application Info Report Data and retrieving email addresses, please wait..." -Status "Processing $x of $psAppInfoWindowsServerCount, $(($x/$psAppInfoWindowsServerCount).ToString("P")) complete..." `
                -PercentComplete ($x / $psAppInfoWindowsServerCount * 100)
        }

        $TmpEmail = $(Get-EmailAddress -firstName $PSItem.$fn_ITManagerFirst -lastName $PSItem.$fn_ITManagerLast)
        # $ITMEmail = $($TmpEmail+$(Get-AdditionalEmailAddress -EmailAddress $TmpEmail))
        $Additional = if ($TmpEmail) { Get-AdditionalEmailAddress -EmailAddress $TmpEmail }else { $null }
        $ITMEmail = "$($TmpEmail+$Additional)".trim()
        $PSItem
    } |
    Select-Object -Property @{Name = $fn_ServerName; Expression = { $PSItem.$fn_ServerName.ToUpper() } },
    @{Name = $fn_IPAddress; Expression = { $PSItem.$fn_IPAddress } },
    @{Name = $fn_OperatingSystem; Expression = { $PSItem.$uCMDBAppInfoOS } },
    @{Name = $fn_AppName; Expression = { $PSItem.$fn_AppName } },
    @{Name = $fn_AppEnvironment; Expression = { $PSItem.$fn_AppEnvironment } },
    @{Name = $fn_PrimarySOP; Expression = { $NotSet } },
    @{Name = $fn_SpecialSOP; Expression = { $NotSet } },
    @{Name = $fn_SOP; Expression = { $NotSet } },
    @{Name = $fn_ClusterName; Expression = { $NotSet } },
    @{Name = $fn_DRTier; Expression = { $PSItem.$fn_DRTier } },
    @{Name = $fn_IsTier1; Expression = { $NotSet } },
    @{Name = $fn_PrimaryShakeoutEmail; Expression = { $NotSet } },
    @{Name = $fn_PrimaryShakeoutPhone; Expression = { $NotSet } },
    @{Name = $fn_SecondaryShakeoutEmail; Expression = { $NotSet } },
    @{Name = $fn_SecondaryShakeoutPhone; Expression = { $NotSet } },
    @{Name = $fn_ShakeoutCompletionStatus; Expression = { $NotSet } },
    @{Name = $fn_PatchDay; Expression = { $NotSet } },
    @{Name = $fn_PatchWindow; Expression = { $NotSet } },
    @{Name = $fn_PatchDate; Expression = { $NotSet } },
    @{Name = $fn_PatchTime; Expression = { $NotSet } },
    @{Name = $fn_TimeZone; Expression = { $NotSet } },
    @{Name = $fn_StdTimeZone; Expression = { $NotSet } },
    @{Name = $fn_ITManagerFirst; Expression = { $PSItem.$fn_ITManagerFirst } },
    @{Name = $fn_ITManagerLast; Expression = { $PSItem.$fn_ITManagerLast } },
    @{Name = $fn_ITManagerEmail; Expression = { $ITMEmail } },
    @{Name = $fn_ITDirectorFirst; Expression = { $PSItem.$fn_ITDirectorFirst } },
    @{Name = $fn_ITDirectorLast; Expression = { $PSItem.$fn_ITDirectorLast } },
    @{Name = $fn_ITDirectorEmail; Expression = { (Get-EmailAddress -firstName $PSItem.$fn_ITDirectorFirst -lastName $PSItem.$fn_ITDirectorLast) } },
    @{Name = $fn_BusDirectorFirst; Expression = { $PSItem.$fn_BusDirectorFirst } },
    @{Name = $fn_BusDirectorLast; Expression = { $PSItem.$fn_BusDirectorLast } },
    @{Name = $fn_BusDirectorEmail; Expression = { (Get-EmailAddress -firstName $PSItem.$fn_BusDirectorFirst -lastName $PSItem.$fn_BusDirectorLast) } },
    @{Name = $fn_VPPOCFirst; Expression = { $PSItem.$fn_VPPOCFirst } },
    @{Name = $fn_VPPOCLast; Expression = { $PSItem.$fn_VPPOCLast } },
    @{Name = $fn_VPPOCEmail; Expression = { (Get-EmailAddress -firstName $PSItem.$fn_VPPOCFirst -lastName $PSItem.$fn_VPPOCLast) } },
    @{Name = $fn_ITVPFirst; Expression = { $PSItem.$fn_ITVPFirst } },
    @{Name = $fn_ITVPLast; Expression = { $PSItem.$fn_ITVPLast } },
    @{Name = $fn_ITVPEmail; Expression = { (Get-EmailAddress -firstName $PSItem.$fn_ITVPFirst -lastName $PSItem.$fn_ITVPLast) } },
    @{Name = $fn_PatchScheduleStatus; Expression = { 'Current' } },
    @{Name = $fn_Comments1; Expression = { $NotSet } },
    @{Name = $fn_Comments2; Expression = { $NotSet } },
    @{Name = $fn_Decommissioned; Expression = { 'False' } },
    @{Name = $fn_RecordStatus; Expression = { 'New' } },
    @{Name = $fn_Retarget; Expression = { 'False' } }

    #endregion

    # Filtering the uCMDB Server report, we only want Windows Servers...
    Write-Host "Filtering uCMDB Server Report on Windows Non-Decommissioned Servers..."
    $psServerInfoWindowsServers = $uCMDBServerObjects | Where-Object( { (
                $PSItem.$uCMDBServerOS -notlike '*2003*' -and `
                    $PSItem.$uCMDBServerOS -notlike '*2000*' -and `
                    $PSItem.$uCMDBServerOS -notlike 'Windows 7*' -and `
                    $PSItem.$uCMDBServerOS -notlike 'Windows 8*') -and `
            ($PSItem.$uCMDBServerOS -like 'Windows*' -or `
                    $PSItem.$uCMDBServerOS -like '*2008*' -or `
                    $PSItem.$uCMDBServerOS -like '*2012*') -and `
                $PSItem.$uCMDBDecomStatus -ne 1 })

    Write-Host "Found " -NoNewline; Write-Host "$($psServerInfoWindowsServers.Count)" -NoNewline -ForegroundColor Green; Write-Host " Active Windows Servers uCMDB Server Report..."

    # We now need to add the servers where the Server Info OS field is blank, but the platform field shows them as Windows.
    Write-Host "Searching through uCMDB Server Info Report for Platform identified Windows Servers, please wait..."
    $psServerInfoPlatformWindowsServers = $uCMDBServerObjects | Where-Object( { $PSItem.$uCMDBServerPlatform -like 'Windows' -and $PSItem.$uCMDBServerOS -eq '' -and $PSItem.$uCMDBDecomStatus -ne 1 })
    Write-Host "Found " -NoNewline; Write-Host $psServerInfoPlatformWindowsServers.Count -NoNewline -ForegroundColor Green; Write-Host " 'Platform' identified Windows Servers in uCMDB Server Info Report..."

    Write-Host "Adding the " -NoNewline; Write-Host $psServerInfoPlatformWindowsServers.Count -NoNewline -ForegroundColor Green; Write-Host " Platform identified Windows Servers to existing array of Windows Servers found in uCMDB Server Info Report..."
    $psServerInfoWindowsServers += $psServerInfoPlatformWindowsServers
    Write-Host "Updated count is now  " -NoNewline; Write-Host $psServerInfoWindowsServers.Count -NoNewline -ForegroundColor Green; Write-Host " of Windows Servers from the uCMDB Server Info Report..."

    # Trimming the uCMCB Server report down to 3 required fields...
    Write-Host "Trimming results, saving subset of fields to new array..."
    $psServerInfoWindowsServersTrimmed = $psServerInfoWindowsServers | 
    Select-Object -Property @{Name = $fn_ServerName; Expression = {
            if ($PSItem.$uCMDBServerFqdn) {
                $PSItem.$uCMDBServerFqdn.ToUpper()
            }
            else {
                $PSItem.$uCMDBServerLabel.ToUpper()
            }
        }
    }, @{Name = $fn_IPAddress; Expression = { $PSItem.$uCMDBServerIPAddress } }

    Write-Host "Filtering out duplicates and sorting data from uCMDB Server Report..."
    $psServerInfoWindowsServersUnique = $psServerInfoWindowsServersTrimmed | Sort-Object -Property $fn_ServerName, $fn_IPAddress -Unique
    Write-Host "Found " -NoNewline; Write-Host $psServerInfoWindowsServersUnique.Count -NoNewline -ForegroundColor Green; Write-Host " 'Active, and Unique' Windows Servers in uCMDB Server Report..."

    $DifferenceObjects = Compare-Object -ReferenceObject $psServerInfoWindowsServersUnique -DifferenceObject $psAppInfoWindowsServersUnique -Property $fn_ServerName, $fn_IPAddress
    $MissingServers = $DifferenceObjects | Where-Object( { $PSItem.SideIndicator -eq '<=' })

    Write-Host "Found " -NoNewline; Write-Host $MissingServers.Count -NoNewline -ForegroundColor Green; Write-Host " uCMDB Server Info objects missing in ApplicationInfo Server Objects report..."

    #region Formatting Uncorrelated uCMDB ServerInfo objects
    Write-Host "Formatting Uncorrelated uCMDB ServerInfo objects, preparing for pipeline..."
    $MissingServersFormatted = $MissingServers | Select-Object -Property @{Name = $fn_ServerName; Expression = { $PSItem.$fn_ServerName } },
    @{Name = $fn_IPAddress; Expression = { $PSItem.$fn_IPAddress } },
    @{Name = $fn_OperatingSystem; Expression = { $PSItem.$uCMDBServerOS } },
    @{Name = $fn_AppName; Expression = { $NotSet } },
    @{Name = $fn_AppEnvironment; Expression = { $NotSet } },
    @{Name = $fn_PrimarySOP; Expression = { $NotSet } },
    @{Name = $fn_SpecialSOP; Expression = { $NotSet } },
    @{Name = $fn_SOP; Expression = { $NotSet } },
    @{Name = $fn_ClusterName; Expression = { $NotSet } },
    @{Name = $fn_DRTier; Expression = { $NotSet } },
    @{Name = $fn_IsTier1; Expression = { $NotSet } },
    @{Name = $fn_PrimaryShakeoutEmail; Expression = { $NotSet } },
    @{Name = $fn_PrimaryShakeoutPhone; Expression = { $NotSet } },
    @{Name = $fn_SecondaryShakeoutEmail; Expression = { $NotSet } },
    @{Name = $fn_SecondaryShakeoutPhone; Expression = { $NotSet } },
    @{Name = $fn_ShakeoutCompletionStatus; Expression = { $NotSet } },
    @{Name = $fn_PatchDay; Expression = { $NotSet } },
    @{Name = $fn_PatchWindow; Expression = { $NotSet } },
    @{Name = $fn_PatchDate; Expression = { $NotSet } },
    @{Name = $fn_ITManagerFirst; Expression = { $NotSet } },
    @{Name = $fn_ITManagerLast; Expression = { $NotSet } },
    @{Name = $fn_ITManagerEmail; Expression = { $NotSet } },
    @{Name = $fn_ITDirectorFirst; Expression = { $NotSet } },
    @{Name = $fn_ITDirectorLast; Expression = { $NotSet } },
    @{Name = $fn_ITDirectorEmail; Expression = { $NotSet } },
    @{Name = $fn_BusDirectorFirst; Expression = { $NotSet } },
    @{Name = $fn_BusDirectorLast; Expression = { $NotSet } },
    @{Name = $fn_BusDirectorEmail; Expression = { $NotSet } },
    @{Name = $fn_VPPOCFirst; Expression = { $NotSet } },
    @{Name = $fn_VPPOCLast; Expression = { $NotSet } },
    @{Name = $fn_VPPOCEmail; Expression = { $NotSet } },
    @{Name = $fn_ITVPFirst; Expression = { $NotSet } },
    @{Name = $fn_ITVPLast; Expression = { $NotSet } },
    @{Name = $fn_ITVPEmail; Expression = { $NotSet } },
    @{Name = $fn_PatchScheduleStatus; Expression = { 'Current' } },
    @{Name = $fn_Comments1; Expression = { $NotSet } },
    @{Name = $fn_Comments2; Expression = { $NotSet } },
    @{Name = $fn_Decommissioned; Expression = { 'False' } },
    @{Name = $fn_RecordStatus; Expression = { 'New' } },
    @{Name = $fn_Retarget; Expression = { 'False' } }

    #endregion

    Write-Host "Merging newly filtered Application Info and Server Info results into temporary master server list..." -ForegroundColor Green
    $TempMasterServerList = $psAppInfoWindowsServersUnique + $MissingServersFormatted | Sort-Object -Property $fn_ServerName

    Write-Host "Starting merge with previous month's MasterServerlist..."

    Write-Host "Importing $PreviousMonthMasterListPath..."
    $PreviousMonthMasterList = Import-Csv -Path $PreviousMonthMasterListPath | ForEach-Object { $PSItem.$fn_ServerName = $PSItem.$fn_ServerName.ToUpper(); $PSItem }
    Write-Host "Last month's Production list '$PreviousMonthMasterListPath' imported successfully..." -ForegroundColor Green

    if ($NewServerInput) {
        Write-Host "Adding additional new servers to temporary master server list"
        $TempMasterServerList += $NewServerInput
        Write-Host "Added " -NoNewline; Write-Host "$($NewServerInput.count)" -NoNewline -ForegroundColor Green; Write-Host " additional servers added to Temporary master server list."
    }

    Write-Host "Finding carry-over servers/servers added to inventory after generation of previous schedule.  This may take a while" -ForegroundColor Yellow
    $NameList = $TempMasterServerList.$fn_ServerName.ToUpper()
    # $CarryOver = $PreviousMonthMasterList | Show-Progress -Activity "Getting carryover servers" | Where-Object {$_.$fn_ServerName.ToUpper() -notin ($TempMasterServerList.$fn_ServerName).ToUpper()}

    $CarryOver = $PreviousMonthMasterList | Show-Progress -Activity "Getting carryover servers" | Where-Object { $_.$fn_ServerName.ToUpper() -notin $NameList }
    $Decommed = ($uCMDBServerObjects | Where-Object { [bool][int]$_.$uCMDBDecomStatus }).$uCMDBServerLabel
    $CarryOver = $CarryOver | Where-Object { $_.$fn_ServerName.split("\")[-1] -notin $Decommed }
    $CarryOver | Export-Csv -Path "$WorkingDirectory\CarryOver.csv" -NoTypeInformation
    $TempMasterServerList += $CarryOver

    Write-Host "Setting up temp arrays..."
    $NewMasterServerList = @()

    [int]$TempMasterCount = $TempMasterServerList.Count
    [int]$PreviousMonthCount = $PreviousMonthMasterList.Count

    # Resetting for next loop
    $x = 0

    Write-Host "Merging patching metadata from previous month's MasterServerlist, to new MasterServerList..."
    $UnCorReport = @()
    $NewServers = @()
    $Interval = [int]($TempMasterCount / 100)
    foreach ($TempMasterRecord in $TempMasterServerList) {
        $ItemFound = $false
        $uncorrelated = $false
        $x++
        if ( $x % $Interval -eq 0) {
            Write-Progress -Id 0 -Activity "Iterating through new MasterServerList: updating patch schedule and management contact info..." -Status "Processing $x of $TempMasterCount, $(($x/$TempMasterCount).ToString('P')) complete..." `
                -PercentComplete ($x / $TempMasterCount * 100)
        } 
        foreach ($PreviousRecord in $PreviousMonthMasterList) {
            if ($TempMasterRecord.$fn_ServerName -eq $PreviousRecord.$fn_ServerName) {
                $ItemFound = $true

                # If we are here, we have an exact application/server match. We pull in all Blade Logic patching schedules from last month...
                if ($TempMasterRecord.$fn_AppName -eq $PreviousRecord.$fn_AppName) {

                    $TempMasterRecord.$fn_PatchDay = $PreviousRecord.$fn_PatchDay.Trim()
                    $TempMasterRecord.$fn_PatchWindow = $PreviousRecord.$fn_PatchWindow.Trim()

                    # Handling the patch date, moving forward either 28 or 35 days depending on number of weeks between patch Tuesdays'.
                    # Skipping if it's a string and not a date in string format.
                    if ($PreviousRecord.$fn_PatchDate.Contains("/")) {
                        # It's a date, we advance the date according to number of days required for next patching cycle...
                        $TempMasterRecord.$fn_PatchDate = (Get-Date -Date $PreviousRecord.$fn_PatchDate).AddDays($NumberOfDaysToAdvanceSchedule).ToString('MM/dd/yyyy')
                    }
                    else {
                        # It's a string, we just trim and continue...
                        $TempMasterRecord.$fn_PatchDate = $PreviousRecord.$fn_PatchDate.Trim()
                    }

                    $TempMasterRecord.$fn_ClusterName = $PreviousRecord.$fn_ClusterName.Trim()
                    $TempMasterRecord.$fn_PrimarySOP = $PreviousRecord.$fn_PrimarySOP.Trim()
                    $TempMasterRecord.$fn_SpecialSOP = $PreviousRecord.$fn_SpecialSOP.Trim()
                    $TempMasterRecord.$fn_SOP = $PreviousRecord.$fn_SOP.Trim()
                    $TempMasterRecord.$fn_IsTier1 = $PreviousRecord.$fn_IsTier1.Trim()
                    $TempMasterRecord.$fn_DRTier = $PreviousRecord.$fn_DRTier.Trim()
                    $TempMasterRecord.$fn_RecordStatus = 'Existing'
                }

                # if we are here, we have a partial match, server name matches, but last month's application name is blank. 
                # In this case we update only the scheduling information from last month.
                elseif (($TempMasterRecord.$fn_AppName -ne '') -and ($PreviousRecord.$fn_AppName -eq '')) {

                    $TempMasterRecord.$fn_PatchDay = $PreviousRecord.$fn_PatchDay.Trim()
                    $TempMasterRecord.$fn_PatchWindow = $PreviousRecord.$fn_PatchWindow.Trim()

                    # Handling the patch date, moving forward either 28 or 35 days depending on number of weeks between patch Tuesdays'.
                    # Skipping if it's a string and not a date in string format.
                    if ($PreviousRecord.$fn_PatchDate.Contains("/")) {
                        # It's a date, we advance the date according to number of days required for next patching cycle...
                        $TempMasterRecord.$fn_PatchDate = (Get-Date -Date $PreviousRecord.$fn_PatchDate).AddDays($NumberOfDaysToAdvanceSchedule).ToString('MM/dd/yyyy')
                    }
                    else {
                        # It's a string, we just trim and continue...
                        $TempMasterRecord.$fn_PatchDate = $PreviousRecord.$fn_PatchDate.Trim()
                    }

                    $TempMasterRecord.$fn_ClusterName = $PreviousRecord.$fn_ClusterName.Trim()
                    $TempMasterRecord.$fn_SOP = $PreviousRecord.$fn_SOP.Trim()
                    $TempMasterRecord.$fn_IsTier1 = $PreviousRecord.$fn_IsTier1.Trim()
                    $TempMasterRecord.$fn_DRTier = $PreviousRecord.$fn_DRTier.Trim()
                    $TempMasterRecord.$fn_RecordStatus = 'Existing'

                }

                # if we are here, we have a second partial match, server name matches but incoming application name is blank. 
                # In this case we update the application name, mgmt contact info, and scheduling information from last month.
                elseif (($TempMasterRecord.$fn_AppName -eq '') -and ($PreviousRecord.$fn_AppName -ne '')) {

                    if ($TempMasterRecord.$fn_ServerName -in $MissingServers.$fn_ServerName) {
                        $TempMasterRecord.$fn_ServerName
                        $uncorrelated = $true
                        $TempMasterRecord.$fn_AppName = $NotSet
                        $TempMasterRecord.$fn_AppEnvironment = $NotSet
                        $TempMasterRecord.$fn_ITManagerFirst = $NotSet
                        $TempMasterRecord.$fn_ITManagerLast = $NotSet
                        $TempMasterRecord.$fn_ITManagerEmail = $NotSet
                        $TempMasterRecord.$fn_ITDirectorFirst = $NotSet
                        $TempMasterRecord.$fn_ITDirectorLast = $NotSet
                        $TempMasterRecord.$fn_ITDirectorEmail = $NotSet
                        $TempMasterRecord.$fn_BusDirectorFirst = $NotSet
                        $TempMasterRecord.$fn_BusDirectorLast = $NotSet
                        $TempMasterRecord.$fn_BusDirectorEmail = $NotSet
                        $TempMasterRecord.$fn_VPPOCFirst = $NotSet
                        $TempMasterRecord.$fn_VPPOCLast = $NotSet
                        $TempMasterRecord.$fn_VPPOCEmail = $NotSet
                        $TempMasterRecord.$fn_ITVPFirst = $NotSet
                        $TempMasterRecord.$fn_ITVPLast = $NotSet
                        $TempMasterRecord.$fn_ITVPEmail = $NotSet
                    }
                    else {
                        # Setting application name and environment and from last month onto this record
                        $TempMasterRecord.$fn_AppName = $PreviousRecord.$fn_AppName.Trim()
                        $TempMasterRecord.$fn_AppEnvironment = $PreviousRecord.$fn_AppEnvironment.Trim()

                        # Setting mgmt contact info from last month as the incoming record has no data from this server record.
                        $TempMasterRecord.$fn_ITManagerFirst = $PreviousRecord.$fn_ITManagerFirst.Trim()
                        $TempMasterRecord.$fn_ITManagerLast = $PreviousRecord.$fn_ITManagerLast.Trim()
                        $TempMasterRecord.$fn_ITManagerEmail = $PreviousRecord.$fn_ITManagerEmail.Trim()
                        $TempMasterRecord.$fn_ITDirectorFirst = $PreviousRecord.$fn_ITDirectorFirst.Trim()
                        $TempMasterRecord.$fn_ITDirectorLast = $PreviousRecord.$fn_ITDirectorLast.Trim()
                        $TempMasterRecord.$fn_ITDirectorEmail = $PreviousRecord.$fn_ITDirectorEmail.Trim()
                        $TempMasterRecord.$fn_BusDirectorFirst = $PreviousRecord.$fn_BusDirectorFirst.Trim()
                        $TempMasterRecord.$fn_BusDirectorLast = $PreviousRecord.$fn_BusDirectorLast.Trim()
                        $TempMasterRecord.$fn_BusDirectorEmail = $PreviousRecord.$fn_BusDirectorEmail.Trim()
                        $TempMasterRecord.$fn_VPPOCFirst = $PreviousRecord.$fn_VPPOCFirst.Trim()
                        $TempMasterRecord.$fn_VPPOCLast = $PreviousRecord.$fn_VPPOCLast.Trim()
                        $TempMasterRecord.$fn_VPPOCEmail = $PreviousRecord.$fn_VPPOCEmail.Trim()
                        $TempMasterRecord.$fn_ITVPFirst = $PreviousRecord.$fn_ITVPFirst.Trim()
                        $TempMasterRecord.$fn_ITVPLast = $PreviousRecord.$fn_ITVPLast.Trim()
                        $TempMasterRecord.$fn_ITVPEmail = $PreviousRecord.$fn_ITVPEmail.Trim()
                    }

                    # Trimming any space from beginning or end of patch schedule
                    $TempMasterRecord.$fn_PatchDay = $PreviousRecord.$fn_PatchDay.Trim()
                    $TempMasterRecord.$fn_PatchWindow = $PreviousRecord.$fn_PatchWindow.Trim()

                    # Handling the patch date, moving forward either 28 or 35 days depending on number of weeks between patch Tuesdays'.
                    # Skipping if it's a string and not a date in string format.
                    if ($PreviousRecord.$fn_PatchDate.Contains("/")) {
                        # It's a date, we advance the date according to number of days required for next patching cycle...
                        $TempMasterRecord.$fn_PatchDate = (Get-Date -Date $PreviousRecord.$fn_PatchDate).AddDays($NumberOfDaysToAdvanceSchedule).ToString('MM/dd/yyyy')
                    }
                    else {
                        # It's a string, we just trim and continue...
                        $TempMasterRecord.$fn_PatchDate = $PreviousRecord.$fn_PatchDate.Trim()
                    }

                    $TempMasterRecord.$fn_ClusterName = $PreviousRecord.$fn_ClusterName.Trim()
                    $TempMasterRecord.$fn_SOP = $PreviousRecord.$fn_SOP.Trim()
                    $TempMasterRecord.$fn_IsTier1 = $PreviousRecord.$fn_IsTier1.Trim()
                    $TempMasterRecord.$fn_DRTier = $PreviousRecord.$fn_DRTier.Trim()
                    $TempMasterRecord.$fn_RecordStatus = 'Existing'
                }
                else {

                    $TempMasterRecord.$fn_PatchDay = $PreviousRecord.$fn_PatchDay.Trim()
                    $TempMasterRecord.$fn_PatchWindow = $PreviousRecord.$fn_PatchWindow.Trim()

                    # Handling the patch date, moving forward either 28 or 35 days depending on number of weeks between patch Tuesdays'.
                    # Skipping if it's a string and not a date in string format.
                    if ($PreviousRecord.$fn_PatchDate.Contains("/")) {
                        # It's a date, we advance the date according to number of days required for next patching cycle...
                        $TempMasterRecord.$fn_PatchDate = (Get-Date -Date $PreviousRecord.$fn_PatchDate).AddDays($NumberOfDaysToAdvanceSchedule).ToString('MM/dd/yyyy')
                    }
                    else {
                        # It's a string, we just trim and continue...
                        $TempMasterRecord.$fn_PatchDate = $PreviousRecord.$fn_PatchDate.Trim()
                    }

                    $TempMasterRecord.$fn_ClusterName = $PreviousRecord.$fn_ClusterName.Trim()
                    $TempMasterRecord.$fn_SOP = $PreviousRecord.$fn_SOP.Trim()
                    $TempMasterRecord.$fn_IsTier1 = $PreviousRecord.$fn_IsTier1.Trim()
                    $TempMasterRecord.$fn_DRTier = $PreviousRecord.$fn_DRTier.Trim()
                    $TempMasterRecord.$fn_RecordStatus = 'Existing'

                }
            }
        } #End foreach ($PreviousRecord in $PreviousMonthMasterList)

        $NewMasterServerList += $TempMasterRecord
        if ($uncorrelated) {
            $UnCorReport += $TempMasterRecord
        }
        if (-not $ItemFound) {
            $NewServers += $TempMasterRecord
        }
    } #End foreach ($TempMasterRecord in $TempMasterServerList)

    Write-Host "New Master List created successfully in memory..." -ForegroundColor Green

    # Reseting from above use...
    $x = 0
    $y = 0

    [int]$NewMasterCount = $NewMasterServerList.Count

    Write-Host "Retrieving number of new uCMDB server records count..."

    [int]$NewRecordCount = ($NewMasterServerList | Where-Object( { $PSItem.$fn_RecordStatus -eq 'New' })).Count
    #[int]$NewRecordCount = $NewMasterServerList.Count

    Write-Host "Found " -NoNewline; Write-Host "$NewRecordCount" -NoNewline -ForegroundColor Green; Write-Host " new server records this month, retrieving email address info..."

    Write-Host "Retrieving email addresses for new server entries found in uCMDB..."

    $MasterServerSchedulePath = "$WorkingDirectory\$Month $Year - MasterServerSchedule.csv"
    $UnCorReportPath = "$WorkingDirectory\$Month $Year - Windows Uncorrelated.csv"
    $NewServerPath = "$WorkingDirectory\$Month $Year - New Servers.csv"

    Write-Output "Determining problematic Primary SOP items"

    $SOPItems = $NewMasterServerList | Where-Object { $_.'Primary SOP' -eq "TRUE" }
    $SOPNoApp = $SOPItems | Where-Object { [system.string]::IsNullOrWhiteSpace($_.$fn_AppName) }
    If ($SOPNoApp) {
        Write-Host "The below servers are assigned Primary SOP but have no application association" -ForegroundColor Red
        Write-Host "Will set Application name to 'Uncorrelated with SOP'" -ForegroundColor Red
        $SOPNoApp.$fn_ServerName
        $NewMasterServerList | Where-Object { $_ -in $SOPNoApp } | ForEach-Object {
            $_.$fn_AppName = "Uncorrelated with SOP"
        }
    }
    # $MultiSOP = $NewMasterServerList | Where-Object {$_.'Primary SOP' -eq "TRUE"}|Group-Object 'Server Name' -NoElement|Where-Object {$_.Count -gt 1}
    $MultiSOP = $SOPItems | Group-Object 'Server Name' -NoElement | Where-Object { $_.Count -gt 1 }
    if ($MultiSOP) {
        $MultiSOP | Export-Csv -Path $WorkingDirectory\MultiSOP.csv -NoTypeInformation
    }

    Write-Host "Creating $Month $Year - MasterServerSchedule at: $MasterServerSchedulePath..."

    $UnCorReport | Where-Object { $_.$fn_OperatingSystem -ne "" } | Sort-Object -Property $fn_ServerName | Export-Csv -Path $UnCorReportPath -NoTypeInformation
    $x = 0
    $Interval = [int]($NewRecordCount / 100)
    $NewMasterServerList | Where-Object( { $PSItem.$fn_RecordStatus -eq 'New' }) | ForEach-Object {
        $x++
        if ($x % $Interval -eq 0) {
            Write-Progress -Id 0 -Activity "Retrieving email addresses for new server entries found in uCMDB..." -Status "Processing $x of $NewRecordCount, $(($x/$NewRecordCount).ToString('P')) complete..." `
                -PercentComplete ($x / $NewRecordCount * 100) #-CurrentOperation 'Looking for new server entries...'
        }
        $TmpEmail = $(Get-EmailAddress -firstName $PSItem.$fn_ITManagerFirst -lastName $PSItem.$fn_ITManagerLast)
        $ITMEmail = $($TmpEmail + $(Get-AdditionalEmailAddress -EmailAddress $TmpEmail))

        $PSItem.$fn_ITManagerEmail = $($ITMEmail)
        $PSItem.$fn_ITDirectorEmail = Get-EmailAddress -firstName $PSItem.$fn_ITDirectorFirst -lastName $PSItem.$fn_ITDirectorLast
        $PSItem.$fn_BusDirectorEmail = Get-EmailAddress -firstName $PSItem.$fn_BusDirectorFirst -lastName $PSItem.$fn_BusDirectorLast
        $PSItem.$fn_VPPOCEmail = Get-EmailAddress -firstName $PSItem.$fn_VPPOCFirst -lastName $PSItem.$fn_VPPOCLast
        $PSItem.$fn_ITVPEmail = Get-EmailAddress -firstName $PSItem.$fn_ITVPFirst -lastName $PSItem.$fn_ITVPLast
    }
    $ServerAddresses = Import-Csv C:\Patching\EmailAddressMappings\ServerAddresses.csv
    Foreach ($Address in $ServerAddresses) {
        $NewMasterServerList | Where-Object { $_.$fn_ServerName -eq $Address.$fn_ServerName } | ForEach-Object {
            $_.$fn_ITManagerEmail = $($Address.$fn_ITManagerEmail)
        }
    }
    $NewMasterServerList | Sort-Object -Property $fn_ServerName | Export-Csv -Path $MasterServerSchedulePath -NoTypeInformation 

    $MultiSOP2 = $NewMasterServerList | Where-Object { $_.'Primary SOP' -eq "TRUE" } | Group-Object 'Server Name' -NoElement | Where-Object { $_.Count -gt 1 } 
    if ($MultiSOP2) {
        $MultiSOP2 | Export-Csv $WorkingDirectory\MultiSOP2.csv -NoTypeInformation
    }
    # Creating Master Server List and other output ...

    $UnCorReport | Where-Object { $_.$fn_OperatingSystem -eq "" } | Sort-Object -Property $fn_ServerName | Export-Csv -Path "$WorkingDirectory\$Month $Year (Windows) - Unknown OS Uncorrelated.csv" -NoTypeInformation

    $NewServers | Export-Csv -Path $NewServerPath -NoTypeInformation
    $TZHash = @{}
    Import-Csv $WorkingDirectory\ServerPatchInfo\TimeZoneList.csv | ForEach-Object {
        $TZHash[$_.Name] = $_.ID
    }
    $SchedHash = @{}
    Write-Output "Creating hash table for schedule lookups"
    $NewMasterServerList | ForEach-Object {
        If ([System.String]::IsNullOrWhiteSpace($_.$fn_AppName)) {
            $SchedHash[$_.$fn_ServerName.Split('.')[0]] = $_
        }
        Else {
            $SchedHash["$($_.$fn_ServerName.Split('.')[0])-$($_.$fn_Appname)"] = $_
        }
    }
    $Schedule = @()

    $i = 0
    $total = $sprecs.count
    $Interval = [int]($Total / 100)
    Write-Output "Building merged schedule..."
    ForEach ($SPRec In $SPRecs) {
        $i++
        $pct = ($i / $Total)
        if ($i % $Interval -eq 0) {
            Write-Progress -Id 0 -Activity "Working" -Status "$($i) of $($total): $($pct.tostring('P')) complete" -PercentComplete ($pct * 100)
        }
        If ([system.string]::IsNullOrWhiteSpace($SPRec.Application)) {
            $Lookup = $SPRec.DomainServer.split("\")[-1]
        }
        Else {
            $Lookup = "$($SPRec.DomainServer.Split("\")[-1])-$($SPRec.Application)"
        }

        $InvRec = $SchedHash[$Lookup]
        $Schedule += Build-Record -SPRec $SPRec -InvRec $InvRec -TZHash $TZHash
    }
    Write-Output "Looking up permanent shakeout resource..."
    Foreach ($Record in $Schedule) {
        $SORec = Try { $SOHash[$Record.$fn_AppName] }Catch { $null }
        if ($SORec) {
            $Record.$fn_PrimaryShakeoutEmail = $SORec.$fn_PrimaryShakeoutEmail
            $Record.$fn_PrimaryShakeoutPhone = $SORec.$fn_PrimaryShakeoutPhone
            $Record.$fn_SecondaryShakeoutEmail = $SORec.$fn_SecondaryShakeoutEmail
            $Record.$fn_SecondaryShakeoutPhone = $SORec.$fn_SecondaryShakeoutPhone
        }
    }
    $OutputSchedule = "$WorkingDirectory\$($Month)-$($Year)_Schedule.csv"
    $Schedule | Export-Csv -Path $OutputSchedule -NoTypeInformation
    if (Read-HostContinue -PromptTitle "Generate SharePoint list input?" -PromptQuestion "Do you wish to create the SharePoint list input files at this time?") {
        Create-MonthlySharePointLists -Month $Month -Year $Year -MasterServerListPath $OutputSchedule
    }

    Write-Host "Script EndTime: $(Get-Date -Format G)" -ForegroundColor Yellow
}
Function Convert-ToExcelTable {
    Param($SourceCsv)
    $SourceCsv = (Get-Item $SourceCsv).FullName
    $OutputXlsx = $SourceCsv.Replace('csv', 'xlsx')
    $Excel = New-Object -ComObject Excel.Application
    [void]$Excel.workbooks.Open($SourceCsv)
    Write-Output "File Opened"
    $Excel.DisplayAlerts = $false
    $ListObject = $Excel.ActiveSheet.ListObjects.Add([Microsoft.Office.Interop.Excel.XlListObjectSourceType]::xlSrcRange, $Excel.ActiveCell.CurrentRegion, $null , [Microsoft.Office.Interop.Excel.XlYesNoGuess]::xlYes)
    $ListObject.Name = "TableData"
    $ListObject.TableStyle = "TableStyleMedium2"
    Write-Output "Table Created"
    $Excel.ActiveWorkbook.SaveAs($OutputXlsx, 51)
    Write-Output "Excel file created."
    $Excel.close
    $Excel.Quit()
}
#endregion Functions
#endregion Global Objects

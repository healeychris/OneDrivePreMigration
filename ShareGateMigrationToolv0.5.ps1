<#	
    .NOTES
    ===========================================================================
    Created with: 	VS
    Created on:   	05/10/2021
    Created by:   	Chris Healey
    Organization: 	
    Filename:     	SharegateMigrationTool.ps1
    Project path:   https://
    ===========================================================================
    .DESCRIPTION
    This script is used to perform migrations using the 
    Sharegate desktop tool and powershell modules.
#>


Clear-Host

######################## Do not change below #####################################################################################################


$AdminSiteURL                            = "https://cnainsurance-admin.sharepoint.com"                          # Admin URL for Tenant
$SharePointHomeURL                       = "https://cnainsurance.sharepoint.com"                                # Home base URL for SharePoint Online
$ListName                                = "OneDriveProject"                                                    # Project site in SharePoint Online to store details  
$host.ui.RawUI.WindowTitle               = "OneDrive Migration Script"                                          # Transaction Log Folder Name
#$CSVDataFile                             = '.\OneDriveMigrationList.csv'                                        # OneDrive Input File to check for users
$MigrationResultsFile                    = ".\ExportResults_$((get-date).ToString('yyyyMMdd_HHmm')).csv"        # Results file from Output
$ProcessedUsers                          = 0
$ErrorCount                              = 0
$BatchesFolder   		                 = 'Batches' 
$RequireConnectMicrosoftSharePointPNP    = $true
$RecordIntoListSharePointPerMigration    = $true                                                                # Record each migration into the $ListName Site 



# FUNCTION - Ask for Batch Name
function GetBatchname () {
    # Get the batch name to search for
        $Global:BatchName = Read-Host -Prompt "Enter the Batch Name to Start Migrations"
        Write-Host `n

        # Exist if no value is entered
        if ($BatchName -eq ""){WriteTransactionsLogs -Task "No batch name was entered, closing application"  -Result Error -ErrorMessage none -ShowScreenMessage true -ScreenMessageColour RED -IncludeSysError false
        Write-Host `n
        TerminateScript
    }
    $host.ui.RawUI.WindowTitle  = "OneDrive PreCheck Script - Working on $BatchName"
        
}



# FUNCTION - Check  Batch Data from Folder
function GetBatchData () {

    # Check if directory exists and has migration user list in
    if (Test-Path ".\$BatchesFolder\$BatchName"){WriteTransactionsLogs -Task "Batch Folder Found"  -Result Error -ErrorMessage none -ShowScreenMessage true -ScreenMessageColour GREEN -IncludeSysError false}

    # Run Functions
    CheckCSVDataFile
    ImportCSVData

    # Create Stats from Migration data
    $UserCSVCount = $Global:OneDriveUsers.count

    WriteTransactionsLogs -Task "$BatchName has $UserCSVCount Users that will be migrated$"  -Result Error -ErrorMessage none -ShowScreenMessage true -ScreenMessageColour GREEN -IncludeSysError false

}

# FUNCTION - Create Output Files
function CreateOutPutFiles () {


    # Results file
	if (! (Test-Path "$MigrationResultsFile")) {
		New-Item -Path . -Name "$MigrationResultsFile" -ItemType "file" -Value 'Displayname,SamaccountName,OneDriveURL,HomeDirectoryPath,Result,SessionId,SiteObjectsCopied,ItemsCopied,Successes,Warnings,Errors,TaskName' -Force | Out-Null
		Add-Content -path "$MigrationResultsFile" -value ""
	}

}


# FUNCTION - Update List in SharePoint 
If ($true -eq $RecordIntoListSharePointPerMigration) {
function AddRecordIntoListSharePoint () {

    #Connect to PNP Online
    Connect-PnPOnline -Url $SharePointHomeURL -Credentials (Get-Credential)
 
    # Get Date
    $DateNow = Get-Date -f g

    #Add List Item - Internal Names of the columns: Value
    Add-PnPListItem -List $ListName -Values @{"Title" = "$Samaccountname";
                                              "Displayname" = "$Displayname";
                                              "HomeDirectory" = "$User.DIRECTORY";
                                              "OneDriveURL" = "$User.ONEDRIVEURL";
                                              "Result" = "$Result";
                                              "SessionId" = "$SessionId";
                                              "SiteObjectsCopied" = "$SiteObjectsCopied";
                                              "ItemsCopied" = "$ItemsCopied";
                                              "Successes" = "$Successes";
                                              "Warnings" = "$Warnings";
                                              "Errors" = "$Errors";
                                              "TaskName" = "$SamaccountName - $Displayname - Migration";
                                              "Batch" = "$BatachName";
                                              "MigrationDate" = "$DateNow"
                                              "DetailedMigrationFile" = "$SharePointHomeURL/sites/$ListName/DetailedMigrationData/$BatchName/$SessionId.xlsx"}

    }
}



# FUNCTION - Terminate script
function TerminateScript () {

    Write-Host `n
    Write-Host 'This script has closed due to the error above' -foregroundcolor white -backgroundcolor RED
    Write-Host `n
    Write-Host "Press any key to end..."

    # Create pause like wait...
    $x = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    Exit
}


# FUNCTION -  Check ShareGate Module is installed
function CheckShareGateModule () {

    WriteTransactionsLogs -Task "Checking Microsoft Sharegate Module"  -Result Information -ErrorMessage none -ShowScreenMessage true -ScreenMessageColour GREEN -IncludeSysError false   
    if (Get-Module -ListAvailable -Name Sharegate ) {
        WriteTransactionsLogs -Task "Found Microsoft Sharegate Module" -Result Information -ErrorMessage none -ShowScreenMessage true -ScreenMessageColour GREEN -IncludeSysError false 	
    } else {
        WriteTransactionsLogs -Task "Failed to locate Microsoft Sharegate Module, it needs to be installed" -Result Error -ErrorMessage "Sharegate Module not installed" -ShowScreenMessage true -ScreenMessageColour RED -IncludeSysError false
    TerminateScript	
    }
}


# FUNCTION -  Check SharePointPNP Module is installed
function CheckSharePNPModule () {

    WriteTransactionsLogs -Task "Checking Microsoft SharePointPNP Module"  -Result Information -ErrorMessage none -ShowScreenMessage true -ScreenMessageColour GREEN -IncludeSysError false   
    if (Get-Module -ListAvailable -Name PnP.PowerShell ) {
        WriteTransactionsLogs -Task "Found Microsoft SharePointPNP Module" -Result Information -ErrorMessage none -ShowScreenMessage true -ScreenMessageColour GREEN -IncludeSysError false 	
    } else {
        WriteTransactionsLogs -Task "Failed to locate Microsoft SharePointPNP Module, it needs to be installed" -Result Error -ErrorMessage "SharePointPNP Module not installed" -ShowScreenMessage true -ScreenMessageColour RED -IncludeSysError false
    TerminateScript	
    }
}


# FUNCTION - Get TXT Data from File
function ImportModuleShareGate () {

    WriteTransactionsLogs -Task "Importing ShareGate Module"   -Result Information none -ShowScreenMessage true -ScreenMessageColour GREEN -IncludeSysError false    
    try {Import-Module Sharegate -ea stop
        WriteTransactionsLogs -Task "Sharegate Module loaded"   -Result Information -ErrorMessage none -ShowScreenMessage true -ScreenMessageColour GREEN -IncludeSysError false
    } 
    catch {WriteTransactionsLogs -Task "Error loading Sharegate Module" -Result Error -ErrorMessage "An error happened importing the Module :" -ShowScreenMessage true -ScreenMessageColour RED -IncludeSysError true
        TerminateScript
    }

}


# FUNCTION -  Check input file exists, terminate script if not
function CheckCSVDataFile () {

    WriteTransactionsLogs -Task "Checking CSV File............"    -Result Information -ErrorMessage none -ShowScreenMessage true -ScreenMessageColour GREEN -IncludeSysError false 
    if (! (Test-Path ".\$BatchesFolder\$BatchName\$CSVDataFile")) {
	    WriteTransactionsLogs -Task "CSV File Check" -Result Error -ErrorMessage "CSV File Not found in expected location" -ShowScreenMessage true -ScreenMessageColour RED -IncludeSysError false
        TerminateScript
    } else {
        WriteTransactionsLogs -Task "Found CSV File..............."    -Result Information -ErrorMessage none -ShowScreenMessage true -ScreenMessageColour GREEN -IncludeSysError false 
    }
}


# FUNCTION - Get TXT Data from File
function ImportCSVData () {

    # Find Latest Results CSV file in Batch\Report directory
    $CSVDataFile = Get-ChildItem | Sort-Object -Descending -Property LastAccessTime | Select-Object -First 1 | Select-Object -ExpandProperty Name


    WriteTransactionsLogs -Task "Importing Data file................"   -Result Information none -ShowScreenMessage true -ScreenMessageColour GREEN -IncludeSysError false    
    try {$Global:OneDriveUsers = Import-Csv ".\$BatchesFolder\$BatchName\$CSVDataFile" -Delimiter "," -ea stop
        WriteTransactionsLogs -Task "Loaded Users Data"   -Result Information -ErrorMessage none -ShowScreenMessage true -ScreenMessageColour GREEN -IncludeSysError false
    } 
    catch {WriteTransactionsLogs -Task "Error loading Users data File" -Result Error -ErrorMessage "An error happened importing the data file, Please Check File" -ShowScreenMessage true -ScreenMessageColour RED -IncludeSysError false
        TerminateScript
    }

}

# FUNCTION - Connect to SharePoint Online PNP
function ConnectMicrosoftSharePointPNP () {

    # Check Connection to SharePointPNP or Connect if not already
    if ($RequireConnectMicrosoftSharePointPNP -eq $true) {

        try {
         try { Get-PnPTenant -ea stop | Out-Null;  WriteTransactionsLogs -Task "Existing SharePointPNP Connection Found" -Result Information -ErrorMessage none -ShowScreenMessage true -ScreenMessageColour GREEN -IncludeSysError false}
         catch {
                WriteTransactionsLogs -Task "Not Connected to SharePointPNP" -Result Information -ErrorMessage none -ShowScreenMessage true -ScreenMessageColour GREEN -IncludeSysError false
                if ($MFALoginRequired -eq $True){$Global:SharePointlogon = Connect-PnPOnline -Url $SharePointHomeURL -UseWebLogin  -ErrorAction Stop | Out-Null}
                if ($MFALoginRequired -eq $False){Connect-PnPOnline -Url $SharePointHomeURL -credential $credentials  -ErrorAction Stop | Out-Null}
            }
        }  
        Catch {
            WriteTransactionsLogs -Task "Unable to Connect to Microsoft SharePointPNP" -Result Error -ErrorMessage "Connect Error" -ShowScreenMessage true -ScreenMessageColour RED -IncludeSysError true
	        TerminateScript
        }
    }
}


# FUNCTION - ProcessMigration
function ProcessMigration () {

    WriteTransactionsLogs -Task "Starting Migration..." -Result Information -ErrorMessage "none" -ShowScreenMessage true -ScreenMessageColour GREEN -IncludeSysError false
    Set-Variable dstSite, dstList

    foreach ($User in $OneDriveUsers) {
        Clear-Variable dstSite
        Clear-Variable dstList

        $Displayname    = $User.Displayname 
        $SamaccountName = $User.SamaccountName 

        #Incremental Mode
        $copysettings = New-CopySettings -OnContentItemExists IncrementalUpdate

        try {
            WriteTransactionsLogs -Task "Processing....$Displayname" -Result Information -ErrorMessage "none" -ShowScreenMessage true -ScreenMessageColour GREEN -IncludeSysError false
            $host.ui.RawUI.WindowTitle = "OneDrive Migration Script - Processing $Displayname" 
            $dstSite = Connect-Site -Url $User.ONEDRIVEURL -Browser
            $dstList = Get-List -Name Documents -Site $dstSite 
            $MigrationData = Import-Document -SourceFolder $User.HomeDirectory -DestinationList $dstList -taskname "$SamaccountName - $Displayname - Migration" -WarningAction:SilentlyContinue -CopySettings $copysettings
            $ProcessedUsers ++
        }
        Catch {
            WriteTransactionsLogs -Task "Failed Processing...$Displayname" -Result ERROR -ErrorMessage "Error:" -ShowScreenMessage true -ScreenMessageColour GREEN -IncludeSysError true
            $errorCount ++
        }
        
        #Remove-SiteCollectionAdministrator -Site $dstSite

                
        $Result                     =   $MigrationData.Result
        $SessionId                  =   $MigrationData.Sessionid
        $SiteObjectsCopied          =   $MigrationData.SiteObjectsCopied
        $ItemsCopied                =   $MigrationData.ItemsCopied
        $Successes                  =   $MigrationData.Successes
        $Warnings                   =   $MigrationData.Warnings
        $Errors                     =   $MigrationData.Errors
        $DetailedMigrationFile      =   "$SharePointHomeURL/sites/OneDriveProject/DetailedMigrationData/$BatchName/$SessionId.xlsx"

        WriteTransactionsLogs -Task "Finished Processing $Displayname | $Result" -Result Information -ErrorMessage "none" -ShowScreenMessage true -ScreenMessageColour GREEN -IncludeSysError false

        # Create array to store and write output data
        $ReportFile = [pscustomobject][ordered]@{}
        $ReportFile | Add-Member -MemberType NoteProperty -Name SamaccountName -Value $Samaccountname -Force
        $ReportFile | Add-Member -MemberType NoteProperty -Name Displayname -Value $Displayname -Force
        $ReportFile | Add-Member -MemberType NoteProperty -Name HomeDirectory -Value $User.DIRECTORY -Force
        $ReportFile | Add-Member -MemberType NoteProperty -Name OneDriveURL -Value $User.ONEDRIVEURL -Force
        $ReportFile | Add-Member -MemberType NoteProperty -Name Result -Value $Result -Force
        $ReportFile | Add-Member -MemberType NoteProperty -Name SessionId -Value $SessionId -Force
        $ReportFile | Add-Member -MemberType NoteProperty -Name SiteObjectsCopied -Value $SiteObjectsCopied -Force
        $ReportFile | Add-Member -MemberType NoteProperty -Name ItemsCopied -Value $ItemsCopied -Force
        $ReportFile | Add-Member -MemberType NoteProperty -Name Successes -Value $Successes -Force
        $ReportFile | Add-Member -MemberType NoteProperty -Name Warnings -Value $Warnings -Force
        $ReportFile | Add-Member -MemberType NoteProperty -Name Errors -Value $Errors -Force
        $ReportFile | Add-Member -MemberType NoteProperty -Name TaskName -Value "$SamaccountName - $Displayname - Migration" -Force

        
        # Export out to file
        $ReportFile | Export-Csv -path ".\$BatchesFolder\$BatchName\$MigrationResultsFile" -Append -NoTypeInformation


    }
    # Update Window title for status
    $host.ui.RawUI.WindowTitle = "OneDrive Migration Script - Processed:$ProcessedUsers / Failed:$ErrorCount"
}



# FUNCTION - WriteLogMessages and Screen Message
function WriteTransactionsLogs  {

    #WriteTransactionsLogs -Task 'Creating folder' -Result information  -ScreenMessage true -ShowScreenMessage true exit #Writes to file and screen, basic display
          
    #WriteTransactionsLogs -Task task -Result Error -ErrorMessage errormessage -ShowScreenMessage true -ScreenMessageColour red -IncludeSysError true #Writes to file and screen and system "error[0]" is recorded
         
    #WriteTransactionsLogs -Task task -Result Error -ErrorMessage errormessage -ShowScreenMessage true -ScreenMessageColour red -IncludeSysError false  #Writes to file and screen but no system "error[0]" is recorded
         


    [CmdletBinding()]
    param(
        [Parameter(Mandatory,ValueFromPipelineByPropertyName)]
        [ValidateNotNullOrEmpty()]
        [string]$Task,
 
        [Parameter(Mandatory,ValueFromPipelineByPropertyName)]
        [ValidateSet('Information','Warning','Error','Completed','Processing')]
        [string]$Result,
 
        [Parameter(ValueFromPipelineByPropertyName)]
        [AllowNull()]
        [string]$ErrorMessage,
    
        [Parameter(Mandatory,ValueFromPipelineByPropertyName)]
        [ValidateSet('True','False')]
        [string]$ShowScreenMessage,
 
        [Parameter(Mandatory,ValueFromPipelineByPropertyName)]
        [string]$ScreenMessageColour,
 
        [Parameter(ValueFromPipelineByPropertyName)]
        [string]$IncludeSysError
 
 )
 
    process {
 
        # Stores Variables
        $LogsFolder      		     = 'Logs'
 
        # Date
        $DateNow = Get-Date -f g
        
        # Error Message
        $SysErrorMessage = $error[0].Exception.message
 
 
        # Check of log files exist for this session
        If ($null -eq $Global:TransactionLog) {$Global:TransactionLog = ".\TransactionLog_$((get-date).ToString('yyyyMMdd_HHmm')).csv"} # Used to capture a running event of the process, error and transactions 
 
        
        # Create Directory Structure
        if (! (Test-Path ".\$LogsFolder")) {new-item -path .\ -name ".\$LogsFolder" -type directory | out-null}
 
 
 
        $TransactionLogScreen = [pscustomobject][ordered]@{}
        $TransactionLogScreen | Add-Member -MemberType NoteProperty -Name "Date"-Value $DateNow 
        $TransactionLogScreen | Add-Member -MemberType NoteProperty -Name "Task" -Value $Task
        $TransactionLogScreen | Add-Member -MemberType NoteProperty -Name "Result" -Value $Result
        $TransactionLogScreen | Add-Member -MemberType NoteProperty -Name "Error" -Value $ErrorMessage
        $TransactionLogScreen | Add-Member -MemberType NoteProperty -Name "SystemError" -Value $SysErrorMessage
        
       
        # Output to screen
       
        if  ($Result -match "Information|Warning" -and $ShowScreenMessage -eq "$true"){
 
        Write-host $TransactionLogScreen.Date  -NoNewline -ForegroundColor GREEN
        Write-host " | " -NoNewline
        Write-Host $TransactionLogScreen.Task  -NoNewline
        Write-host " | " -NoNewline
        Write-host $TransactionLogScreen.Result -ForegroundColor $ScreenMessageColour 
        }
 
       if  ($Result -eq "Error" -and $ShowScreenMessage -eq "$true" -and $IncludeSysError -eq "$false"){
       Write-host $TransactionLogScreen.Date  -NoNewline -ForegroundColor GREEN
       Write-host " | " -NoNewline
       Write-Host $TransactionLogScreen.Task  -NoNewline
       Write-host " | " -NoNewline
       Write-host $TransactionLogScreen.Result -ForegroundColor $ScreenMessageColour -NoNewline 
       Write-host " | " -NoNewline
       Write-Host $ErrorMessage  -ForegroundColor $ScreenMessageColour
       }
 
       if  ($Result -eq "Error" -and $ShowScreenMessage -eq "$true" -and $IncludeSysError -eq "$true"){
       Write-host $TransactionLogScreen.Date  -NoNewline -ForegroundColor GREEN
       Write-host " | " -NoNewline
       Write-Host $TransactionLogScreen.Task  -NoNewline
       Write-host " | " -NoNewline
       Write-host $TransactionLogScreen.Result -ForegroundColor $ScreenMessageColour -NoNewline 
       Write-host " | " -NoNewline
       Write-Host $ErrorMessage -NoNewline -ForegroundColor $ScreenMessageColour
       if (!$SysErrorMessage -eq $null) {Write-Host " | " -NoNewline}
       Write-Host $SysErrorMessage -ForegroundColor $ScreenMessageColour
       Write-Host
       }
 
     
 
 
        $TransactionLogFile = [pscustomobject][ordered]@{}
        $TransactionLogFile | Add-Member -MemberType NoteProperty -Name "Date"-Value "$datenow"
        $TransactionLogFile | Add-Member -MemberType NoteProperty -Name "Task"-Value "$task"
        $TransactionLogFile | Add-Member -MemberType NoteProperty -Name "Result"-Value "$result"
        $TransactionLogFile | Add-Member -MemberType NoteProperty -Name "Error"-Value "$ErrorMessage"
        $TransactionLogFile | Add-Member -MemberType NoteProperty -Name "SystemError"-Value "$SysErrorMessage"
 
        $TransactionLogFile | Export-Csv -Path ".\$LogsFolder\$TransactionLog" -Append -NoTypeInformation
 
 
        # Clear Error Messages
        $error.clear()
    }   
 
}






#### Funcion run order
GetBatchname
GetBatchData
CheckSharePointModule
CheckSharePNPModule 
ImportModuleShareGate
ConnectMicrosoftSharePointPNP
#CheckCSVDataFile
#ImportCSVData
ProcessMigration
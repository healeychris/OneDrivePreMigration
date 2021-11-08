<#	
    .NOTES
    ===========================================================================
    Created with: 	VS
    Created on:   	05/10/2021
    Created by:   	Chris Healey
    Organization: 	
    Filename:     	SharegateMigrationTool.ps1
    Project path:   https://
    Version :       0.5
    ===========================================================================
    .DESCRIPTION
    This script is used to perform migrations using the 
    Sharegate desktop tool and powershell modules.
#>


Clear-Host

######################## Do not change below #####################################################################################################


$AdminSiteURL                            = "https://cnainsurance-admin.sharepoint.com"                          # Admin URL for Tenant
$SharePointHomeURL                       = "https://cnainsurance.sharepoint.com"                                # Home base URL for SharePoint Online
$SharePointProjectSiteName               = "https://cnainsurance.sharepoint.com/sites/OneDriveProject"          # Project site in SharePoint Online to store details  
$PostMigrationListName                   = "PostMigration"                                                      # Project site in SharePoint Online to store details
$host.ui.RawUI.WindowTitle               = "OneDrive Migration Script"                                          # Transaction Log Folder Name
$PreMigrationReports                     = 'PreMigrationReports'                                                # Reports Directory
$ShareGateReportsExport                  = 'ShareGateReports'                                                   # Sharegate Directory for export reports
$MigrationResultsFile                    = ".\ExportResults_$((get-date).ToString('yyyyMMdd_HHmm')).csv"        # Results file from Output
$ProcessedUsers                          = 0                                                                    # Default var of completed users
$ErrorCount                              = 0                                                                    # Default var of failed users
$BatchesFolder   		                 = 'Batches'                                                            # Batch folder directory
$RequireConnectMicrosoftSharePointPNP    = $true                                                                # Should data be writting to SharePointPNP
$RecordIntoListSharePointPerMigration    = $true                                                                # Record each migration into the $ListName Site 
$RequireConnectShareGate                 = $true
#$Reports                                 = 'Reports'                                                           # Reports Directory
$MFALoginRequired                        = $false                                                               # Do connections require MFA
$LogsFolder                              = 'Log'                                                                # Log Files to store details in 
$MigrationReports                        = 'MigrationReports'  
$ExcludeListFile                         ='.\ExcludeUsers.txt'                                                  # List of Users to exclude by SAmaccountName

                                                    





function DisplayExtendedInfo () {

    # Display to notify the operator before running
    Clear-Host
    Write-Host 
    Write-Host 
    Write-Host  '----------------------------------------------------------------------------------'	
	Write-Host  '            OneDrive Migration Tool for ShareGate'                                    -ForegroundColor Green
	Write-Host  '----------------------------------------------------------------------------------'
    Write-Host
    Write-Host
    Write-Host  '  This script is to be used wth the Pre-migration tool to generate the batch of   '   -ForegroundColor Yellow
    Write-Host  '  Users that have been verified and and can be moved to 365.                      '   -ForegroundColor Yellow 
    Write-Host  '                                                                                  '          
    Write-Host  '                                                                                  '   -ForegroundColor YELLOW
    Write-Host  '----------------------------------------------------------------------------------'
    Write-Host 
}



function AskForCreds () {


    # Asking for creds if they don't exist and the embedded acount is not used

    if ($MFALoginRequired -eq $False -and !$credentials) {

    WriteTransactionsLogs -Task "Asking for Credentrials as MFA is not required"  -Result Error -ErrorMessage none -ShowScreenMessage true -ScreenMessageColour YELLOW -IncludeSysError false

        $Global:credentials = Get-Credential -Message "Enter Login details for Office 365"
        Write-host 'The login details entered are stored in this shell Windows. And will be used until the window is closed' -ForegroundColor YELLOW
    
    }

}




# FUNCTION - Ask for Batch Name
function GetBatchname () {
    # Get the batch name to search for
        $Global:BatchName = Read-Host -Prompt "Enter the Batch Name to Start Migrations"
        Write-Host `n

        #if (Test-Path ".\$BatchesFolder\$BatchName")

        # Exist if no value is entered
        if ($BatchName -eq ""){WriteTransactionsLogs -Task "No batch name was entered, closing application"  -Result Error -ErrorMessage none -ShowScreenMessage true -ScreenMessageColour RED -IncludeSysError false
        Write-Host `n
        TerminateScript
    }
    $host.ui.RawUI.WindowTitle  = "OneDrive PreCheck Script - Working on $BatchName"
        
}



# FUNCTION - Check Batch Data from Folder
function GetBatchData () {

    # Check if directory exists and has migration user list in
    if (Test-Path ".\$BatchesFolder\$BatchName"){WriteTransactionsLogs -Task "Batch Directory Found"  -Result Information -ErrorMessage none -ShowScreenMessage true -ScreenMessageColour GREEN -IncludeSysError false}

    # Run Functions
    CheckCSVDataFile
    ImportCSVData

    # Create Stats from Migration data
    $UserCSVCount = $Global:OneDriveUsers.UserPrincipalName.count

    WriteTransactionsLogs -Task "$BatchName has $UserCSVCount Users that will be migrated"  -Result Information -ErrorMessage none -ShowScreenMessage true -ScreenMessageColour YELLOW -IncludeSysError false

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

    # Get Date
    $DateNow = Get-Date

    try {
        # Add List Item - Internal Names of the columns: Value
          Add-PnPListItem -ea stop -ContentType "Item" -List $PostMigrationListName  -Values @{"Title" = "$Samaccountname";
                                              "Displayname" = "$Displayname";
                                              "HomeDirectory" = "$HomeDirectory";
                                              "OneDriveURL" = "$OneDriveURL";
                                              "Result" = "$Result";
                                              "SessionId" = "$SessionId";
                                              "SiteObjectsCopied" = "$SiteObjectsCopied";
                                              "ItemsCopied" = "$ItemsCopied";
                                              "Successes" = "$Successes";
                                              "Warnings" = "$Warnings";
                                              "Errors" = "$Errors";
                                              "TaskName" = "$SamaccountName - $Displayname - Migration";
                                              "Batch" = "$Global:BatchName";
                                              "MigrationDate" = "$DateNow";
                                              "DetailedMigrationFile" = "$SharePointHomeURL/sites/OneDriveProject/Detailsmigrationdata/$SessionId.xlsx"} | Out-Null
    
    
    WriteTransactionsLogs -Task "Added PNP record to $PostMigrationListName libary"  -Result Information -ErrorMessage none -ShowScreenMessage true -ScreenMessageColour GREEN -IncludeSysError false   

    }
    Catch {WriteTransactionsLogs -Task "Failed to add PNP record to $PostMigrationListName libary"  -Result Error -ErrorMessage Failed: -ShowScreenMessage true -ScreenMessageColour RED -IncludeSysError true}   
         }
}



# FUNCTION - Update List in SharePoint with ExportReport
If ($true -eq $RecordIntoListSharePointPerMigration) {
function AddExportReportIntoListSharePoint () {

    # Export the Report from Sharegate
    Try {Export-Report -SessionId $SessionId -Path ".\$BatchesFolder\$BatchName\$ShareGateReportsExport\$SessionId" | Out-Null
         WriteTransactionsLogs -Task "Export of Sharegate Report Completed"  -Result Information -ErrorMessage none -ShowScreenMessage true -ScreenMessageColour GREEN -IncludeSysError false
         
         # Upload of ShareGate Report into SharePoint List
         Add-PnPFile -Path ".\$BatchesFolder\$BatchName\$ShareGateReportsExport\$SessionId.xlsx" -Folder "Detailsmigrationdata" -Values @{SamaccountName="$SamaccountName"} | Out-Null
         WriteTransactionsLogs -Task "Upload of Sharegate Report Completed"  -Result Information -ErrorMessage none -ShowScreenMessage true -ScreenMessageColour GREEN -IncludeSysError false}   
    
    Catch {WriteTransactionsLogs -Task "ShareGate report Export/Upload Failed"  -Result ERROR -ErrorMessage none -ShowScreenMessage true -ScreenMessageColour RED -IncludeSysError true}
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
    if (Get-Module -ListAvailable -Name PnP.PowerShell) {
        WriteTransactionsLogs -Task "Found Microsoft SharePointPNP Module" -Result Information -ErrorMessage none -ShowScreenMessage true -ScreenMessageColour GREEN -IncludeSysError false 	
    } else {
        WriteTransactionsLogs -Task "Failed to locate Microsoft SharePointPNP Module, it needs to be installed" -Result Error -ErrorMessage "SharePointPNP Module not installed" -ShowScreenMessage true -ScreenMessageColour RED -IncludeSysError false
    TerminateScript	
    }
}


# FUNCTION - Import Module
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

    
    # Find Latest Results CSV file in Batch\Report directory
    $Global:CSVDataFile = Get-ChildItem ".\$BatchesFolder\$BatchName\$PreMigrationReports" | Sort-Object -Descending -Property LastAccessTime | Select-Object -First 1 | Select-Object -ExpandProperty Name

    WriteTransactionsLogs -Task "Checking CSV File............"    -Result Information -ErrorMessage none -ShowScreenMessage true -ScreenMessageColour GREEN -IncludeSysError false 
    if (! (Test-Path ".\$BatchesFolder\$BatchName\$PreMigrationReports\$CSVDataFile")) {
	    WriteTransactionsLogs -Task "CSV File Check" -Result Error -ErrorMessage "CSV File Not found in expected location" -ShowScreenMessage true -ScreenMessageColour RED -IncludeSysError false
        TerminateScript
    } else {
        WriteTransactionsLogs -Task "Found CSV File..............."    -Result Information -ErrorMessage none -ShowScreenMessage true -ScreenMessageColour GREEN -IncludeSysError false 
    }
}


# FUNCTION - Get CSV Data from File
function ImportCSVData () {

    WriteTransactionsLogs -Task "Importing Data file................$CSVDataFile"   -Result Information none -ShowScreenMessage true -ScreenMessageColour GREEN -IncludeSysError false    
    try {$Global:OneDriveUsers = Import-Csv ".\$BatchesFolder\$BatchName\$PreMigrationReports\$CSVDataFile" -Delimiter "," -ea stop
        WriteTransactionsLogs -Task "Loaded Users Data"   -Result Information -ErrorMessage none -ShowScreenMessage true -ScreenMessageColour GREEN -IncludeSysError true
    } 
    catch {WriteTransactionsLogs -Task "Error loading Users data File" -Result Error -ErrorMessage "An error happened importing the data file, Please Check File" -ShowScreenMessage true -ScreenMessageColour RED -IncludeSysError false
          TerminateScript
    }

}


# FUNCTION - Check Exclude User list
function CheckExcludeList () {

    WriteTransactionsLogs -Task "Checking For Exclude List File............"    -Result Information -ErrorMessage none -ShowScreenMessage true -ScreenMessageColour GREEN -IncludeSysError false 
    if (! (Test-Path $ExcludeListFile)) {
	    WriteTransactionsLogs -Task "Exclude List File Check" -Result Information -ErrorMessage "Exclude File Not found in expected location" -ShowScreenMessage true -ScreenMessageColour YELLOW -IncludeSysError false
        $ExcludeListFileNotFound = $false
    }else {
        WriteTransactionsLogs -Task "Exclude List File Check Located..........."    -Result Information -ErrorMessage none -ShowScreenMessage true -ScreenMessageColour GREEN -IncludeSysError false 
    }
}


# FUNCTION - Import Exclude User list
function ImportExcludeList () {

    if ($null -eq $ExcludeListFileNotFound){
       WriteTransactionsLogs -Task "Importing Exclude List File............"    -Result Information -ErrorMessage none -ShowScreenMessage true -ScreenMessageColour GREEN -IncludeSysError false 
        
        try {$Global:ExcludeListUsers =  Get-content $ExcludeListFile
            $ExcludeListUsersCount = $ExcludeListUsers.count
            WriteTransactionsLogs -Task "Imported Exclude List File and has $ExcludeListUsersCount Users listed!"    -Result Warning -ErrorMessage none -ShowScreenMessage true -ScreenMessageColour YELLOW -IncludeSysError false}
        
        Catch {WriteTransactionsLogs -Task "Imported Exclude List Failed, Job will Continue"    -Result Error -ErrorMessage none -ShowScreenMessage true -ScreenMessageColour RED -IncludeSysError false
              $ExcludeListFileNotFound = $false}

    }

}


# FUNCTION - Check CSV Data for Valid Users
function ValidateUserCSVData () {

    #Check for excluded users
    #Check for Not Valid users in CSV

    # Build complete Var list to pass for processing

    # Store a copy of the userlist Modified in the Batch folder for reference
    #$MigrationFileUsed = ".\$BatchesFolder\$BatchName\$TXTDataFile" + "_ValidatedList.txt"
    #Copy-Item $TXTDataFile $MigrationFileUsed


}




# FUNCTION - Connect to SharePoint Online PNP
function ConnectMicrosoftSharePointPNP () {

    # Check Connection to SharePointPNP or Connect if not already
    if ($RequireConnectMicrosoftSharePointPNP -eq $true) {

     try {
         try { Get-PnPTenant -ea stop | Out-Null;  WriteTransactionsLogs -Task "Existing SharePointPNP Connection Found" -Result Information -ErrorMessage none -ShowScreenMessage true -ScreenMessageColour GREEN -IncludeSysError false}
         catch {
                WriteTransactionsLogs -Task "Not Connected to SharePointPNP" -Result Information -ErrorMessage none -ShowScreenMessage true -ScreenMessageColour GREEN -IncludeSysError false
                if ($MFALoginRequired -eq $True){$Global:SharePointPNPlogon = Connect-PnPOnline -Url $SharePointProjectSiteName -UseWebLogin  -ErrorAction Stop | Out-Null}
                if ($MFALoginRequired -eq $False){Connect-PnPOnline -Url $SharePointHomeURL -credential $credentials  -ErrorAction Stop | Out-Null}
                WriteTransactionsLogs -Task "Connected to SharePointPNP" -Result Information -ErrorMessage "none" -ShowScreenMessage true -ScreenMessageColour GREEN -IncludeSysError false
            }
        }  
        Catch {
            WriteTransactionsLogs -Task "Unable to Connect to Microsoft SharePointPNP" -Result Error -ErrorMessage "Connect Error" -ShowScreenMessage true -ScreenMessageColour RED -IncludeSysError true
	        TerminateScript
        }
    }
}



# FUNCTION - Connect to ShareGate
function ConnectShareGate () {

    # Check Connection to ShareGate Connect if not already
    if ($RequireConnectShareGate -eq $true) {

     try {
         try {Get-List -site $Global:ShareGateLogon  -ea stop | Out-Null;  WriteTransactionsLogs -Task "Existing ShareGate Connection Found" -Result Information -ErrorMessage none -ShowScreenMessage true -ScreenMessageColour GREEN -IncludeSysError false}
         catch {
                WriteTransactionsLogs -Task "Not Connected to ShareGate" -Result Information -ErrorMessage none -ShowScreenMessage true -ScreenMessageColour GREEN -IncludeSysError false
                if ($MFALoginRequired -eq $True){$Global:ShareGatelogon = Connect-Site -Url $AdminSiteURL -Browser -DisableSSO -ErrorAction Stop}
                if ($MFALoginRequired -eq $False){Connect-Site -Url $AdminSiteURL  -credential $credentials  -ErrorAction Stop | Out-Null}
                WriteTransactionsLogs -Task "Connected to SharGate" -Result Information -ErrorMessage "none" -ShowScreenMessage true -ScreenMessageColour GREEN -IncludeSysError false
            }
        }  
        Catch {
            WriteTransactionsLogs -Task "Unable to Connect to ShareGate" -Result Error -ErrorMessage "Connect Error: " -ShowScreenMessage true -ScreenMessageColour RED -IncludeSysError true
	        TerminateScript
        }
    }
}




# FUNCTION - ProcessMigration
function ProcessMigration () {

    WriteTransactionsLogs -Task "Starting Migration..." -Result Information -ErrorMessage "none" -ShowScreenMessage true -ScreenMessageColour GREEN -IncludeSysError false
    
    # Connect to base SharPoint Site and Store creds
    #$Global:ShareGateConnection = Connect-Site -Url $AdminSiteURL -Browser -DisableSSO


    # Create Var
    Set-Variable dstSite, dstList

    foreach ($User in $OneDriveUsers) {
        Clear-Variable dstSite
        Clear-Variable dstList

        $Displayname        = $User.Displayname 
        $SamaccountName     = $User.SamaccountName 
        $HomeDirectory      = $User.HomeDirectory
        $OneDriveURL        = $User.ONEDRIVEURL
        $HomeDirectorySize  = $User.HomeDirectorySize

        #$HomeDirectorySize = "{0:n2}" -f ($HomeDirectorySize.SizeinBytes/1MB)

        $HomeDirectorySizeCalc = "{0:n2}" -f  $HomeDirectorySize/1mb 

        #Incremental Mode
        # $copysettings = New-CopySettings -OnContentItemExists IncrementalUpdate

        try {
            WriteTransactionsLogs -Task "Processing....$Displayname - Size $HomeDirectorySizeCalc in MB )" -Result Information -ErrorMessage "none" -ShowScreenMessage true -ScreenMessageColour GREEN -IncludeSysError false
            $host.ui.RawUI.WindowTitle = "OneDrive Migration Script - Processing $Displayname" 
            $dstSite = Connect-Site -Url $OneDriveURL -UseCredentialsFrom $Global:ShareGatelogon
            $dstList = Get-List -Name Documents -Site $dstSite 
            $MigrationData = Import-Document -SourceFolder $HomeDirectory -DestinationList $dstList -taskname "$SamaccountName - $Displayname - Migration" -WarningAction:SilentlyContinue #-CopySettings $copysettings
            
            # Count additional user completed
            $ProcessedUsers ++

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
           $ReportFile | Add-Member -MemberType NoteProperty -Name HomeDirectory -Value $HomeDirectory -Force
           $ReportFile | Add-Member -MemberType NoteProperty -Name OneDriveURL -Value $OneDriveURL  -Force
           $ReportFile | Add-Member -MemberType NoteProperty -Name Result -Value $Result -Force
           $ReportFile | Add-Member -MemberType NoteProperty -Name SessionId -Value $SessionId -Force
           $ReportFile | Add-Member -MemberType NoteProperty -Name SiteObjectsCopied -Value $SiteObjectsCopied -Force
           $ReportFile | Add-Member -MemberType NoteProperty -Name ItemsCopied -Value $ItemsCopied -Force
           $ReportFile | Add-Member -MemberType NoteProperty -Name Successes -Value $Successes -Force
           $ReportFile | Add-Member -MemberType NoteProperty -Name Warnings -Value $Warnings -Force
           $ReportFile | Add-Member -MemberType NoteProperty -Name Errors -Value $Errors -Force
           $ReportFile | Add-Member -MemberType NoteProperty -Name TaskName -Value "$SamaccountName - $Displayname - Migration" -Force

        
           # Export out to file
           $ReportFile | Export-Csv -path ".\$BatchesFolder\$BatchName\$MigrationReports\$MigrationResultsFile" -Append -NoTypeInformation

           # Run Add PNP record for ShrePoint Update
           AddRecordIntoListSharePoint
           AddExportReportIntoListSharePoint
           

        }
        Catch {
            WriteTransactionsLogs -Task "Failed Processing...$Displayname" -Result ERROR -ErrorMessage "Error:" -ShowScreenMessage true -ScreenMessageColour RED -IncludeSysError true
            $errorCount ++
        }
        
       
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
        #$LogsFolder      		     = 'Logs'
 
        # Date
        $DateNow = Get-Date -f g
        
        # Error Message
        $SysErrorMessage = $error[0].Exception.message
 
 
        # Check of log files exist for this session
        If ($null -eq $Global:TransactionLog) {$Global:TransactionLog = ".\TransactionLog_$((get-date).ToString('yyyyMMdd_HHmm')).csv"} # Used to capture a running event of the process, error and transactions 
 
        
        # Create Directory Structure
        #if (! (Test-Path ".\$LogsFolder")) {new-item -path .\ -name ".\$LogsFolder" -type directory | out-null}
 
 
 
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
 
        $TransactionLogFile | Export-Csv -Path ".\$BatchesFolder\$BatchName\$LogsFolder\$TransactionLog" -Append -NoTypeInformation
 
 
        # Clear Error Messages
        $error.clear()
    }   
 
}






#### Funcion run order
DisplayExtendedInfo
GetBatchname
AskForCreds
CheckSharePNPModule 
ImportModuleShareGate
GetBatchData
CheckExcludeList
ImportExcludeList
ConnectMicrosoftSharePointPNP
ConnectShareGate
ValidateUserCSVData 
ProcessMigration
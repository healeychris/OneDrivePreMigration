<#	
    .NOTES
    ===========================================================================
    Created with: 	VS
    Created on:   	27/09/2021
    Created by:   	Chris Healey
    Organization: 	
    Filename:     	OneDrivePreCheck.ps1
    Project path:   https://
    Version :       0.5
    ===========================================================================
    .DESCRIPTION
    This script is used check users before performing OneDrive Migrations. 
    - Checks the user is vaild in AD
    - Checks the user has a SharePoint Licence
    - Checks the user has a valid home directory in AD
    - Calculates the total size and files on HomeDirectory
    - Collections information on OneDrive if provisioned / size
    - Provisions OneDrive if vaild to do so based on licence 
    - Batching and reporting per folder structure
#>


Clear-Host

######################## CAN BE CHANGED #####################################################################################################

$MFALoginRequired                        = $true                                                                # Is MFA login required
$AdminSiteURL                            = "https://cnainsurance-admin.sharepoint.com"                          # Admin URL for Tenant
$CheckSharePointProvision                = $true                                                                # Check if Provisioned in SharePoint (Slow process if enabled/true)
$GetHomeDriveSize                        = $false                                                                # Calculate the size of the users HomeDirectory (Slow process if enabled/true) 
$ProvisionSharePointContainer            = $false                                                                # Provision OneDrive container 
$GCServer                                = 'xxx'                                                 # GC to be used to find user information
$DCServer                                = 'xxx'                                                         # DC to be used to find user information
$UsePSSessionActiveDirectory             = $true                                                                # Use Active Directory module from Domain Controller

# Migration Servers/Sites
$PittsburgDCList = 'Server1|Server2'
$GreenfordDCList = 'Server3|Server4'
$AuroraDCList    = 'kdenvr02|kreadr02|kcranr02|KPLNTR01|KAUSTR01|khousr01|KDGRVR01|kkansr01|KALBQR01|KCLBSR01|KDLUTR01|KSFRAR01|Klrckr01|KSYRCR01|KFHCCR01|khuntr01|KINDYR01|KELLCR01|KBOSTR01|kbre1r01|KALBYR01|KCHARR01|KGRPDR01|KCHVYR01|Kphnxr01|klosar01|KRICHR01|KPHILR01|KRCKHR01|KSEATR01|KPIT1R01|KLOUIR01|KOKCYR01|kyorkr01|KMILWR01|KMELVR01|kwashr01|KSLKCR01|KWHP1R01|klosnr01|kportr01|KSANDR01|kw2wpr01|KQNCYR01|KTAMPR01|KW2CGR01|KMTRIR01|kslccr01|KNVBRR01|KW2MNR01|kbirmr01|kminnr01|KPRSPR01|KW2VCR01|Kw2tor01|KSACRR01|KDALPR01|kch1r500|kmater02|kch1r400|kch1r100|kw7srr01|kch1r801|kch1r101|kch1r300|kch1r600'

# Fix local issues
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

######################## Do not change below #####################################################################################################

$host.ui.RawUI.WindowTitle               = "OneDrive PreCheck Script"                                           # Transaction Log Folder Name
$TXTDataFile                             = '.\OneDriveUserList.txt'                                             # OneDrive Input File to check for users
$ExportResultsFile                       = ".\ExportResults_$((get-date).ToString('yyyyMMdd_HHmm')).csv"        # Results file from Output
$MissingUsers                            = ".\NotFoundUsers.txt"                                                # Users excluded due to not been found in AD 
$ReportFolder                            = 'Reports'                                                            # Folder containing the output reports
$MigrationReports                        = 'MigrationReports'
$ShareGateReportsExport                  = 'ShareGateReports'                                                   # Sharegate Directory for export reports
#$GCServer                               = 'BGB01DC1181.national.core.bbc.co.uk:3268'                           # GC to be used to find user information
$Location                                = Get-Location | Select-Object -ExpandProperty path                    # Current running directory
$ErrorCount                              = 0                                                                    # Counters to record errors
$CompletedCount                          = 0                                                                    # Counters to record Completed
$RequireConnectMicrosoftSharePoint       = $True
$RequireConnectMicrosoftOnline           = $True
$BatchesFolder   		                 = 'Batches'                                                            # Root Folder name
$LogsFolder                              = 'Log'                                                                # Log Files to store details in
$FailedFolder                            = 'Failed'                                                             # Failed Reports / Users
$PreMigrationReports                     = 'PreMigrationReports'                                                # PreMigrationReports Folder
$ExcludeListFile                         ='.\ExcludeUsers.txt'                                                  # List of Users to exclude by SAmaccountName


                                                                





function DisplayExtendedInfo () {

    # Display to notify the operator before running
    Clear-Host
    Write-Host 
    Write-Host 
    Write-Host  '----------------------------------------------------------------------------------'	
	Write-Host  '            OneDrive Pre-migration Check tool'                                        -ForegroundColor Green
	Write-Host  '----------------------------------------------------------------------------------'
    Write-Host
    Write-Host 
    Write-Host  '  This script is used to Check OneDrive users before Migration of data          '     -ForegroundColor YELLOW
    Write-Host  '                                                                                '          
    Write-Host  '** You must have at least the SharePoint Admin Role and Admin SharePoint Access *** ' -ForegroundColor YELLOW
    Write-Host  '----------------------------------------------------------------------------------'
    Write-Host 
}


# FUNCTION - As for Batch Name
function Batchname {

    # Prompt to input Batch file name
    $Global:BatchName = Read-Host 'Enter Batch Name (Leave blank for default)'
    if ($BatchName -eq "") {$Global:BatchName = "Batch_$((get-date).ToString('yyyyMMdd_HHmm'))"}
    Write-Host `n

    # Check if Batch folder already exists, prompt for new batch name if it does
    $ChkBatchFile = "$BatchName"
    if (Test-Path ".\$BatchesFolder\$ChkBatchFile") {
	Write-Host "Batch name " -NoNewLine
	Write-Host "$BatchName" -NoNewLine -ForegroundColor YELLOW
	Write-Host " has already been used, please use a different name"
	Write-Host `n
	$Global:BatchName = Read-Host 'Enter new Batch Name (Leave blank for default)'
	if ($BatchName -eq "") {$Global:BatchName = "Batch_$((get-date).ToString('yyyyMMdd_HHmm'))"}
	Write-Host `n
    }
}

# FUNCTION - Create Folders for Data
function CreateFolderStructure () {

    # Check that the BatchesFolder exists, create it if not
    if (! (Test-Path ".\$BatchesFolder\$BatchName")) {
        #Write-Host "Creating "$BatchName" sub-folder..."
		new-item -path .\ -name ".\$BatchesFolder\$BatchName" -type directory | out-null
    }

    
    # Check that the $LogsFolder exists, create it if not
    if (! (Test-Path ".\$BatchesFolder\$BatchName\$LogsFolder")) {
        #Write-Host "Creating '$LogsFolder' sub-folder..."
        new-item -path .\ -name ".\$BatchesFolder\$BatchName\$LogsFolder" -type directory | out-null
    }
       
   
    # Check that the FailedFolder exists, create it if not
    if (! (Test-Path ".\$BatchesFolder\$BatchName\$FailedFolder")) {
        #Write-Host "Creating '$FailedFolder' sub-folder..."
        new-item -path .\ -name ".\$BatchesFolder\$BatchName\$FailedFolder" -type directory | out-null
    }

        # Check that the ReportFolder exists, create it if not
    if (! (Test-Path ".\$BatchesFolder\$BatchName\$PreMigrationReports")) {
        #Write-Host "Creating '$ReportFolder' sub-folder..."
        new-item -path .\ -name ".\$BatchesFolder\$BatchName\$ReportFolder" -type directory | out-null
    }

      # Check that the ShareGate Reports directory exists, create it if not
    if (! (Test-Path ".\$BatchesFolder\$BatchName\$ShareGateReportsExport")) {
        #Write-Host "Creating '$ShareGateReportsExport' sub-folder..."
        new-item -path .\ -name ".\$BatchesFolder\$BatchName\$ShareGateReportsExport" -type directory | out-null
    }

          # Check that the Migration Reports directory exists, create it if not
    if (! (Test-Path ".\$BatchesFolder\$BatchName\$MigrationReports")) {
        #Write-Host "Creating '$MigrationReports' sub-folder..."
        new-item -path .\ -name ".\$BatchesFolder\$BatchName\$MigrationReports" -type directory | out-null
    }
    
}





function CreateOutPutFiles () {

    # Results file
	if (! (Test-Path ".\$BatchesFolder\$BatchName\$ReportFolder\$ExportResultsFile")) {
		New-Item -Path . -Name ".\$BatchesFolder\$BatchName\$ReportFolder\$ExportResultsFile" -ItemType "file" -Value 'UserFoundInAD,Displayname,UserPrincipalName,HomeDirectory,SharePointLicence,OneDriveProvisioned,OneDriveCurrentSize,OneDriveURL,PerferredMigrationServer,IdentifiedHomeDrive,HomeDirectorySize,HomeDirectoryNumFiles,SamaccountName' -Force | Out-Null
		Add-Content -path ".\$BatchesFolder\$BatchName\$ReportFolder\$ExportResultsFile" -value ""
	}

    # Missing Users File
	if (! (Test-Path ".\$BatchesFolder\$BatchName\$FailedFolder\$MissingUsers")) {
		New-Item -Path . -Name ".\$BatchesFolder\$BatchName\$FailedFolder\$MissingUsers" -ItemType "file" -Value 'SAMaccountName' | Out-Null
		Add-Content -path ".\$BatchesFolder\$BatchName\$FailedFolder\$MissingUsers" -value ""
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


# FUNCTION -  Check Active Directory Module is installed
function ActiveDirectoryModule () {

    if ($false -eq $UsePSSessionActiveDirectory) {

        WriteTransactionsLogs -Task "Checking ActiveDirectory  Module" -Result Information -ErrorMessage none -ShowScreenMessage true -ScreenMessageColour GREEN -IncludeSysError false 
        if (Get-Module -ListAvailable -Name ActiveDirectory) {
            WriteTransactionsLogs -Task "Found ActiveDirectory Module" -Result Information -ErrorMessage none -ShowScreenMessage true -ScreenMessageColour GREEN -IncludeSysError false
            Import-Module ActiveDirectory 	
        } else {
            WriteTransactionsLogs -Task "Failed to locate Active Directory Module" -Result Error -ErrorMessage "Active Directory Module not installed" -ShowScreenMessage true -ScreenMessageColour RED -IncludeSysError false 
            TerminateScript	
        }
    }
}


# FUNCTION - Load Active Directory Module via PS Session
function PSSessionActiveDirectory () {

    if ($true -eq $UsePSSessionActiveDirectory){

        WriteTransactionsLogs -Task "Connecting to $DCServer for Active Directory Module" -Result Information -ErrorMessage none -ShowScreenMessage true -ScreenMessageColour GREEN -IncludeSysError false 
        try {$PSSessions = Get-PSSession
            if ($PSSessions | Where-Object {$_.Computername -eq $DCServer}) {Remove-PSSession -ComputerName $DCServer} # Used to clear the existing session if exists
            $session = New-PSSession -ComputerName $DCserver -Authentication Kerberos -WarningAction  SilentlyContinue -ea STOP 
            Import-PSSession -Session $session -Module ActiveDirectory | Out-Null
            WriteTransactionsLogs -Task "Loaded Active Directory Module" -Result Information -ErrorMessage none -ShowScreenMessage true -ScreenMessageColour GREEN -IncludeSysError false 
        }
        Catch {WriteTransactionsLogs -Task "Could not load or connect for Active Directory Module" -Result ERROR -ErrorMessage "PS Session Failed: $error[0].Exception.message " -ShowScreenMessage true -ScreenMessageColour RED -IncludeSysError false 
            TerminateScript
        }
        
    }

}


$UsePSSessionActiveDirectory

# FUNCTION -  Check MSole Module is installed
function MSonlineModule () {

    WriteTransactionsLogs -Task "Checking Microsoft Online Module"  -Result Information -ErrorMessage none -ShowScreenMessage true -ScreenMessageColour GREEN -IncludeSysError false   
    if (Get-Module -ListAvailable -Name MSonline) {
        WriteTransactionsLogs -Task "Found Microsoft Online Module" -Result Information -ErrorMessage none -ShowScreenMessage true -ScreenMessageColour GREEN -IncludeSysError false 	
    } else {
        WriteTransactionsLogs -Task "Failed to locate Microsoft Online Module, it needs to be installed" -Result Error -ErrorMessage "Online Module not installed" -ShowScreenMessage true -ScreenMessageColour RED -IncludeSysError false
    TerminateScript	
    }
}


# FUNCTION -  Check SharePoint Module is installed
function SharePointModule () {

    WriteTransactionsLogs -Task "Checking Microsoft SharePoint Module"  -Result Information -ErrorMessage none -ShowScreenMessage true -ScreenMessageColour GREEN -IncludeSysError false   
    if (Get-Module -ListAvailable -Name Microsoft.Online.SharePoint.PowerShell ) {
        WriteTransactionsLogs -Task "Found Microsoft SharePoint Module" -Result Information -ErrorMessage none -ShowScreenMessage true -ScreenMessageColour GREEN -IncludeSysError false 	
    } else {
        WriteTransactionsLogs -Task "Failed to locate Microsoft SharePoint Module, it needs to be installed" -Result Error -ErrorMessage "SharePoint Module not installed" -ShowScreenMessage true -ScreenMessageColour RED -IncludeSysError false
    TerminateScript	
    }
}


# FUNCTION - Get Credentions
function AskForAdminCreds () {

    # Asking for creds if they don't exist 
    if ($MFALoginRequired -eq $false) {
        WriteTransactionsLogs -Task "Asking for Service Account Credentials"  -Result Information -ErrorMessage none -ShowScreenMessage true -ScreenMessageColour GREEN -IncludeSysError false
        $Global:credentials = Get-Credential -Message "Enter Login details for Office 365"
    }
}

# FUNCTION - Connect to 365 Microsoft Online
function ConnectMicrosoftOnline () {

    # Check Connection to 365 or Connect if not already
    if ($RequireConnectMicrosoftOnline -eq $true) {

        try {
         try { Get-MsolCompanyInformation -ea stop | Out-Null;  WriteTransactionsLogs -Task "Existing Msole Connection Found" -Result Information -ErrorMessage none -ShowScreenMessage true -ScreenMessageColour GREEN -IncludeSysError false}
         catch {
                WriteTransactionsLogs -Task "Not Connected to Msole" -Result Information -ErrorMessage none -ShowScreenMessage true -ScreenMessageColour GREEN -IncludeSysError false
                if ($MFALoginRequired -eq $True){Connect-MsolService  -ErrorAction Stop | Out-Null}
                if ($MFALoginRequired -eq $False){Connect-MsolService -Credential $Global:credentials  -ErrorAction Stop | Out-Null}
            }
        }  
        Catch {
            WriteTransactionsLogs -Task "Unable to Connect to Microsoft Online" -Result Error -ErrorMessage "Connect Error" -ShowScreenMessage true -ScreenMessageColour RED -IncludeSysError false
	        TerminateScript
        }
    }
}


# FUNCTION - Connect to SharePoint Online
function ConnectMicrosoftSharePoint () {

    # Check Connection to SharePoint or Connect if not already
    if ($RequireConnectMicrosoftSharePoint -eq $true) {

        try {
         try { Get-SPOTenant -ea stop | Out-Null;  WriteTransactionsLogs -Task "Existing SharePoint Connection Found" -Result Information -ErrorMessage none -ShowScreenMessage true -ScreenMessageColour GREEN -IncludeSysError false}
         catch {
                WriteTransactionsLogs -Task "Not Connected to SharePoint" -Result Information -ErrorMessage none -ShowScreenMessage true -ScreenMessageColour GREEN -IncludeSysError false
                if ($MFALoginRequired -eq $True){$Global:SharePointlogon = Connect-SPOService -Url $AdminSiteURL  -ErrorAction Stop | Out-Null}
                if ($MFALoginRequired -eq $False){Connect-SPOService -Url $AdminSiteURL -credential $credentials  -ErrorAction Stop | Out-Null}
            }
        }  
        Catch {
            WriteTransactionsLogs -Task "Unable to Connect to Microsoft SharePoint" -Result Error -ErrorMessage "Connect Error" -ShowScreenMessage true -ScreenMessageColour RED -IncludeSysError true
	        TerminateScript
        }
    }
}


# FUNCTION -  Check User Permissions
function CheckPermissions () {

    $ValidPermissions = ''
    WriteTransactionsLogs -Task "Checking Permissions"    -Result Information -ErrorMessage none -ShowScreenMessage true -ScreenMessageColour GREEN -IncludeSysError false 
   if ($ValidPermissions -like $null){try {$ValidPermissions = Get-MsolRoleMember -RoleObjectId 62e90394-69f5-4237-9190-012177145e10 | Where-Object {$_.emailaddress -eq $userAdminID}; WriteTransactionsLogs -Task "Found Admin in Global Administrators" -Result Information -ErrorMessage "none" -ShowScreenMessage true -ScreenMessageColour GREEN -IncludeSysError false} catch {WriteTransactionsLogs -Task "Permissions Error" -Result Information -ErrorMessage "Error happened searching Rbac Group" -ShowScreenMessage true -ScreenMessageColour RED -IncludeSysError false }}
   if ($ValidPermissions -like $null){try {$ValidPermissions = Get-MsolRoleMember -RoleObjectId f28a1f50-f6e7-4571-818b-6a12f2af6b6c | Where-Object {$_.emailaddress -eq $userAdminID}; WriteTransactionsLogs -Task "Found Admin in SharePoint Service Administrator" -Result Information -ErrorMessage "none" -ShowScreenMessage true -ScreenMessageColour GREEN -IncludeSysError false} catch {WriteTransactionsLogs -Task "Permissions Error" -Result Information -ErrorMessage "Error happened searching Rbac Group" -ShowScreenMessage true -ScreenMessageColour RED -IncludeSysError false}}   
   if ($ValidPermissions -like $null) {
       WriteTransactionsLogs -Task "Current user has no Permissions to perform the required actions" -Result Information -ErrorMessage none -ShowScreenMessage true -ScreenMessageColour GREEN -IncludeSysError false 
    TerminateScript	}
}

# FUNCTION - Get Userlogon Identity
function FindAdminLogonID () {

    if ($MFALoginRequired -eq $false) {$Global:userAdminID = $credentials.username}
    if ($MFALoginRequired -eq $true) {$Global:userAdminID = $SharePointlogon.account}
}

# FUNCTION -  Check input file exists, terminate script if not
function CheckTXTDataFile () {

    WriteTransactionsLogs -Task "Checking TXT File............"    -Result Information -ErrorMessage none -ShowScreenMessage true -ScreenMessageColour GREEN -IncludeSysError false 
    if (! (Test-Path $TXTDataFile)) {
	    WriteTransactionsLogs -Task "TXT File Check" -Result Error -ErrorMessage "TXT File Not found in expected location" -ShowScreenMessage true -ScreenMessageColour RED -IncludeSysError false
        TerminateScript
    } else {
        WriteTransactionsLogs -Task "Found TXT File..............."    -Result Information -ErrorMessage none -ShowScreenMessage true -ScreenMessageColour GREEN -IncludeSysError false 

        # Store a copy of the userlist in the Batch folder for reference
        $MigrationFileUsed = ".\$BatchesFolder\$BatchName\$TXTDataFile" + "_OriginalList.txt"
        Copy-Item $TXTDataFile $MigrationFileUsed
    }
}


# FUNCTION - Get TXT Data from File
function ImportTXTData () {

    WriteTransactionsLogs -Task "Importing Data file................"   -Result Information none -ShowScreenMessage true -ScreenMessageColour GREEN -IncludeSysError false    
    try {$Global:OneDriveObjects = Get-content $TXTDataFile -ea stop
        WriteTransactionsLogs -Task "Loaded Users Data"   -Result Information -ErrorMessage none -ShowScreenMessage true -ScreenMessageColour GREEN -IncludeSysError false
    } 
    catch {WriteTransactionsLogs -Task "Error loading Users data File" -Result Error -ErrorMessage "An error happened importing the data file, Please Check File" -ShowScreenMessage true -ScreenMessageColour RED -IncludeSysError false
        TerminateScript
    }

}


# FUNCTION - Connecto to GC for all Domain Objects
function ConnectToGCServer () {

    WriteTransactionsLogs -Task "Connecting to $DCServer"   -Result Information -ErrorMessage none -ShowScreenMessage true -ScreenMessageColour GREEN -IncludeSysError false 
    $TestPathDC = Test-Path "AD2:"

    if ($TestPathDC -eq "$true"){
        WriteTransactionsLogs -Task "Connection already exists"   -Result Information -ErrorMessage none -ShowScreenMessage true -ScreenMessageColour GREEN -IncludeSysError false 
    
    } Else {
        try {New-PSDrive -Name AD2 -PSProvider activedirectory -server $DCServer -Root "//rootdse/" -Scope Global | Out-Null
            WriteTransactionsLogs -Task "Connected to $DCServer" -Result Information -ErrorMessage none -ShowScreenMessage true -ScreenMessageColour GREEN -IncludeSysError false
        }
        Catch {WriteTransactionsLogs -Task "Failed Connecting to $DCServer" -Result Error -ErrorMessage "No Vaild connection:" -ShowScreenMessage true -ScreenMessageColour RED -IncludeSysError true
               TerminateScript
        }
    }

}


# FUNCTION - Check Exclude User list
function CheckExcludeList () {

    WriteTransactionsLogs -Task "Checking For Exclude List File............"    -Result Information -ErrorMessage none -ShowScreenMessage true -ScreenMessageColour GREEN -IncludeSysError false 
    if (! (Test-Path $ExcludeListFile)) {
	    WriteTransactionsLogs -Task "Exclude List File Check" -Result Information -ErrorMessage "Exclude File Not found in expected location" -ShowScreenMessage true -ScreenMessageColour YELLOW -IncludeSysError false
        $ExcludeListFileNotFound = $false
    } else {
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





# FUNCTION - Process Users
function ProcessUsers () {

    Write-host "################## Processing Users ################## " -ForegroundColor YELLOW
    Write-host ''


    ####################################### Main loop for all Users #####################################################

        Foreach ($UserObject in $OneDriveObjects) {

        # Markers to record data
        $UserError  = 0
        $UserValid  = ''
        $UserUPN    = ''
        $UserValid  = $true


        # Find User in Active Directory to be processed or set a marker to exclude them.
        #region GetADUserInfo
        
        # Clear any existing details
        $ADUserInfo    = ''
        $DisplayName   = ''
        $HomeDirectory = ''
        $UserUPN       = ''
        
        # Change location to AD
        # Set-Location AD2:

               
        # Find AD user Info and Store to var
        $ADUserInfo = Get-Aduser -LDAPFilter "(SamaccountName=$UserObject)"  -Properties *
        
        # Change location back to system
       # Set-location $Location
        
        # Create a marker to exclude user from other functions
        if ($null -eq $ADUserInfo){$UserFoundInAD = "$false"
            WriteTransactionsLogs -Task "$UserObject Not found in AD" -Result information -ShowScreenMessage true -ScreenMessageColour Yellow -IncludeSysError false
            $UserValid = $false
            $UserObject | out-file ".\$BatchesFolder\$BatchName\$FailedFolder\$MissingUsers" -Encoding ASCII -Append
        } Else {
            $UserFoundInAD = "$true"
            # Build Simple String
            $DisplayName        = $ADUserInfo.displayname
            $HomeDirectory      = $ADUserInfo.HomeDirectory
            $HomeDirectoryPath  = $ADUserInfo.HomeDirectory + "\*"
            $UserUPN            = $ADUserInfo.UserPrincipalName
            $Mail               = $ADUserInfo.Mail
            WriteTransactionsLogs -Task "$DisplayName found in AD" -Result information -ShowScreenMessage true -ScreenMessageColour Green -IncludeSysError false
            $host.ui.RawUI.WindowTitle = "OneDrive PreCheck Script - Processing $DisplayName"
            $ADUserInfo | Add-Member -MemberType NoteProperty -Name UserFoundInAD -Value $true -Force
            $UserValid = $true
            
        }
        #endregion GetADUserInfo



         # Remove User if found on exclude list 
         #region CheckAgainstExcludeList
        if ($null -eq $ExcludeListFileNotFound){
            
            # Compare exclude list with current user
            If ($ExcludeListUsers -contains "$UserObject") {$UserValid = $false
                WriteTransactionsLogs -Task "$UserObject found in exclude list and Remove from processing" -Result information -ShowScreenMessage true -ScreenMessageColour Yellow -IncludeSysError false}
        }

        #endregion CheckAgainstExcludeList


        # Check if User has a HomeDrive mapped in Active Directory and get server location
        #region CheckHomeDrive

        if ($UserValid -eq $true) {
             WriteTransactionsLogs -Task "$DisplayName has valid homedrive - $HomeDirectory " -Result information -ErrorMessage none -ShowScreenMessage true -ScreenMessageColour GREEN -IncludeSysError false
            
            # Find Server name in Path
            [string]$HomedriveServer = $ADUserInfo.HomeDirectory -split '\\' | select-object -Skip 2 -Last 1

            if ($HomedriveServer -match $PittsburgDCList) {$DCRegion = "Pittsburg" ; $SetPerferredMigrationServer = "Pittsburg Server1"
                $ADUserInfo | Add-Member -MemberType NoteProperty -Name PerferredMigrationServer -Value $SetPerferredMigrationServer -Force
                $ADUserInfo | Add-Member -MemberType NoteProperty -Name IdentifiedHomeDrive -Value $HomedriveServer -Force
                WriteTransactionsLogs -Task "$DisplayName assigned migration server $SetPerferredMigrationServer" -Result information -ErrorMessage none -ShowScreenMessage true -ScreenMessageColour GREEN -IncludeSysError false}

            if ($HomedriveServer -match $GreenfordDCList) {$DCRegion = "Greenford" ; $SetPerferredMigrationServer = "Greenford Server2"
                $ADUserInfo | Add-Member -MemberType NoteProperty -Name PerferredMigrationServer -Value $SetPerferredMigrationServer -Force
                $ADUserInfo | Add-Member -MemberType NoteProperty -Name IdentifiedHomeDrive -Value $HomedriveServer -Force
                WriteTransactionsLogs -Task "$DisplayName assigned migration server $SetPerferredMigrationServer" -Result information -ErrorMessage none -ShowScreenMessage true -ScreenMessageColour GREEN -IncludeSysError false}

            if ($HomedriveServer -match $AuroraDCList) {$DCRegion = "Aurora" ; $SetPerferredMigrationServer = "Aurora Server3"
                $ADUserInfo | Add-Member -MemberType NoteProperty -Name PerferredMigrationServer -Value $SetPerferredMigrationServer -Force
                $ADUserInfo | Add-Member -MemberType NoteProperty -Name IdentifiedHomeDrive -Value $HomedriveServer -Force
                WriteTransactionsLogs -Task "$DisplayName assigned migration server $SetPerferredMigrationServer" -Result information -ErrorMessage none -ShowScreenMessage true -ScreenMessageColour GREEN -IncludeSysError false}

            if ($null -eq $DCRegion -and $HomeDirectory -like "*\\*") {WriteTransactionsLogs -Task "$DisplayName Homedrive Check" -Result Error -ErrorMessage "Unidentified Server" -ShowScreenMessage true -ScreenMessageColour RED -IncludeSysError false
                $ADUserInfo | Add-Member -MemberType NoteProperty -Name IdentifiedHomeDrive -Value "Unidentified Server" -Force
                $UserValid = $false}

            if ($null -eq $DCRegion -or $null -eq $HomeDirectory) {WriteTransactionsLogs -Task "$DisplayName Homedrive is not assigned to Profile" -Result Error -ErrorMessage "No Migration Possible" -ShowScreenMessage true -ScreenMessageColour YELLOW -IncludeSysError false
                $ADUserInfo | Add-Member -MemberType NoteProperty -Name IdentifiedHomeDrive -Value "User has no Homedrive Directory" -Force
                $ADUserInfo | Add-Member -MemberType NoteProperty -Name PerferredMigrationServer -Value "No Server selected/No HomeDirectory" -Force
                $UserValid = $true
            
            }
                        
        
        }
        #endregion CheckHomeDrive

         

        # Get the size of the users Home Directory
        #region GetHomeDirectorySize

        If ($UserValid -eq $true -and $GetHomeDriveSize -eq $true){
            WriteTransactionsLogs -Task "Calculating HomeDirectory Size.....Please Wait" -Result information -ErrorMessage "none" -ShowScreenMessage true -ScreenMessageColour GREEN -IncludeSysError false
            try {$HomeDirectoryVisable = Test-Path $HomeDirectoryPath
                if ($HomeDirectoryVisable -eq $True) {
                    $HomeDriveSize = Get-FolderSizeInfo $ADUserInfo.HomeDirectory -ErrorAction Stop
                    $TotalSize  = $HomeDriveSize.totalSize
                    $TotalFiles = $HomeDriveSize.totalfiles
                    WriteTransactionsLogs -Task "$Displayname has the following stats TotalSize $TotalSize / TotalFiles $TotalFiles " -Result information -ErrorMessage "none" -ShowScreenMessage true -ScreenMessageColour GREEN -IncludeSysError false
                    $ADUserInfo | Add-Member -MemberType NoteProperty -Name HomeDirectorySize -Value $TotalSize -Force
                    $ADUserInfo | Add-Member -MemberType NoteProperty -Name HomeDirectoryNumFiles -Value $TotalFiles -Force
                }
                Else {
                   WriteTransactionsLogs -Task "Failed to get the HomeDirectory details" -Result Error -ErrorMessage "No Access or other error" -ShowScreenMessage true -ScreenMessageColour YELLOW -IncludeSysError false
                   $ADUserInfo | Add-Member -MemberType NoteProperty -Name HomeDirectorySize -Value 'No Data' -Force
                   $ADUserInfo | Add-Member -MemberType NoteProperty -Name HomeDirectoryNumFiles -Value 'No Data' -Force
                }
                
            }
            Catch {WriteTransactionsLogs -Task "Failed to get the HomeDirectory details" -Result Error -ErrorMessage "No Access or other error" -ShowScreenMessage true -ScreenMessageColour YELLOW -IncludeSysError true
                $ADUserInfo | Add-Member -MemberType NoteProperty -Name HomeDirectorySize -Value 'No Data' -Force
                $ADUserInfo | Add-Member -MemberType NoteProperty -Name HomeDirectoryNumFiles -Value 'No Data' -Force
            }
        }
        #endregion GetHomeDirectorySize



        # Get if User is valid in Msole
        #region GetMsolUser

        If ($UserValid -eq $true){
            try {$MsolUser = Get-MsolUser -UserPrincipalName $Mail -ErrorAction Stop
                WriteTransactionsLogs -Task "Found $Mail in Azure " -Result information -ErrorMessage "none" -ShowScreenMessage true -ScreenMessageColour GREEN -IncludeSysError false
             }
            Catch {WriteTransactionsLogs -Task "Failed to locate $Mail in Azure" -Result Error -ErrorMessage "$Mail was not found in Msol" -ShowScreenMessage true -ScreenMessageColour RED -IncludeSysError false
                $UserValid = $false
            }
        }
        #endregion GetMsolUser


        # Check Msole User is valid / health
        #region CheckMsolUser
        
        If ($UserValid -eq $true){
            If ($MsolUser.ValidationStatus -eq 'Healthy') {
                WriteTransactionsLogs -Task "Azure User Health Check Passed" -Result information -ErrorMessage "none" -ShowScreenMessage true -ScreenMessageColour GREEN -IncludeSysError false    
            
            }Else {
                WriteTransactionsLogs -Task "Azure User Health Check Failed" -Result information -ErrorMessage "Health Check Failed with result $MsolUser.ValidationStatus" -ShowScreenMessage true -ScreenMessageColour RED -IncludeSysError false
                $UserValid = $false
            }
        }
        #endregion CheckMsolUser


        # Check User has a SharePoint Licence
        #region CheckSharePointSKU

        If ($UserValid -eq $true){
            WriteTransactionsLogs -Task "Checking for SharePoint SKU" -Result information -ErrorMessage "none" -ShowScreenMessage true -ScreenMessageColour GREEN -IncludeSysError false
             if ($MsolUser.Licenses.ServiceStatus.ServicePlan.ServiceName -eq "SHAREPOINTENTERPRISE"){$SharePointLicFound = $true
                WriteTransactionsLogs -Task "User is assigned a SharePoint SKU" -Result information -ErrorMessage "none" -ShowScreenMessage true -ScreenMessageColour GREEN -IncludeSysError false
                $ADUserInfo | Add-Member -MemberType NoteProperty -Name SharePointLicence -Value $SharePointLicFound -Force
        
            } Else {
                WriteTransactionsLogs -Task "Check for SharePoint SKU Failed" -Result information -ErrorMessage "User has not SharePoint SKU or Failed check" -ShowScreenMessage true -ScreenMessageColour RED -IncludeSysError false
                $ADUserInfo | Add-Member -MemberType NoteProperty -Name SharePointLicence -Value "Not Assigned or has error" -Force
                $UserValid = $false
            }
        }
        #endregion CheckSharePointSKU


        # Get current Onedrive details
        #region GetOneDriveDetails

        If ($UserValid -eq $true -and $CheckSharePointProvision -eq $true){ 
            WriteTransactionsLogs -Task "Checking OneDrive for Provisioned Container.... Please Wait" -Result Information -ErrorMessage none -ShowScreenMessage true -ScreenMessageColour GREEN -IncludeSysError false 
            $OneDriveDetails = Get-SPOSite -Template "SPSPERS" -Limit ALL -includepersonalsite $True -Filter "owner -eq $mail"
            If ($OneDriveDetails -like $null){WriteTransactionsLogs -Task "OneDrive not provisioned for $mail" -Result Information -ErrorMessage none -ShowScreenMessage true -ScreenMessageColour YELLOW -IncludeSysError false
                $NotProvisionSharePointContainer = False

            } Else {
                WriteTransactionsLogs -Task "$mail is Provisioned for OneDrive" -Result Information -ErrorMessage none -ShowScreenMessage true -ScreenMessageColour GREEN -IncludeSysError false
                $ADUserInfo | Add-Member -MemberType NoteProperty -Name OneDriveProvisioned -Value $true -Force
                $OneDriveCurrentSize = $OneDriveDetails.StorageUsageCurrent
                $OneDriveURL = $OneDriveDetails.url
                $ADUserInfo | Add-Member -MemberType NoteProperty -Name OneDriveCurrentSize -Value $OneDriveCurrentSize -Force
                $ADUserInfo | Add-Member -MemberType NoteProperty -Name OneDriveURL -Value $OneDriveURL -Force
                WriteTransactionsLogs -Task "Current OneDrive Size is $OneDriveCurrentSize" -Result Information -ErrorMessage none -ShowScreenMessage true -ScreenMessageColour GREEN -IncludeSysError false
            }
        }
        #endregion GetOneDriveDetails

        
        # Provisioned OneDrive if required
        #region ProvisionedOneDrive

        If ($UserValid -eq $true -and $ProvisionSharePointContainer -eq $true -and $NotProvisionSharePointContainer -eq $false){
            WriteTransactionsLogs -Task "Requesting OneDrive Provisioning" -Result Information -ErrorMessage none -ShowScreenMessage true -ScreenMessageColour YELLOW -IncludeSysError false
            try {Request-SPOPersonalSite -UserEmails $Mail -NoWait
                WriteTransactionsLogs -Task "Requesting Completed" -Result Information -ErrorMessage none -ShowScreenMessage true -ScreenMessageColour GREEN -IncludeSysError false
                $ADUserInfo | Add-Member -MemberType NoteProperty -Name OneDriveProvisioned -Value $true -Force}
            Catch {WriteTransactionsLogs -Task "Request OneDrive Provisioning Failed" -Result Information -ErrorMessage "Error:" -ShowScreenMessage true -ScreenMessageColour RED -IncludeSysError true
                #$UserValid = $false
            }
        }
        #endregion ProvisionedOneDrive


        # Export Results 
        #region ExportResults
        
        If ($UserValid -eq $true){
            $ADUserInfo |  Select-Object UserFoundInAD,Displayname,UserPrincipalName,HomeDirectory,SharePointLicence,OneDriveProvisioned,OneDriveCurrentSize,OneDriveURL,PerferredMigrationServer,IdentifiedHomeDrive,HomeDirectorySize,HomeDirectoryNumFiles,SamaccountName |export-csv ".\$BatchesFolder\$BatchName\$ReportFolder\$ExportResultsFile" -NoTypeInformation -Append
             Write-host "################## Completed OneDrive Check for $Displayname ################## " -ForegroundColor YELLOW
             Write-host ''
             $CompletedCount ++
        }
        #endregion ExportResults


        #Failed Users 
        #region FailedUsers
        
        If ($UserValid -eq $false){
            
             Write-host "################## Faied OneDrive Check for $UserObject ################## " -ForegroundColor RED
             Write-host ''
             $errorcount ++
        }
        #endregion FailedUsers

        
        # Clear Results 
        #region ClearResults
        
        If ($ADUserInfo){Remove-Variable ADUserInfo -force}
        #endregion ClearResults

        
    }   ######################### End of Foreach ##########################

    $host.ui.RawUI.WindowTitle = "OneDrive PreCheck Script - Completed"
    Write-host ''
    Write-host "Total Completed $CompletedCount" -ForegroundColor GREEN -NoNewline
    Write-host " | " -NoNewline
    Write-host "Total Failed $ErrorCount" -ForegroundColor RED
    WriteTransactionsLogs -Task "Total Completed $CompletedCount and Total Failed $ErrorCount" -Result Information -ErrorMessage "Information" -ShowScreenMessage false -ScreenMessageColour white -IncludeSysError false


}   ################## End of Process Function ########################



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
        # if (! (Test-Path ".\$BatchesFolder\$BatchName\$LogsFolder")) {new-item -path .\ -name ".\$BatchesFolder\$BatchName\$LogsFolder" -type directory | out-null}
        # created aove in the structure section of the code
 
 
 
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

# FUNCTION - Get Folder Sizes
Function Get-FolderSizeInfo {

    [cmdletbinding()]
    [alias("gsi")]
    [OutputType("FolderSizeInfo")]

    Param(
        [Parameter(Position = 0, Mandatory, HelpMessage = "Enter a file system path like C:\Scripts.", ValueFromPipeline, ValueFromPipelineByPropertyName)]
        [ValidateNotNullOrEmpty()]
        [alias("PSPath")]
        [string[]]$Path,
        [Parameter(HelpMessage = "Include hidden directories")]
        [switch]$Hidden
    )

    Begin {
        Write-Verbose "Starting $($MyInvocation.MyCommand)"
    } #Begin

    Process {
        foreach ($item in $path) {
            $cPath = (Convert-Path -literalpath $item)
            Write-Verbose "Measuring $cPath on $([System.Environment]::MachineName)"

            if (Test-Path -literalpath $cPath) {

                $d = [System.IO.DirectoryInfo]::new($cPath)

                $files = [system.collections.arraylist]::new()

                If ($psversiontable.psversion.major -gt 5  ) {
                    #this .NET class is not available in Windows PowerShell 5.1
                    $opt = [System.IO.EnumerationOptions]::new()
                    $opt.RecurseSubdirectories = $True

                    if ($hidden) {
                        Write-Verbose "Including hidden files"
                        $opt.AttributesToSkip = "SparseFile", "ReparsePoint"
                    }
                    else {
                        $opt.attributestoSkip = "Hidden"
                    }

                    $data = $($d.GetFiles("*", $opt))
                    if ($data.count -gt 1) {
                        $files.AddRange($data)
                    }
                    elseif ($data.count -eq 1) {
                        [void]($files.Add($data))
                    }

                } #if newer that Windows PowerShell 5.1
                else {
                    Write-Verbose "Using legacy code"
                    #need to account for errors when accessing folders without permissions
                    #a function to recurse and get all non-hidden directories

                    Function _enumdir {
                        [cmdletbinding()]
                        Param([string]$Path, [switch]$Hidden)
                        # write-host $path -ForegroundColor cyan
                        $path = Convert-Path -literalpath $path
                        $ErrorActionPreference = "Stop"
                        try {
                            $di = [System.IO.DirectoryInfo]::new($path)
                            if ($hidden) {
                                $top = $di.GetDirectories()
                            }
                            else {
                                $top = ($di.GetDirectories()).Where( {$_.attributes -notmatch "hidden"})
                            }
                            $top
                            foreach ($t in $top) {
                                $params = @{
                                    Path   = $t.fullname
                                    Hidden = $Hidden
                                }
                                _enumdir @params
                            }
                        }
                        Catch {
                           # Write-Warning "Failed on $path. $($_.exception.message)."
                        }
                    } #enumdir

                    # get files in the root of the folder
                    if ($hidden) {
                        Write-Verbose "Including hidden files"
                        $data = $d.GetFiles()
                    }
                    else {
                        #get files in current location
                        $data = $($d.GetFiles()).Where({$_.attributes -notmatch "hidden"})
                    }

                    if ($data.count -gt 1) {
                        $files.AddRange($data)
                    }
                    elseif ($data.count -eq 1) {
                        [void]($files.Add($data))
                    }

                    #get a list of all non-hidden subfolders
                    Write-Verbose "Getting subfolders (Hidden = $Hidden)"
                    $eParam = @{
                        Path   = $cpath
                        Hidden = $hidden
                    }
                    $all = _enumdir @eparam

                    #get the files in each subfolder
                    Write-Verbose "Getting files from $($all.count) subfolders"

                    ($all).Foreach( {
                            Write-Verbose $_.fullname
                            $ErrorActionPreference = "Stop"
                            Try {
                                if ($hidden) {
                                    $data = (([System.IO.DirectoryInfo]"$($_.fullname)").GetFiles())
                                }
                                else {
                                    $data = (([System.IO.DirectoryInfo]"$($_.fullname)").GetFiles()).where({$_.Attributes -notmatch "Hidden"})
                                }
                                if ($data.count -gt 1) {
                                    $files.AddRange($data)
                                }
                                elseif ($data.count -eq 1) {
                                    [void]($files.Add($data))
                                }
                            }
                            Catch {
                                #Write-Warning "Failed on $path. $($_.exception.message)."
                                #clearing the variable as a precaution
                                Clear-variable data
                            }
                        })
                } #else 5.1

                If ($files.count -gt 0) {
                    Write-Verbose "Found $($files.count) files"
                    # there appears to be a bug with the array list in Windows PowerShell
                    # where it doesn't always properly enumerate. Passing the list
                    # items via ForEach appears to solve the problem and doesn't
                    # adversely affect PowerShell 7. Addeed in v2.22.0. JH
                    $stats = $files.foreach( {$_}) | Measure-Object -property length -sum
                    $totalFiles = $stats.count
                    $totalSize = $stats.sum
                }
                else {
                    Write-Verbose "Found an empty folder"
                    $totalFiles = 0
                    $totalSize = 0
                }

                [pscustomobject]@{
                    PSTypename   = "FolderSizeInfo"
                    Computername = [System.Environment]::MachineName
                    Path         = $cPath
                    Name         = $(Split-Path $cpath -leaf)
                    TotalFiles   = $totalFiles
                    TotalSize    = $totalSize
                }
            } #test path
            else {
                Write-Warning "Can't find $Path on $([System.Environment]::MachineName)"
            }

        } #foreach item
    } #process
    End {
        Write-Verbose "Ending $($MyInvocation.MyCommand)"
    }
} 







#### Function Run Order
 DisplayExtendedInfo
 Batchname
 CreateFolderStructure
 CreateOutPutFiles
 ActiveDirectoryModule
 PSSessionActiveDirectory
 MSonlineModule
 SharePointModule
 CheckTXTDataFile
 AskForAdminCreds
 ConnectMicrosoftOnline
 ConnectMicrosoftSharePoint
 #ConnectToGCServer
 FindAdminLogonID
 CheckPermissions
 ImportTXTData
 CheckExcludeList
 ImportExcludeList 
 ProcessUsers

 
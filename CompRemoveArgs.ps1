[CmdletBinding(DefaultParameterSetName = 'Dispose')]
Param
(
    [Parameter(ParameterSetName = 'Dispose', Mandatory = $true, ValueFromPipelineByPropertyName = $true, ValueFromPipeline = $true)]
    [Parameter(ParameterSetName = 'Individual', Mandatory = $true, ValueFromPipelineByPropertyName = $true, ValueFromPipeline = $true)]
    [string]$ComputerName = $env:COMPUTERNAME,
    [Parameter(ParameterSetName = 'Dispose')]
    [switch]$Dispose = $false,
    [Parameter(ParameterSetName = 'Individual')]
    [switch]$AD,
    [Parameter(ParameterSetName = 'Individual')]
    [switch]$Azure,
    [Parameter(ParameterSetName = 'Individual')]
    [switch]$Intune,
    [Parameter(ParameterSetName = 'Individual')]
    [switch]$Autopilot,
    [Parameter(ParameterSetName = 'Individual')]
    [switch]$MECM
)

# Logging Funciton
Function LogIt {
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory = $true)]
        [ValidateNotNullorEmpty()]
        [string]$string
    )
    $logtext = "$(get-date -Format "yyyy-MM-d HH:mm:ss:fff zzz") - $string" 
    Write-Host $logText
    $logtext | Out-File -FilePath "$LocalLog" -Encoding unicode -Append
}

# Create Log Location (If not exist, it should)
if (!(Test-Path "C:\Temp")) {
    New-Item -ItemType Directory -Path "C:\Temp"
}


if ($null -eq $ComputerName) {
    $ComputerName = $env:COMPUTERNAME
    [bool]$DiffComp = $false
}
$ComputerName = $ComputerName.ToUpper()

# Starting Values
$LoggedOnUser = (qwinsta /SERVER:$env:COMPUTERNAME) -replace '\s{2,22}', ',' | ConvertFrom-Csv | Where-Object { $_ -like "*Acti*" } | Select-Object -ExpandProperty Username
$LocalLog = "C:\Temp\$ComputerName.log"
$NetworkFolder = "\\Server\Share\Removals"
$NetworkLog = "$NetworkFolder\$ComputerName.log"

# Confirm Entered Hostname
if ($ComputerName -notlike $env:COMPUTERNAME) {
    [bool]$DiffComp = $true
    $prompt = Read-Host -Prompt "You have entered:$ComputerName. Machine you are currently on is $env:Computername. To remove $ComputerName, please physically go to that machine and run script again.
    Press any key to continue..."
    if ($prompt) {
        LogIt "Issue with Computer Name:$ComputerName. To remove $Computername, please physically go to that machine and run script again."
        $errCR = "$errCR`nIssue with Computer Name:$ComputerName. To remove $Computername, please physically go to that machine and run script again."
        $err++
        exit 1
    }
}
else {
    [bool]$DiffComp = $false
    $prompt = Read-Host -Prompt "You have entered:$ComputerName. THIS IS THE COMPTUER YOU ARE CURRENTLY ON. Is this the computer you want to purge from the network? (y/N)"
    if (($prompt -notlike "y") -and ($prompt -notlike "yes")) {
        LogIt "Issue with Computer Name:$ComputerName. You have made an invalid selection at the prompt. Please only respond with 'y' or 'yes'."
        $errCR = "$errCR`nIssue with Computer Name:$ComputerName. You have made an invalid selection at the prompt. Please only respond with 'y' or 'yes'."
        $err++
        exit 1
    }
}

if ($Dispose) {
    $Intune = $true
    $Autopilot = $true
    $Azure = $true
    $MECM = $true
    $AD = $true
}


$NetworkLog = "$NetworkFolder\$ComputerName.log"

# Begin
LogIt "ComputerName: $computername"
LogIt "========Log-Begin========"


# Confirmation of selected arguments. For logging puproses
$Selection = "Removing $computername from:"
if ($Intune -or $Dispose) { $Selection = "$Selection Intune," }
if ($Autopilot -or $Dispose ) { $Selection = "$Selection Auto Pilot," }
if ($Azure -or $Dispose) { $Selection = "$Selection Azure Active Directory," }
if ($AD -or $Dispose) { $Selection = "$Selection Active Directory," }
if ($MECM -or $Dispose) { $Selection = "$Selection Configurtation Manager," }
LogIt $($Selection -replace ".$")

# Install Dependencies needed based on selected arguments
Try {
    LogIt " Installing Dependencies. This may take a few minutes. Please wait..."
    Set-ExecutionPolicy -Scope Process -Force Unrestricted
    if ($(Get-PackageProvider | Where-Object { ($_.name -eq "NuGet") }).Version -lt "2.8.5.201") {
        LogIt "NuGet not found or version below minimum needed. Installing\Updating...."
        try {
            Install-PackageProvider -Name NuGet -MinimumVersion "2.8.5.201" -Scope CurrentUser -Force
            LogIt "Compatible NuGet Installed"
        }
        Catch {
            LogIt "$($_.Exception.Message)"
            LogIt "Issue with NuGet Dependency. NuGet 2.8.5.201 or newer. Module not able to be installed"
            $errCR = "$errCR`nIssue with NuGet Dependency. NuGet 2.8.5.201 or newer. Module not able to be installed"
            $err++
            Exit 1
        }
    }
    Else {
        LogIt "NuGet already installed and version is at least 2.8.5.201"
    }

    if ($Intune -or $Autopilot -or $Azure -or $Dispose) {
        if (Get-Module -ListAvailable -Name Microsoft.Graph.Intune) {
            LogIt "Microsoft Graph Intune already Installed"
        }
        else {
            try {
                Install-Module Microsoft.Graph.Intune -Scope CurrentUser -Force 
                LogIt "Microsoft Graph Intune Installed"
            }
            Catch {
                LogIt "$($_.Exception.Message)"
                LogIt "Issue with Microsoft Graph Dependency. Module not able to be installed"
                $errCR = "$errCR`nIssue with Microsoft Graph Dependency. Module not able to be installed"
                $err++
                exit 1
            }
        }
    }
    if ($Azure -or $Dispose) {
        <#
        if (Get-Module -ListAvailable -Name AzureAD) {
            LogIt "AzureAD already Installed"
        }
        else { 
            Install-Module -Name AzureAD -Scope CurrentUser -Force -AllowClobber 
            LogIt "AzureAD Installed"
        }#>
        #<#
        if (Get-Module -ListAvailable -Name AzureADPreview) {
            LogIt "AzureADPreview already Installed"
        }
        else { 
            try {
                Install-Module -Name AzureADPreview -Scope CurrentUser -Force -AllowClobber 
                LogIt "AzureADPreview Installed"
            }
            Catch {
                LogIt "$($_.Exception.Message)"
                LogIt "Issue with Azure AD Dependency. Module not able to be installed"
                $errCR = "$errCR`nIssue with Azure AD Dependency. Module not able to be installed"
                $err++
                exit 1
            }
        }#>

        if (Get-Module -ListAvailable -Name msal.ps) {
            LogIt "msal.ps already Installed"
        }
        else { 
            Try {
                Install-Package msal.ps -Scope CurrentUser -Force 
                LogIt "msal.ps Installed"
            }
            Catch {
                LogIt "$($_.Exception.Message)"
                LogIt "Issue with Microsoft Aithentication Library Dependency. Module not able to be installed"
                $errCR = "$errCR`nIssue with Microsoft Aithentication Library Dependency. Module not able to be installed"
                $err++
                exit 1
            }
        }

        if (Get-Module -ListAvailable -Name MSOnline) {
            LogIt "MSOnline already Installed"
        }
        else { 
            Try {
                Install-Module -Name MSOnline -Scope CurrentUser -Force -ErrorAction SilentlyContinue 
                LogIt "MSOnline Installed"
            }
            Catch {
                LogIt "$($_.Exception.Message)"
                LogIt "Issue with Microsoft Online Dependency. Module not able to be installed"
                $errCR = "$errCR`nIssue with Microsoft Online Dependency. Module not able to be installed"
                $err++
                exit 1
            }
        }
    }
}
Catch {
    LogIt "$($_.Exception.Message)" 
    LogIt "Issue with Dependency Installations"
    $errCR = "$errCR`nIssue with Dependency Installations"
    #$err++
    Exit 1
}
LogIt "Successful Install of all Dependencies"

# Load required modules based on arguments
Try {
    # Cloud (Intune, Azure AD, or Auto Pilot) Only
    if ($Intune -or $AutoPilot -or $Azure -or $Dispose) {
        If ("Microsoft.Graph.Intune" -in $(Get-Module).Name) {
            LogIt "Microsoft Graph Intune already Imported"
        }
        Else {
            Try {
                Import-Module Microsoft.Graph.Intune -Force -ErrorAction Stop
                LogIt "Microsoft Graph Intune successfully Imported"
            }
            Catch {
                LogIt "$($_.Exception.Message)"
                LogIt "Issue with Microsoft Graph. Module not able to be imported"
                $errCR = "$errCR`nIssue with Microsoft Graph. Module not able to be imported"
                $err++
                exit 1
            }
        }
        Try {
            LogIt "Logged on user is: $LoggedOnUser"
            # Create the ACE
            $identity = "Contoso\$LoggedOnUser"
            $rights = 'FullControl' #Other options: [enum]::GetValues('System.Security.AccessControl.FileSystemRights')
            $inheritance = 'ContainerInherit, ObjectInherit' #Other options: [enum]::GetValues('System.Security.AccessControl.Inheritance')
            $propagation = 'None' #Other options: [enum]::GetValues('System.Security.AccessControl.PropagationFlags')
            $type = 'Allow' #Other options: [enum]::GetValues('System.Security.AccessControl.AccessControlType')
            $Rules = New-Object System.Security.AccessControl.FileSystemAccessRule($identity, $rights, $inheritance, $propagation, $type)
            # Grant access to admin home folder to non-admin account of logged in user. Should aleviate some issues with EdgeWebView2 during MSGraph/msal (Intune Powershell) Login.
            LogIt "Getting permissions of $env:USERPROFILE"
            $ACL = get-acl -Path $env:USERPROFILE
            $ACL.AddAccessRule($Rules)
            LogIt "Attempeting to grant permissions of $env:USERPROFILE folder to $LoggedOnUser. This may take a few minutes. Please wait..."
            Set-Acl -Path "$env:USERPROFILE" -AclObject $Acl
            LogIt "Granted permissions of $env:USERPROFILE folder to $LoggedOnUser"
        }
        Catch {
            LogIt "$($_.Exception.Message)"
            LogIt "Unable to grant permissions of $env:USERPROFILE to $LoggedOnUser"
            Continue
        }
    }
    # Azure Active Directory Only
    if ($Azure -or $Dispose) {
        <#
            If ("AzureAD" -in $(Get-Module).Name) {
                LogIt "AzureAD already Imported"
            }
            Else {
                Import-Module AzureAD -Force -ErrorAction Stop
                LogIt "AzureAD successfully Imported"
            }#>
        #<#
        If ("AzureADPreview" -in $(Get-Module).Name) {
            LogIt "AzureADPreview already Imported"
        }
        Else {
            Try {
                Import-Module AzureADPreview -Force -ErrorAction Stop
                LogIt "AzureADPreview successfully Imported"
            }
            Catch {
                LogIt "$($_.Exception.Message)"
                LogIt "Issue with Azure AD. Module not able to be imported"
                $errCR = "$errCR`nIssue with Azure AD. Module not able to be imported"
                $err++
                exit 1
            }
        }#>
    }
    # Config Manager only
    if ($MECM -or $Dispose) {
        If ("ConfigurationManager" -in $(Get-Module).Name) {
            LogIt "ConfigurationManager already Imported"
        }
        Else {
            Try {
                #Import-Module $env:SMS_ADMIN_UI_PATH.Replace('i386', 'ConfigurationManager.psd1') -Force -ErrorAction Stop
                #Import-Module "\\server\share$\Program Files\Microsoft Configuration Manager\AdminConsole\bin\ConfigurationManager\ConfigurationManager.psd1"  -Force -ErrorAction Stop
                Import-Module "\\server\ConfigurationManager_bin\ConfigurationManager.psd1" -Force -ErrorAction Stop # Might need to change this based on location and permission.
                LogIt "ConfigurationManager successfully Imported"
            }
            Catch {
                LogIt "$($_.Exception.Message)"
                LogIt "Issue with Configuration Manager. Module not able to be imported"
                $errCR = "$errCR`nIssue with Configuration Manager. Module not able to be imported"
                $err++
                exit 1
            }
        }
    }
    LogIt "Successfully Imported Modules"
}
Catch {
    LogIt "$($_.Exception.Message)"
    #$err++
    LogIt "Issue with Importing Modules"
    $errCR = "$errCR`nIssue with Importing Modules"
}

# Cloud Conneciton
if ($Intune -or $AutoPilot -or $Azure -or $Dispose) {
    Try {
        LogIt "Authenticating with MS Graph and Azure AD..." 
        #Get user cred and store to secure string variable 
        $Creds = Get-Credential -Message "Enter credentials for Azure logon (Use UPN for usename, most likely email address)" -UserName "@contoso.com"
        LogIt "Entered $($Creds.UserName)"
        #Authenticate to Cloud stuffs
        LogIt "Connecting to MSGraph..."
        $intuneId = Connect-MSGraph -Credential $Creds -ErrorAction Stop
        LogIt "Connecting to AzureAD..."
        $aadId = Connect-AzureAD -AccountId $intuneId.UPN -Credential $Creds -ErrorAction Stop
        $TenantId = $aadId.TenantId
        LogIt "Getting Azure User Info"
        $AccountId = $(Get-AzureADUser -ObjectId $($intuneId.UPN)).ObjectId
        LogIt "UPN: $($intuneId.UPN) is Azure ID:$AccountId"
        LogIt "Successful connection to Azure Cloud. Please note, times in next section are in GMT."
    }
    Catch {
        if ($Creds.UserName -notlike "*@finra.org") {
            LogIt "Issue with Initial Cloud Connection. UPN should be email address. $($Creds.UserName) was entered. Verify your UPN\password combination. Then try again."
            LogIt "$($_.Exception.Message)"
            $err++
            $errCR = "$errCR`nIssue with Initial Cloud Connection.  UPN should be email address. $($Creds.UserName) was entered. Verify your UPN\password combination. Then try again."
            exit 1
        }
        else {
            LogIt "Issue with Initial Cloud Connection. Verify that you are a member of PPL_ROLE_TEC_DskSppt (AD Group, for Delegated Cloud Permissions), and that you typed your password correctly. Then try again."
            LogIt "$($_.Exception.Message)"
            $err++
            $errCR = "$errCR`nIssue with Initial Cloud Connection. Verify that you are a member of PPL_ROLE_TEC_DskSppt (AD Group, for Delegated Cloud Permissions), and that you typed your password correctly. Then try again."
            exit 1
        }
    }

    # Set schedule and Reason for PIM elevation (Will appear in Azure logs and notify watchers)
    $Schedule = New-Object Microsoft.Open.MSGraph.Model.AzureADMSPrivilegedSchedule
    $Schedule.Type = "Once"
    $Schedule.StartDateTime = (Get-Date).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ss.fffZ")
    $Schedule.endDateTime = ((Get-Date).AddMinutes(10)).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ss.fffZ")
    $Reason = "$($intuneId.UPN): Removing Device - $ComputerName"

    #Get all PIM roles in tenant
    $AzureADMSPrivilegedRoleDefinition = Get-AzureADMSPrivilegedRoleDefinition -ProviderId 'aadRoles' -ResourceId $TenantId
    # Fetch all PIM role RoleDefinitionIds assigned to the current user.
    $AzureADMSPrivilegedRoleAssignment = (Get-AzureADMSPrivilegedRoleAssignment -ProviderId 'aadRoles' -ResourceId $TenantId -Filter "subjectId eq '$($AccountId)'").RoleDefinitionId
    # Fetch IDs for desired roles
    $Roles = 'Cloud Device Administrator'
    foreach ($role in $Roles) {
        $RoleID = ($AzureADMSPrivilegedRoleDefinition | Where-Object { $role -eq $_.DisplayName }).Id
        if ($AzureADMSPrivilegedRoleAssignment -contains $RoleID) {
            LogIt "Successfully found $role in assignements for $($intuneId.UPN)"
            Try {
                LogIt $(Open-AzureADMSPrivilegedRoleAssignmentRequest -ProviderId 'aadRoles' -ResourceId $TenantId -RoleDefinitionId $RoleID -SubjectId $AccountId -Type 'UserAdd' -AssignmentState 'Active' -Schedule $Schedule -Reason $Reason)
            }
            Catch {
                LogIt "$($_.Exception.Message)"
                $err++
                LogIt "Error with PIM $role. Unknown connection issue."
                $errCR = "$errCR`nIssue with $role PIM. Unknown connection issue."
                exit 1
            }
            LogIt "Waiting a little bit for PIM"
            Start-Sleep -Seconds 25
            LogIt "Resuming Removal"
        }
        else {
            #LogIt "$($_.Exception.Message)"
            $err++
            LogIt "Error with PIM $role. Your account has not been delegated $role. Submit AMG request or contant IAM."
            $errCR = "$errCR`nIssue with $role PIM. Your account has not been deledated $role. Submit AMG request or contant IAM."
            exit 1
        }
    }
    # Get Intune device data, this is needed to remove from AutoPilot devices later, and/or used to delete Intune device
    if ($Intune -or $AutoPilot -or $Dispose) {
        Try {
            LogIt "Retrieving " 
            LogIt "Intune " 
            LogIt "managed device record/s..." 
            [array]$IntuneDevices = Get-IntuneManagedDevice -Filter "deviceName eq '$computername'" -ErrorAction Stop
            If ($IntuneDevices.Count -ge 1) {
                LogIt "Successfully retrieved Intune record for $($IntuneDevice.deviceName)"
                # Remove From Intune
                If ($Intune -or $Dispose) {
                    $i = 0;
                    foreach ($IntuneDevice in $IntuneDevices) {
                        LogIt "   Deleting DeviceName: $($IntuneDevice.deviceName)  |  Id: $($IntuneDevice.Id)  |  AzureADDeviceId: $($IntuneDevice.azureADDeviceId)  |  SerialNumber: $($IntuneDevice.serialNumber) ..." 
                        Remove-IntuneManagedDevice -managedDeviceId $IntuneDevice.Id -Verbose -ErrorAction Stop
                        LogIt "Successfully removed $($IntuneDevice.deviceName) from Intune"
                        $i++
                    }
                }
            }
            Else {
                LogIt "$computername Not found in Intune!"
            }
        }
        Catch {
            if ($i -le 1) {
                LogIt "$($_.Exception.Message)"
                LogIt "Error with Finding or Deleting Intune Device!"
                $errCR = "$errCR`nIssue with Finding or Deleting Intune Device"
                $err++
                $_
            }
            else {
                LogIt "Error with Finding or Deleting Intune Device! Appears to be false positive from Co-Magement device duplicate."
                LogIt $errCR ="$errCR`nError with Finding or Deleting Intune Device! Appears to be false positive from Co-Magement device duplicate."
            }
        }
        # Remove from Auto Pilot
        if ($AutoPilot -or $Dispose) {
            LogIt "Retrieving " 
            LogIt "Autopilot "  
            LogIt "device registration..." 
            If ($IntuneDevices.Count -ge 1) {
                Try {
                    $AutopilotDevices = New-Object System.Collections.ArrayList
                    foreach ($IntuneDevice in $IntuneDevices) {
                        $URI = "https://graph.microsoft.com/beta/deviceManagement/windowsAutopilotDeviceIdentities?`$filter=contains(serialNumber,'$($IntuneDevice.serialNumber)')"
                        $AutopilotDevice = Invoke-MSGraphRequest -Url $uri -HttpMethod GET -ErrorAction Stop
                        [void]$AutopilotDevices.Add($AutopilotDevice)
                    }
                    if ($AutopilotDevices -ge 1) {
                        LogIt "Successfully retrieved AutoPilot device information for $($IntuneDevice.serialNumber)."
            
                        foreach ($device in $AutopilotDevices) {
                            LogIt "   Deleting SerialNumber: $($Device.value.serialNumber)  |  Model: $($Device.value.model)  |  Id: $($Device.value.id)  |  GroupTag: $($Device.value.groupTag)  |  ManagedDeviceId: $($device.value.managedDeviceId) ..."
                            $URI = "https://graph.microsoft.com/beta/deviceManagement/windowsAutopilotDeviceIdentities/$($device.value.Id)"
                            $AutopilotDevice = Invoke-MSGraphRequest -Url $uri -HttpMethod DELETE -ErrorAction Stop
                            LogIt "Successfully removed record from AutoPilot for $($IntuneDevice.serialNumber)." 
                        }
                    }
                }
                Catch {
                    if ($AutopilotDevices -ge 1) {
                        LogIt "$($_.Exception.Message)"
                        LogIt "Error with Finding or Deleting AutoPilot Device!"
                        $err++
                        $errCR = "$errCR`nError with Finding or Deleting AutoPilot Device!"
                        $_
                    }
                    else {
                        LogIt "Warning, no AutoPilot Device informaiton found for $($IntuneDevice.serialNumber). (This is not an Error. Please continue.)"
                    }
                }
            }
            Else {
                LogIt "$computername not Found in Intune. No data to search AutoPilot with."
            }
        }
    }
}

#Remove From Azure AD
if ($Azure -or $Dispose) {
    Try {
        LogIt "Retrieving " 
        LogIt "Azure AD "  
        LogIt "device record/s..."  
        [array]$AzureADDevices = Get-AzureADDevice -SearchString $computername -All:$true -ErrorAction Stop
        If ($AzureADDevices.Count -ge 1) {
            LogIt "Successfully retrieved Azure record for $($AzureADDevice.DisplayName) " 
            Foreach ($AzureADDevice in $AzureADDevices) {
                if ($computername -eq $txtCN.Text) {
                    LogIt "   Deleting Hostname: $($AzureADDevice.DisplayName)  |  ObjectId: $($AzureADDevice.ObjectId)  |  DeviceId: $($AzureADDevice.DeviceId) ..."
                    Remove-AzureADDevice -ObjectId $AzureADDevice.ObjectId -ErrorAction Stop
                    LogIt "Successfully Removed $($AzureADDevice.DisplayName) from Azure." 
                }
            }      
        }
        Else {
            LogIt "$computername Not found in Azure!"
        }
    }
    Catch {
        LogIt "Error with Finding or Deleting Azure Device!" 
        $err++
        $errCR = "$errCR`nIssue with Finding or Deleting Azure Device"
        $_
    }
}
# Remove from Configuraiton Manager
if ($MECM -or $Dispose) {
    Try {
        LogIt "Retrieving " 
        LogIt "ConfigMgr " 
        LogIt "device record/s..." 
        $SiteCode = (Get-PSDrive -PSProvider CMSITE -ErrorAction SilentlyContinue).Name
        if ($null -eq $SiteCode) { $SiteCode = "" }
        $ProviderMachineName = "" # SMS Provider machine name
        if ($null -eq (Get-PSDrive -Name $SiteCode -PSProvider CMSite -ErrorAction SilentlyContinue)) {
            New-PSDrive -Name $SiteCode -PSProvider CMSite -Root $ProviderMachineName
        }
        Set-Location ("$SiteCode" + ":") -ErrorAction Stop
        [array]$ConfigMgrDevices = Get-CMDevice -Name $computername -Fast -ErrorAction Stop
        IF ($NuLL -eq $ConfigMgrDevices) {
            LogIt "$Computername Not found in Configuration Manager!"
        }
        Else {
            LogIt "Successfully found $($ConfigMgrDevice.Name) in Configuration Manager." 
            foreach ($ConfigMgrDevice in $ConfigMgrDevices) {
                LogIt "   Deleting Name: $($ConfigMgrDevice.Name)  |  ResourceID: $($ConfigMgrDevice.ResourceID)  |  SMSID: $($ConfigMgrDevice.SMSID)  |  UserDomainName: $($ConfigMgrDevice.UserDomainName) ..."
                Remove-CMDevice -InputObject $ConfigMgrDevice -Force -ErrorAction Stop
                LogIt "Successfully Deleted $($ConfigMgrDevice.Name) from Configuraiton Manager."
            }
        }
    }
    Catch {
        LogIt "Error with Finding and Removing Device from Configuraiton Manager!"
        $_
        $err++
        $errCR = "$errCR`nIssue with Configuration Manager"
    }
}

# Navigate back to C drive (Leave MECM)
Set-Location $env:SystemDrive

# Check for Admin
if ((New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {                        
    # Remove from Active Directory
    if ($AD -or $Dispose) {
        Try {
            LogIt "Retrieving "  
            LogIt "Active Directory " 
            LogIt "computer account..."    
            $Searcher = [ADSISearcher]::new()
            $Searcher.Filter = "(sAMAccountName=$computername`$)"
            [void]$Searcher.PropertiesToLoad.Add("distinguishedName")
            $ComputerAccount = $Searcher.FindOne()
            If ($ComputerAccount) {
                LogIt "Successfully found $ComputerAccount in Active Directory."
                LogIt "   Deleting computer account..." 
                $DirectoryEntry = $ComputerAccount.GetDirectoryEntry(); $Result = $DirectoryEntry.DeleteTree()
                LogIt "Successfully Removed "
            }
            Else {
                LogIt "$Computername not found in Active Directory!"
            }
        }
        Catch {
            LogIt "Error with Finding or Deleting from Active Directory!"
            $_
            $err++
            $errCR = "$errCR`nIssue with Finding or Deleting from Active Directory"
        }
    }
    # Only Clear TPM and remove BIOS password if running locally. (Do not attempt to clear BIOS password or TPM when targetting different computer).
   
    <# Check for TPM
    LogIt "Checking TPM: " 
    if ((get-tpm).TpmReady) { 
        LogIt "TPM is currently configured"
        LogIt "Clearing TPM... " 
        # Clear TPM
        Clear-Tpm
        if (!((get-tpm).TpmReady)) {
            LogIt "TPM Successfully cleared."
        } 
        else {
            LogIt "Error Clearing TPM, please Clear-TPM manually as administrator."
        }
    } 
    else { 
        LogIt "TPM is unconfigured"
    }#>
}

#Dump all errors to end of log
if ($errCR) {
    LogIt "$errCR"
}

#Copy log up to network share
$NetworkLog = "$NetworkFolder\$computername.log"
$LocalLog = $LocalLog
Try {
    Copy-item -Path "$LocalLog" -Destination "$NetworkFolder" -Force
    if (Test-Path "$NetworkLog") {
        "$(get-date -Format "yyyy-MM-d HH:mm:ss:fff zzz") - ========Log-Complete========" | Out-File -FilePath "$NetworkLog" -Encoding unicode -Append
    }
}
Catch {
    $err++
    LogIt "Trouble sending log to network share.`nPlease check $NetworkLog and\or $LocalLog"
    $errCR = "$errCR `nTrouble sending log to network share.`nPlease check $NetworkLog and\or $LocalLog"
}

#wait a little bit
start-sleep -Seconds 5
# Check Error Handling
if ($err -ge 1) {
    # Review logs and see what failed
    Write-Host -ForegroundColor Red -BackgroundColor Black "$(get-date -Format "yyyy-MM-d HH:mm:ss:fff zzz") - $err error(s) detected`n$errCR"
    "$(get-date -Format "yyyy-MM-d HH:mm:ss:fff zzz") - $err error(s) detected`n$errCR" | Out-File -FilePath "$NetworkLog" -Encoding unicode -Append
 
}
else {
    Write-Host -ForegroundColor Green "$(get-date -Format "yyyy-MM-d HH:mm:ss:fff zzz") - No Errors found. Great Success! Please proceede to wipe machine!"
    "$(get-date -Format "yyyy-MM-d HH:mm:ss:fff zzz") - No Errors found. Great Success! Please proceede to wipe machine!" | Out-File -FilePath "$NetworkLog" -Encoding unicode -Append
}
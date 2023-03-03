# Include for Windows Presentaiton Framework (WPF)
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName System.Windows.Forms
# XML for GUI
[xml]$XAMLWindow = @"
<Window
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        Title="CompRemover" Height="363" Width="242" ResizeMode="NoResize" WindowStartupLocation="CenterScreen" Topmost="True">
    <Grid Background="#FF233E66" Margin="0,0,0,0">
        <TextBox Name="txtComputerName" HorizontalAlignment="Left" Height="32" Margin="29,49,0,0" VerticalAlignment="Top" Width="180" FontWeight="Bold" FontSize="18"/>
        <TextBox Name="txtComputerNameConfirm" HorizontalAlignment="Left" Height="32" Margin="29,106,0,0" VerticalAlignment="Top" Width="180" FontWeight="Bold" FontSize="18"/>
        <Button Name="btnBegin" Content="Begin" HorizontalAlignment="Left" Margin="145,290,0,0" VerticalAlignment="Top"/>
        <Label Name="lblComputerName" Content="Computer Name:" HorizontalAlignment="Left" Height="44" Margin="31,5,0,0" VerticalAlignment="Top" Width="170" FontWeight="Bold" FontSize="20"/>
        <CheckBox Name="cbDiffCompName" Content="Different Computer Name?" HorizontalAlignment="Left" Margin="38,86,0,0" VerticalAlignment="Top"/>
        <CheckBox Name="cbAAD" Content="Azure Active Directory" HorizontalAlignment="Left" Margin="38,225,0,0" VerticalAlignment="Top" Visibility="Visible"/>
        <CheckBox Name="cbAP" Content="AutoPilot" HorizontalAlignment="Left" Margin="38,205,0,0" VerticalAlignment="Top" Visibility="Visible"/>
        <CheckBox Name="cbIntune" Content="Intune" HorizontalAlignment="Left" Margin="38,185,0,0" VerticalAlignment="Top" Visibility="Visible"/>
        <CheckBox Name="cbAD" Content="Active Directory" HorizontalAlignment="Left" Margin="38,245,0,0" VerticalAlignment="Top" Visibility="Visible"/>
        <CheckBox Name="cbMECM" Content="Configuraiton Manager" HorizontalAlignment="Left" Margin="38,265,0,0" VerticalAlignment="Top" Visibility="Visible"/>
        <CheckBox Name="cbAll" Content="All" HorizontalAlignment="Left" Margin="38,295,0,0" VerticalAlignment="Top" Visibility="Visible"/>
        <Button Name="btnClear" Content="Clear" HorizontalAlignment="Left" Margin="93,290,0,0" VerticalAlignment="Top"/>
        <Button Name="btnDone" Content="Exit" HorizontalAlignment="Left" Margin="93,290,0,0" VerticalAlignment="Top" Visibility="Hidden" />
        <TextBox Name="txtOutput" HorizontalAlignment="Left" Margin="245,10,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="375" Height="300" VerticalScrollBarVisibility="Visible"/>
        <Button Name="btnTest" Content="Test" HorizontalAlignment="Left" Margin="184,182,0,0" VerticalAlignment="Top"/>
        </Grid>
</Window>
"@
# Create the Window Object
$Reader = (New-Object System.Xml.XmlNodeReader $XAMLWindow)
$Window = [Windows.Markup.XamlReader]::Load( $Reader )

# Get Form Items
$txtCN = $window.FindName("txtComputerName")
$txtCNC = $window.FindName("txtComputerNameConfirm")
$txtO = $window.FindName("txtOutput")
$btnC = $window.FindName("btnClear")
$btnB = $window.FindName("btnBegin")
$btnD = $window.FindName("btnDone")
#$btnT = $window.FindName("btnTest")
$cbDC = $window.FindName("cbDiffCompName")
$cbI = $window.FindName("cbIntune")
$cbAP = $window.FindName("cbAP")
$cbAAD = $window.FindName("cbAAD")
$cbAD = $window.FindName("cbAD")
$cbA = $window.FindName("cbAll")
$cbCM = $window.FindName("cbMECM")


# Functions
# Logging Funciton
Function LogIt {
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory = $true)]
        [ValidateNotNullorEmpty()]
        [string]$string,
        [Parameter(Mandatory = $false)]
        $nnl
    )
    if ($nnL) {
        $txtO.Text = "$($txtO.Text) $string"
    }
    else {
        $txtO.Text = "$($txtO.Text) $string `n"
    }
    "$(get-date -Format "yyyy-MM-d HH:mm:ss:fff zzz") - $string" | Out-File -FilePath "$LocalLog" -Encoding unicode -Append
}
# Reset Funciton
function ResetForm {
    $txtCN.IsReadOnly = $true
    $txtO.IsReadOnly = $true
    $txtCN.Text = $env:ComputerName
    $txtCNC.Visibility = "Hidden"
    $btnB.Visibility = "Visible"
    $btnC.Visibility = "Visible"
    $btnC.Content = "Clear"
    $btnD.Visibility = "Hidden"
    $txtO.Text = $null
    $cbDC.IsChecked = $false
    $cbAP.IsChecked = $false
    $cbCM.IsChecked = $false
    $cbAD.IsChecked = $false
    $cbAAD.IsChecked = $false
    $cbA.IsChecked = $false
    $cbI.IsChecked = $false
    $Window.Width = "242"
    $global:ComputerName = $env:COMPUTERNAME
    $global:LocalLog = "C:\Temp\$global:ComputerName.log"
    $global:NetworkLog = "$NetworkFolder\$global:ComputerName.log"
}

# Set Starting Values
ResetForm
#Set-ExecutionPolicy -Scope Process -Force Unrestricted
$global:ComputerName = $env:COMPUTERNAME
$global:LocalLog = "C:\Temp\$global:ComputerName.log"
$NetworkFolder = "\\Server\share\Removals"
$global:NetworkLog = "$NetworkFolder\$global:ComputerName.log"
$LoggedOnUser = (qwinsta /SERVER:$env:COMPUTERNAME) -replace '\s{2,22}', ',' | ConvertFrom-Csv | Where-Object { $_ -like "*Acti*" } | Select-Object -ExpandProperty Username

# Create Log Location (If not exist, it should)
if (!(Test-Path "C:\Temp")) {
    New-Item -ItemType Directory -Path "C:\Temp"
}

<# Test Button... Do Stuff. Remember to move button object back into xml above and Enable btnT in Get Form Items
#<Button Name="btnTest" Content="Test" HorizontalAlignment="Left" Margin="184,182,0,0" VerticalAlignment="Top"/>
$btnT.add_click({
    $Window.Width = "650"
    LogIt "$($global:computername.ToUpper())"
    LogIt "$NetworkLog"
    #Copy log up to network share
    $NetworkLog = "$NetworkFolder\$global:computername.log"
    LogIt "$NetworkLog"
    LogIt "$global:NetworkLog"
    $LocalLog = $global:LocalLog
    #Copy-item -Path "$LocalLog" -Destination "$NetworkFolder" -Force
    Try {
    if (Test-Path "$NetworkLog") {
        "$(get-date -Format "yyyy-MM-d HH:mm:ss:fff zzz") - ========Log-Complete========" | Out-File -FilePath "$NetworkLog" -encoding unicode -Append
    }
}
 Catch {
        $err++
        $errCR = "$errCR `nTrouble sending log to network share.`nPlease check $NetworkLog and\or $LocalLog"
    }
    })#>


# Different Computer Name Check Box Checked
$cbDC.add_Checked({
        $txtCN.IsReadOnly = $false
        $txtCN.Text = $null
        $txtCNC.Text = $null
        $btnB.Visibility = "Hidden"
        $txtCNC.Visibility = "Visible"
        $pwbLB.Visibility = "Hidden"
        $lblLB.Visibility = "Hidden"
        $cbLB.Visibility = "Hidden"
        $pwbLB.Password = $null
        $cbLB.IsChecked = $false
    })
# Different Computer Name Check Box Unchecked
$cbDC.add_Unchecked({
        $txtCN.IsReadOnly = $true
        $txtCN.Text = $env:ComputerName
        $txtCNC.Visibility = "Hidden"
        $btnB.Visibility = "Visible"
        $global:ComputerName = $env:COMPUTERNAME
        $global:LocalLog = "C:\Temp\$global:ComputerName.log"
        $global:NetworkLog = "$NetworkFolder\$global:ComputerName.log"
    })


# Logic to uncheck All box if any options unchecked. Yes I probably could have made this a check box group instead. Didn't feel like it!
$cbI.add_Unchecked({
        $cbA.IsChecked = $false
    })
$cbAD.add_Unchecked({
        $cbA.IsChecked = $false
    })
$cbAAD.add_Unchecked({
        $cbA.IsChecked = $false
    })
$cbCM.add_Unchecked({
        $cbA.IsChecked = $false
    })
$cbAP.add_Unchecked({
        $cbA.IsChecked = $false
    })

    
# Compare text boxes when manually typing Comptuer Name
$txtCN.add_TextChanged({
        if (($($txtCN.Text.ToUpper()) -eq $($txtCNC.Text.ToUpper())) -and ($txtCN.Text.Length -ge 1)) {
            $btnB.Visibility = "Visible"
            $global:computername = $txtCN.Text
            $global:LocalLog = "C:\Temp\$global:ComputerName.log"
            $global:NetworkLog = "$NetworkFolder\$global:ComputerName.log"
        }
        else {
            $btnB.Visibility = "Hidden"
        }
    })

# Compare text boxes when manually typing Comptuer Name Confirm
$txtCNC.add_TextChanged({
        if (($txtCNC.Text.ToUpper() -eq $txtCN.Text.ToUpper()) -and ($txtCNC.Text.Length -ge 1)) {
            $btnB.Visibility = "Visible"
            $global:computername = $txtCNC.Text
            $global:LocalLog = "C:\Temp\$global:computername.log"
            $global:NetworkLog = "$NetworkFolder\ComputerName.log"
        }
        else {
            $btnB.Visibility = "Hidden"
        }
    })

# Clear Button, essentially reset form. Might be better off making this a funciton to call at certain points, such as resetting form for invalid computername detections.
$btnC.add_click({
        ResetForm
    })

# All Check Box is Checked
$cbA.add_Checked({
        $cbAP.IsChecked = $true
        $cbCM.IsChecked = $true
        $cbAD.IsChecked = $true
        $cbAAD.IsChecked = $true
        $cbI.IsChecked = $true
    })

#Done\Exit button is clicked. Close window    
$btnD.add_click({
        $Window.Close()
    })

# Begin Button, run scripts based on selection(s)
$btnB.add_click({
        $global:computername = $txtCN.Text
        $global:LocalLog = "C:\Temp\$global:computername.log"
        $global:NetworkLog = "$NetworkFolder\$global:computername.log"
        # Reset error count, which should already be 0
        $err = 0
        #$err++;$errCR = "$errCR`nThis is a test error Bwahahaha"
        # Check if Comptuer name box is blank. Already have logic in both text boxes to make sure text matches. So either button works, using top box, since this is also true for autopopulated on program launch for current hostname
        if ($null -eq $txtCN.Text) {
            $txtO.Text = "Computer Name can not be blank.`nPlease wait for app to reset, and try again."
            $Window.Width = "650"
            $txtCNC.Visibility = "Hidden"
            $btnB.Visibility = "Visible"
            $btnC.Visibility = "Visible"
            $btnD.Visibility = "Hidden"
            Start-Sleep -Seconds 5
            $cbDC.IsChecked = $false
            $cbLB.IsChecked = $false
            $Window.Width = "242"
        }

        # Check that at least one of the essential check boxes are checked
        elseif (!($cbI.IsChecked -or $cbAP.IsChecked -or $cbAAD.IsChecked -or $cbCM.IsChecked -or $cbAD.IsChecked -or $cbLB.IsChecked)) { 
            $txtO.Text = "You must select at least one check box.`nPlease wait for app to reset, and try again."
            $Window.Width = "650"
            $txtCNC.Visibility = "Hidden"
            $btnB.Visibility = "Visible"
            $btnC.Visibility = "Visible"
            $btnD.Visibility = "Hidden"
            Start-Sleep -Seconds 5
            $cbDC.IsChecked = $false
            $cbLB.IsChecked = $false
            $Window.Width = "242"
        } 
        else {
            # Set Comptuername to top textbox value
            $global:computername = $txtCN.Text
            # Resize window to show output to user
            $Window.Width = "650"
            # Swap out buttons
            $btnB.Visibility = "Hidden"
            $btnC.Visibility = "Hidden"
            $btnD.Visibility = "Visible"
        }
            # Make sure we are on the C: drive
            Set-Location $env:SystemDrive
            # Start Logging
            LogIt "$($global:computername.ToUpper())"
            LogIt "========Log-Begin========"

            $Selection = "Removing $global:computername from:"
            if ($cbI.IsChecked) { $Selection = "$Selection Intune," }
            if ($cbAP.IsChecked) { $Selection = "$Selection Auto Pilot," }
            if ($cbAAD.IsChecked) { $Selection = "$Selection Azure Active Directory," }
            if ($cbAD.IsChecked) { $Selection = "$Selection Active Directory," }
            if ($cbCM.IsChecked) { $Selection = "$Selection Configurtation Manager," }
            LogIt $($selection -replace ".$")
            Try {
                LogIt " Installing Dependencies..." -nnl $true
                Set-ExecutionPolicy -Scope Process -Force Unrestricted
                if ($(Get-PackageProvider | Where-Object { ($_.name -eq "NuGet") }).Version -lt "2.8.5.201") {
                    LogIt "NuGet not found or version below minimum needed. Installing\Updating...."
                    Install-PackageProvider -Name NuGet -MinimumVersion "2.8.5.201" -Scope CurrentUser -Force
                    LogIt "Compatible NuGet Installed"
                }
                Else {
                    LogIt "NuGet already installed and version is at least 2.8.5.201"
                }

                # Install Dependencies needed based on checked boxes
                if ($cbI.IsChecked -or $cbAP.IsChecked -or $cbAAD.IsChecked) {
                    if (Get-Module -ListAvailable -Name Microsoft.Graph.Intune) {
                        LogIt "Microsoft Graph Intune already Installed"
                    }
                    else {
                        Install-Module Microsoft.Graph.Intune -Scope CurrentUser -Force 
                        LogIt "Microsoft Graph Intune Installed"
                    }
                }

                if ($cbAAD.IsChecked) {
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
                        Install-Module -Name AzureADPreview -Scope CurrentUser -Force -AllowClobber 
                        LogIt "AzureADPreview Installed"
                    }#>

                    if (Get-Module -ListAvailable -Name msal.ps) {
                        LogIt "msal.ps already Installed"
                    }
                    else { 
                        Install-Package msal.ps -Scope CurrentUser -Force 
                        LogIt "msal.ps Installed"
                    }

                    if (Get-Module -ListAvailable -Name MSOnline) {
                        LogIt "MSOnline already Installed"
                    }
                    else { 
                        Install-Module -Name MSOnline -Scope CurrentUser -Force -ErrorAction SilentlyContinue 
                        LogIt "MSOnline Installed"
                    }
                }
                LogIt "Success"
            }
            Catch {
                LogIt "$($_.Exception.Message)" 
                $errCR = "$errCR`nIssue with Dependencies"
                $err++
                Return
            }
        
            # Load required modules based on checked boxes
            if ($cbI.IsChecked -or $cbAP.IsChecked -or $cbAAD.IsChecked -or $cbCM.IsChecked) {
                Try {
                    # Cloud (Intune, Azure AD, or Auto Pilot) Only
                    if ($cbI.IsChecked -or $cbAP.IsChecked -or $cbAAD.IsChecked) {
                        If ("Microsoft.Graph.Intune" -in $(Get-Module).Name) {
                            LogIt "Microsoft Graph Intune already Imported"
                        }
                        Else {
                            Import-Module Microsoft.Graph.Intune -Force -ErrorAction Stop
                            LogIt "Microsoft Graph Intune successfully Imported"
                        }
                        Try {
                            LogIt "Logged on user is: $LoggedOnUser"
                            # Create the ACE
                            $identity = "NASDCORP\$LoggedOnUser"
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
                            $errCR = "$errCR`nUnable to grant permissions of $env:USERPROFILE to $LoggedOnUser"
                            $err++
                            Continue
                        }
                    }
                    # Azure Active Directory Only
                    if ($cbAAD.IsChecked) {
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
                            Import-Module AzureADPreview -Force -ErrorAction Stop
                            LogIt "AzureADPreview successfully Imported"
                        }#>
                    }
                    # Config Manager only
                    if ($cbCM.IsChecked) {
                        If ("ConfigurationManager" -in $(Get-Module).Name) {
                            LogIt "ConfigurationManager already Imported"
                        }
                        Else {
                            #Import-Module $env:SMS_ADMIN_UI_PATH.Replace('i386', 'ConfigurationManager.psd1') -Force -ErrorAction Stop
                            #Import-Module "\\server\share$\Program Files\Microsoft Configuration Manager\AdminConsole\bin\ConfigurationManager\ConfigurationManager.psd1"  -Force -ErrorAction Stop
                            Import-Module "\\server\ConfigurationManager_bin\ConfigurationManager.psd1" -Force -ErrorAction Stop # Might need to change this based on location and permission.
                            LogIt "ConfigurationManager successfully Imported"
                        }
                    }
                    LogIt "Success"
                }
                Catch {
                    LogIt "$($_.Exception.Message)"
                    $err++
                    $errCR = "$errCR`nIssue with Modules"
                }
            }

            # Cloud Conneciton
            if ($cbI.IsChecked -or $cbAP.IsChecked -or $cbAAD.IsChecked) {
                Try {
                    LogIt "Authenticating with MS Graph and Azure AD..." -nnl $true
                    #Authenticate to Cloud stuffs
                    $intuneId = Connect-MSGraph -ErrorAction Stop
                    $aadId = Connect-AzureAD -AccountId $intuneId.UPN -ErrorAction Stop
                    #Connect-MgGraph -ClientId "62369bec-7731-4576-8806-766ace939ad0" -TenantId "79a2fbbe-f26b-4ab8-ab20-5223e4c2f62d" #Sample App in Dev
                    $TenantId = $aadId.TenantId
                    $AccountId = $(Get-AzureADUser -ObjectId $($intuneId.UPN)).ObjectId
                    LogIt "Success"

                    # Elevate Self with PIM (Should prompt MFA\2FA)

                    # Set schedule and Reason for PIM elevation (Will appear in Azure logs and notify watchers)
                    $Schedule = New-Object Microsoft.Open.MSGraph.Model.AzureADMSPrivilegedSchedule
                    $Schedule.Type = "Once"
                    $Schedule.StartDateTime = (Get-Date).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ss.fffZ")
                    $Schedule.endDateTime = ((Get-Date).AddMinutes(10)).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ss.fffZ")
                    $Reason = "$($intuneId.UPN): Removing Device - $ComputerName"

                    # Fetch all PIM role assignments for the current user.
                    $AzureADMSPrivilegedRoleDefinition = Get-AzureADMSPrivilegedRoleDefinition -ProviderId 'aadRoles' -ResourceId $TenantId
                    # Fetch IDs for desired roles
                    $Roles = 'Cloud Device Administrator' #,'Intune Administrator'#, 'Azure AD Joined Device Local Administrator'
                    foreach ($role in $Roles) {
                        $RoleID = ($AzureADMSPrivilegedRoleDefinition | Where-Object { $role -eq $_.DisplayName }).Id
                        Try {
                            LogIt $(Open-AzureADMSPrivilegedRoleAssignmentRequest -ProviderId 'aadRoles' -ResourceId $TenantId -RoleDefinitionId $RoleID -SubjectId $AccountId -Type 'UserAdd' -AssignmentState 'Active' -Schedule $Schedule -Reason $Reason)
                        }
                        Catch {
                            $err++
                            LogIt "Error with PIM $_"
                            $errCR = "$errCR`nIssue with $role PIM"
                        }
                    }
                    Start-Sleep -Seconds 25
                    LogIt "Waiting a little bit for PIM"
                    # Get Intune device data, this is needed to remove from AutoPilot devices later, and/or used to delete Intune device
                    if ($cbI.IsChecked -or $cbAP.IsChecked) {
                        Try {
                            LogIt "Retrieving " -nnl $true
                            LogIt "Intune " -nnl $true
                            LogIt "managed device record/s..." -nnl $true
                            [array]$IntuneDevices = Get-IntuneManagedDevice -Filter "deviceName eq '$global:computername'" -ErrorAction Stop
                            If ($IntuneDevices.Count -ge 1) {
                                LogIt "Success"
                                # Remove From Intune
                                If ($cbI.IsChecked) {
                                    foreach ($IntuneDevice in $IntuneDevices) {
                                        LogIt "   Deleting DeviceName: $($IntuneDevice.deviceName)  |  Id: $($IntuneDevice.Id)  |  AzureADDeviceId: $($IntuneDevice.azureADDeviceId)  |  SerialNumber: $($IntuneDevice.serialNumber) ..." -nnl $true
                                        Remove-IntuneManagedDevice -managedDeviceId $IntuneDevice.Id -Verbose -ErrorAction Stop
                                        LogIt "Success"
                                    }
                                }
                            }
                            Else {
                                LogIt "Not found!"
                            }
                        }
                        Catch {
                            LogIt "Error with Finding or Deleting Intune Device!"
                            $errCR = "$errCR`nIssue with Intune"
                            $err++
                            $_
                        }
                        # Remove from Auto Pilot
                        if ($cbAP.IsChecked) {
                            If ($IntuneDevices.Count -ge 1) {
                                Try {
                                    LogIt "Retrieving " -nnl $true
                                    LogIt "Autopilot "  -nnl $true
                                    LogIt "device registration..." -nnl $true
                                    $AutopilotDevices = New-Object System.Collections.ArrayList
                                    foreach ($IntuneDevice in $IntuneDevices) {
                                        $URI = "https://graph.microsoft.com/beta/deviceManagement/windowsAutopilotDeviceIdentities?`$filter=contains(serialNumber,'$($IntuneDevice.serialNumber)')"
                                        $AutopilotDevice = Invoke-MSGraphRequest -Url $uri -HttpMethod GET -ErrorAction Stop
                                        [void]$AutopilotDevices.Add($AutopilotDevice)
                                    }
                                    LogIt "Success"
                        
                                    foreach ($device in $AutopilotDevices) {
                                        LogIt "   Deleting SerialNumber: $($Device.value.serialNumber)  |  Model: $($Device.value.model)  |  Id: $($Device.value.id)  |  GroupTag: $($Device.value.groupTag)  |  ManagedDeviceId: $($device.value.managedDeviceId) ..."
                                        $URI = "https://graph.microsoft.com/beta/deviceManagement/windowsAutopilotDeviceIdentities/$($device.value.Id)"
                                        $AutopilotDevice = Invoke-MSGraphRequest -Url $uri -HttpMethod DELETE -ErrorAction Stop
                                        LogIt "Success" 
                                    }
                                }
                                Catch {
                                    LogIt "Error!"
                                    $err++
                                    $errCR = "$errCR`nIssue with AutoPilot"
                                    $_
                                }
                            }
                        }
                    }
                }
                Catch {
                    LogIt "Error!"
                    LogIt "$($_.Exception.Message)"
                    $err++
                    $errCR = "$errCR`nIssue with Cloud Connection"
                    Return
                }
            }

            #Remove From Azure AD
            if ($cbAAD.IsChecked) {
                Try {
                    LogIt "Retrieving " -nnl $true
                    LogIt "Azure AD "  -nnl $true
                    LogIt "device record/s..."  -nnl $true
                    [array]$AzureADDevices = Get-AzureADDevice -SearchString $global:computername -All:$true -ErrorAction Stop
                    If ($AzureADDevices.Count -ge 1) {
                        LogIt "Success" 
                        Foreach ($AzureADDevice in $AzureADDevices) {
                            if ($global:computername -eq $txtCN.Text) {
                                LogIt "   Deleting Hostname: $($AzureADDevice.DisplayName)  |  ObjectId: $($AzureADDevice.ObjectId)  |  DeviceId: $($AzureADDevice.DeviceId) ..."
                                Remove-AzureADDevice -ObjectId $AzureADDevice.ObjectId -ErrorAction Stop
                                LogIt "Success" 
                            }
                        }      
                    }
                    Else {
                        LogIt "Not found!"
                    }
                }
                Catch {
                    LogIt "Error!" 
                    $err++
                    $errCR = "$errCR`nIssue with Azure Active Directory"
                    $_
                }
            }
            # Remove from Configuraiton Manager
            if ($cbCM.IsChecked) {
                Try {
                    LogIt "Retrieving " -nnl $true
                    LogIt "ConfigMgr " -nnl $true
                    LogIt "device record/s..." -nnl $true
                    $SiteCode = (Get-PSDrive -PSProvider CMSITE -ErrorAction SilentlyContinue).Name
                    if ($null -eq $SiteCode) { $SiteCode = "SMS" }
                    $ProviderMachineName = "SMS" # SMS Provider machine name
                    if ($null -eq (Get-PSDrive -Name $SiteCode -PSProvider CMSite -ErrorAction SilentlyContinue)) {
                        New-PSDrive -Name $SiteCode -PSProvider CMSite -Root $ProviderMachineName
                    }
                    Set-Location ("$SiteCode" + ":") -ErrorAction Stop
                    [array]$ConfigMgrDevices = Get-CMDevice -Name $global:computername -Fast -ErrorAction Stop
                    LogIt "Success" 
                    foreach ($ConfigMgrDevice in $ConfigMgrDevices) {
                        LogIt "   Deleting Name: $($ConfigMgrDevice.Name)  |  ResourceID: $($ConfigMgrDevice.ResourceID)  |  SMSID: $($ConfigMgrDevice.SMSID)  |  UserDomainName: $($ConfigMgrDevice.UserDomainName) ..."
                        Remove-CMDevice -InputObject $ConfigMgrDevice -Force -ErrorAction Stop
                        LogIt "Success"
                    }
                }
                Catch {
                    LogIt "Error!"
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
                if ($cbAD.IsChecked) {
                    Try {
                        LogIt "Retrieving " -nnl $true 
                        LogIt "Active Directory " -nnl $true
                        LogIt "computer account..."    -nnl $true
                        $Searcher = [ADSISearcher]::new()
                        $Searcher.Filter = "(sAMAccountName=$global:computername`$)"
                        [void]$Searcher.PropertiesToLoad.Add("distinguishedName")
                        $ComputerAccount = $Searcher.FindOne()
                        If ($ComputerAccount) {
                            LogIt "Success"
                            LogIt "   Deleting computer account..." -nnl $true
                            $DirectoryEntry = $ComputerAccount.GetDirectoryEntry(); $Result = $DirectoryEntry.DeleteTree()
                            LogIt "Success"
                        }
                        Else {
                            LogIt "Not found!"
                        }
                    }
                    Catch {
                        LogIt "Error!"
                        $_
                        $err++
                        $errCR = "$errCR`nIssue with Active Directory"
                    }
                }

            #Dump all errors to end of log
            if ($errCR) {
                LogIt "$errCR"
            }


            #Copy log up to network share
            $NetworkLog = "$NetworkFolder\$global:computername.log"
            $LocalLog = $global:LocalLog
            Try {
                Copy-item -Path "$LocalLog" -Destination "$NetworkFolder" -Force
                if (Test-Path "$NetworkLog") {
                    "$(get-date -Format "yyyy-MM-d HH:mm:ss:fff zzz") - ========Log-Complete========" | Out-File -FilePath "$NetworkLog" -encoding unicode -Append
                }
            }
            Catch {
                $err++
                $errCR = "$errCR `nTrouble sending log to network share.`nPlease check $NetworkLog and\or $LocalLog"
            }

            #wait a little bit
            start-sleep -Seconds 5
            # Check Error Handling
            if ($err -ge 1) {
                # Review logs and see what failed
                $txtO.Text += "$err error(s) detected`n$errCR"
                $btnC.Visibility = "Visible"
                $btnC.Content = "Reset"
            }
            else {
                if ($cbDC.IsChecked) {
                    # Script ran successfully
                    $txtO.Text = "$($txtO.Text) `nNo Errors found. Great Success! Please make sure to remove BIOS password and clear TPM of target comptuer."
                }
                else {
                    # Script ran successfully, time to reimage
                    $txtO.Text = "$($txtO.Text) `nNo Errors found. Great Success! Please proceede to reimage machine!"
                }
                $btnD.Visibility = "Visible"
            }
        }

    })
$Window.ShowDialog() | Out-Null
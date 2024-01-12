[CmdletBinding()]
Param
(
    [Parameter(Mandatory = $false)]
    $aPRD,
    [Parameter(Mandatory = $false)]
    $aDK,
    [Parameter(Mandatory = $false)]
    $aKey
)
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName System.Windows.Forms
# XML for GUI
[xml]$XAMLWindow = @"
<Window
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        Title="Summon_IDE" Height="450" Width="330" ResizeMode="CanMinimize" WindowStartupLocation="CenterScreen">
        <Grid Background="#FFA39161">
        <TextBlock Name="tbStatus" HorizontalAlignment="Left" Margin="30,40,0,0" TextWrapping="Wrap" Text=" " VerticalAlignment="Top" Height="45" Width="262" Background="#33000000" FontSize="10"/>
        <Button Name="btnReset" Content="Reset" HorizontalAlignment="Left" Margin="30,372,0,0" VerticalAlignment="Top"/>
        <Button Name="btnEXit" Content="Exit" HorizontalAlignment="Left" Margin="72,372,0,0" VerticalAlignment="Top"/>
        <TextBox Name="txtPRD" HorizontalAlignment="Left" Margin="30,127,0,0" TextWrapping="Wrap" Text="" VerticalAlignment="Top" Width="237" Height="25" Visibility="Visible"/>
        <Label Name="lblPRD" Content="Enter Project Root Directory:" HorizontalAlignment="Left" Margin="30,96,0,0" VerticalAlignment="Top" Visibility="Hidden"/>
        <Button Name="btnSet" Content="Set" HorizontalAlignment="Left" Margin="149,164,0,0" VerticalAlignment="Top" Visibility="Hidden"/>
        <Label Name="lblAppKey" Content="Enter AppKey/DomainKey for the app (e.g. lcapx, quick-start):" HorizontalAlignment="Left" Margin="30,96,0,0" VerticalAlignment="Top" Visibility="Hidden" FontSize="10"/>
        <TextBox Name="txtAppKey" HorizontalAlignment="Left" Margin="30,127,0,0" VerticalAlignment="Top" Width="262" Height="25" Visibility="Hidden"/>
        <Button Name="btnAppKey" Content="Set" HorizontalAlignment="Left" Margin="149,164,0,0" VerticalAlignment="Top" Visibility="Hidden"/>
        <Button Name="btnCyberark" Content="Open CyberArk Portal" HorizontalAlignment="Left" Margin="123,164,0,0" VerticalAlignment="Top" ToolTip="https://<domain>privilegecloud.cyberark.com" Visibility="Hidden"/>
        <TextBox Name="txtApiKey" Background="Transparent" BorderThickness="0" IsReadOnly="True" Text="Enter password for username user (Please fetch password from CyberArk Password Vault):" TextWrapping="Wrap" HorizontalAlignment="Left" Margin="30,78,0,0" VerticalAlignment="Top" Width="262" Height="44" FontSize="10"/>
        <PasswordBox Name="pwbApiKey" HorizontalAlignment="Left" Margin="30,127,0,0" VerticalAlignment="Top" Width="262" Height="25" Visibility="Hidden"/>
        <Button Name="btnApiKey" Content="Set" HorizontalAlignment="Left" Margin="271,164,0,0" VerticalAlignment="Top" Visibility="Hidden"/>
        <ListBox Name="lstIDE" HorizontalAlignment="Left" Margin="30,127,0,0" Width="262" Height="232" VerticalAlignment="Top" Visibility="Hidden"/>
        <TextBox Name="txtError" HorizontalAlignment="Left" Margin="30,127,0,0" Width="262" Height="232" VerticalAlignment="Top" Visibility="Hidden" Text="Error" TextWrapping="Wrap"/>
        <Button Name="btnDebug" Content="Debug" HorizontalAlignment="Left" Margin="275,372,0,0" VerticalAlignment="Top"/>
        <Button Name="btnCopyUsername" Content="Copy Username" HorizontalAlignment="Left" Margin="30,164,0,0" VerticalAlignment="Top" Visibility="Visible"/>
        <Button Name="btnCopyAndLaunch" Content="Copy &amp; Launch" HorizontalAlignment="Left" Margin="170,372,0,0" VerticalAlignment="Top" Visibility="Hidden"/>
        <TextBox Name="txtDebug" HorizontalAlignment="Left" Margin="325,46,0,0" Text="" VerticalAlignment="Top" Height="346" Width="302" ScrollViewer.HorizontalScrollBarVisibility="Auto" ScrollViewer.VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Auto" VerticalScrollBarVisibility="Auto"/>
        <Button Name="btnLaunch" Content="Launch" HorizontalAlignment="Left" Margin="225,372,0,0" VerticalAlignment="Top" Visibility="Visible"/>
        <Button Name="btnLaC" Content="Launch &amp; Close" HorizontalAlignment="Left" Margin="132,372,0,0" VerticalAlignment="Top" Visibility="Visible"/>
        <Button Name="btnBrowse" Content="..." Margin="272,130,0,0" HorizontalAlignment="Left" VerticalAlignment="Top" Width="20" Height="20" Visibility="Hidden"/>
    </Grid>
</Window>
"@
# Create the Window Object
$Reader = (New-Object System.Xml.XmlNodeReader $XAMLWindow)
$Window = [Windows.Markup.XamlReader]::Load( $Reader )

## Variables

# Get Form Items
# Button
$buttonReset = $window.FindName("btnReset")
$buttonSet = $window.FindName("btnSet")
$buttonExit = $window.FindName("btnEXit")
$buttonAppKey = $window.FindName("btnAppKey")
$buttonCyberark = $window.FindName("btnCyberark")
$buttonDebug = $window.FindName("btnDebug")
$buttonApiKey = $window.FindName("btnApiKey")
$buttonLaunch = $window.FindName("btnLaunch")
$buttonLaunchAndClose = $window.FindName("btnLaC")
$buttonBrowseDir = $window.FindName("btnBrowse")
$buttonCopyUsername = $window.FindName("btnCopyUsername")
$btnCopyAndLaunch = $window.FindName("btnCopyAndLaunch")

# TextBox
$global:txtBoxPRD = $window.FindName("txtPRD")
$txtBoxDebug = $window.FindName("txtDebug")
$txtBoxAppKey = $window.FindName("txtAppKey")
$txtBoxApiKey = $window.FindName("txtApiKey")
$global:txtBoxError = $window.FindName("txtError")

# Label
$labelPRD = $window.FindName("lblPRD")
$labelAppKey = $window.FindName("lblAppKey")

#Password Box
$pwApiKey = $window.FindName("pwbApiKey")

# List
$ideList = $window.FindName("lstIDE")

#TextBlock
$textBoxStatus = $window.FindName("tbStatus")


# Local
$version = "v1.0.1.8"
$summonPath = $($env:path -split (';')) -like "*Cyberark Conjur\Summon*"
$hDir = "C:\Users\$((qwinsta /SERVER:$env:COMPUTERNAME) -replace '\s{2,22}', ',' | ConvertFrom-Csv | Where-Object { $_ -like "*Acti*" } | Select-Object -ExpandProperty Username)"
[array]$instIDEs = @(
    [pscustomobject]@{IDE = "CMD"; path = "C:\WINDOWS\system32\cmd.exe" },
    [pscustomobject]@{IDE = "PowerShell"; path = "C:\WINDOWS\System32\WindowsPowerShell\v1.0\powershell.exe" },
    [pscustomobject]@{IDE = "IntelliJ IDEA"; path = $(Get-Item "C:\Program Files*\JetBrains\IntelliJ IDEA*\bin\idea64.exe" -ErrorAction SilentlyContinue).FullName },
    [pscustomobject]@{IDE = "IntelliJ IDEA"; path = $(Get-Item "$hDir\AppData\Local\JetBrains\IntelliJ IDEA*\bin\idea64.exe" -ErrorAction SilentlyContinue).FullName },
    [pscustomobject]@{IDE = "PyCharm"; path = $(Get-Item "C:\Program Files*\JetBrains\PyCharm *\bin\pycharm64.exe" -ErrorAction SilentlyContinue).FullName },  
    [pscustomobject]@{IDE = "VSCode"; path = $(Get-Item "$hDir\AppData\Local\Programs\Microsoft VS Code\Code.exe" -ErrorAction SilentlyContinue).FullName }
    [pscustomobject]@{IDE = "VSCode"; path = $(Get-Item "$hDir\AppData\Local\Programs\Microsoft VS Code\bin\Code.cmd" -ErrorAction SilentlyContinue).FullName }
    [pscustomobject]@{IDE = "VSCode"; path = $(Get-Item "C:\Program Files*\Microsoft VS Code\Code.exe" -ErrorAction SilentlyContinue).FullName }
    [pscustomobject]@{IDE = "VSCode"; path = $(Get-Item "C:\Program Files*\Microsoft VS Code\bin\Code.cmd" -ErrorAction SilentlyContinue).FullName }
    [pscustomobject]@{IDE = "VSCode"; path = $(Get-Item "C:\Program Files*\Microsoft VS Code\*\bin\Code.cmd" -ErrorAction SilentlyContinue).FullName }
    [pscustomobject]@{IDE = "VSCode"; path = $(Get-Item "C:\Program Files*\Microsoft VS Code\*\Code.exe" -ErrorAction SilentlyContinue).FullName }
)
# Remove empty path objects of array, assumed not installed IDE.
Foreach ($IDE in $instIDEs) {
    if ($null -ne $IDE.path) {
        if ($IDE.path.count -eq 1) {
            $ideList.Items.Add($("$($IDE.IDE): $($IDE.path)")) | Out-Null
        }
        else {
            ForEach ($Path in $($IDE.path)) {
                $ideList.Items.Add(  $("$($IDE.IDE): $($Path)")) | Out-Null
            }
        }
    }
}

## Functions
# Logging Funciton
Function LogIt {
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory = $true)]
        [ValidateNotNullorEmpty()]
        [string]$string,
        [Parameter(Mandatory = $false)]
        $nnl
    )<#
    if ($nnL) {
        $txtO.Text = "$($txtO.Text) $string"
    }
    else {
        $txtO.Text = "$($txtO.Text) $string `n"
    }#>
    $txtBoxDebug.text += "$(Get-Date -Format "yyyy-MM-d HH:mm:ss:fff zzz") - $string`n" #| Out-File -FilePath "$LocalLog" -Encoding unicode -Append
}
# Reset Funciton
function ResetForm {
    $txtBoxDebug.text | Out-File -FilePath "C:\temp\Summon_IDE-debug.log" -Encoding unicode -Append
    if ($summonPath) {
        LogIt "Found $summonPath"
        $textBoxStatus.Text = " Summon environment has been set."
        $textBoxStatus.ToolTip = "Retrieved from system Path."
        $textBoxStatus.Background = "#008337"
        $global:txtBoxPRD.Visibility = "Visible"
        $buttonBrowseDir.Visibility = "Visible"
        $labelPRD.Visibility = "Visible"            
        $buttonSet.Visibility = "Visible"
        $txtBoxPRD.ToolTip = "Enter the path for a local folder conatining your project and the secrets.yml file."
    }
    else {
        LogIt "Summon not found in path."
        $textBoxStatus.Height = "45"
        $textBoxStatus.Text = ' Summon CLI doesnt exist under system path. It is usually available under: "C:\Program Files\Cyberark Conjur\Summon". Please verify and try again!'
        $textBoxStatus.ToolTip = "Check Software Center or contact Tech Support if needed." 
        $textBoxStatus.Background = "#ab2f42"
        $buttonBrowseDir.Visibility = "Hidden"
        $labelPRD.Visibility = "Hidden"
        $buttonSet.Visibility = "Hidden"
    }
    $ideList.UnselectAll()
    $Window.Width = "331"
    $global:AppKey = $null
    $global:txtBoxPRD.Text = $null
    $txtBoxDebug.text = $null
    $txtBoxAppKey.Text = $null
    $global:txtBoxError.Text = $null
    $pwApiKey.Password = $null
    $buttonAppKey.Visibility = "Hidden"
    $buttonCyberark.Visibility = "Hidden"
    $buttonCopyUsername.Visibility = "Hidden"
    $btnCopyAndLaunch.Visibility = "Hidden"
    $txtBoxAppKey.Visibility = "Hidden"
    $global:txtBoxError.Visibility = "Hidden"
    $labelAppKey.Visibility = "Hidden"
    $pwApiKey.Visibility = "Hidden"
    $txtBoxApiKey.Visibility = "Hidden"
    $buttonApiKey.Visibility = "Hidden"
    $buttonLaunch.Visibility = "Hidden"
    $buttonLaunchAndClose.Visibility = "Hidden"
    $ideList.Visibility = "Hidden"
    $buttonDebug.Visibility = "Visible"
    $textBoxStatus.Height = "25"
    # Set passwed arguments
    if ($aPRD) {
        $global:txtBoxPRD.Text = "$aPRD"
    }
    if ($aDK) {
        $txtBoxAppKey.Text = "$aDK"
    }
    if ($aKey) {
        $pwApiKey.Password = "$aKey"
    }

}


# Set Button
$buttonSet.add_click({
        LogIt "Pressed Button Set"
        if ($global:txtBoxPRD.Text.Length -le 0) {
            LogIt "Textbox Project Root Directory Empty"
            $textBoxStatus.Background = "#ab2f42"
            $textBoxStatus.Text = " Must enter path for Project Root Directory."
        }
        else {
            LogIt "Textbox Project Root Directory Accepted"
            $global:PRD = $global:txtBoxPRD.Text
            LogIt "$global:PRD.text entered for Project Root Directory"
            if (Test-Path $global:PRD -ErrorAction SilentlyContinue) {
                LogIt "Textbox Project Root Directory path found"
                $global:Yml = "$global:PRD\secrets.yml"
                try {
                    if (Test-Path $global:Yml) {
                        LogIt "Secrets.yml found in $global:PRD"
                        $textBoxStatus.Background = "#008337"
                        $textBoxStatus.Text = " Found Project Directory Root with secret.yml."
                        $labelAppKey.Visibility = "Visible"
                        $labelPRD.Visibility = "Hidden"
                        $global:txtBoxPRD.Visibility = "Hidden"
                        $buttonBrowseDir.Visibility = "Hidden"
                        $buttonSet.Visibility = "Hidden"
                        $txtBoxAppKey.Visibility = "Visible"
                        $buttonAppKey.Visibility = "Visible"
                        $textBoxStatus.ToolTip = "Enter the AppKey below, may also be referred to as a Domain Key" 
                    }
                    else {
                        LogIt "secrets.yml not found in $global:PRD"
                        $textBoxStatus.Background = "#ab2f42"
                        $textBoxStatus.Text = " Specified path exists, Need secrets.yml in root."
                    }
                }
                catch {
                    LogIt "An error has occured with $global:Yml"
                    $textBoxStatus.Background = "#ab2f42"
                    $textBoxStatus.Text = " Specified path exists, Need secrets.yml in root."
                }    
            }
            else {
                LogIt "$global:PRD not found on this computer"
                $textBoxStatus.Background = "#ab2f42"
                $textBoxStatus.Text = " Specified project Directory Root not found."
            }
        }
    })
    
# Cyberark button
$buttonCyberark.add_click({
        LogIt "Cyberark button pressed. Opening in default browser"
        Start-Process "https://<domain>.cyberark.com"
    })

# Copy Username Button
$buttonCopyUsername.add_click({
        LogIt "Copy username button pressed. Writing AUTHN_LOGIN: $env:AUTHN_LOGIN to Clipboard"
        Set-Clipboard -Value "$env:AUTHN_LOGIN"
    })

# AppKey Button
$buttonAppKey.add_Click({
        LogIt "Button AppKey pressed"
        if ($txtBoxAppKey.Text.Length -le 0) {
            LogIt "Submitted AppKey password detected as blank"
            $textBoxStatus.Background = "#ab2f42"
            $textBoxStatus.Text = " Cannot leave AppKey blank."
        }
        else {
            LogIt "AppKey is: $($txtBoxAppKey.Text)"
            $global:AppKey = $txtBoxAppKey.Text
            $textBoxStatus.Background = "#008337"
            $textBoxStatus.Text = " AppKey/DomainKey has been set."
            $env:AUTHN_LOGIN = "k8s-apps/$global:AppKey/$global:AppKey-localdev"
            $env:CONJUR_AUTHN_LOGIN = "host/$env:AUTHN_LOGIN"
            LogIt "AUTHN_LOGIN set to: $env:AUTHN_LOGIN"
            LogIt "CONJUR_AUTHN_LOGIN set to:  $env:CONJUR_AUTHN_LOGIN"
            $buttonCyberark.Visibility = "Visible"
            $buttonCopyUsername.Visibility = "Visible"
            $labelAppKey.Visibility = "Hidden"
            $txtBoxAppKey.Visibility = "Hidden"
            $buttonAppKey.Visibility = "Hidden"
            $pwApiKey.Visibility = "Visible"
            $txtBoxApiKey.Visibility = "Visible"
            $txtBoxApiKey.Text = "Enter password for username $env:AUTHN_LOGIN (Please fetch password from CyberArk Password Vault):"
            $buttonApiKey.Visibility = "Visible"
            $textBoxStatus.ToolTip = "Retrieve API Key for this App. Cyberark button below can be used if needed."
        }
    })

# API Button
$buttonApiKey.add_Click({
        LogIt "Button API pressed"
        if ($pwApiKey.Password.Length -le 0) {
            LogIt "Submitted API Key detected as blank"
            $textBoxStatus.Background = "#ab2f42"
            $textBoxStatus.Text = " Cannot leave API Key blank."
        }
        else {
            LogIt "API Key password is: $($pwApiKey.Password.length) characters long"
            $textBoxStatus.Text = " Select desired IDE below"
            $textBoxStatus.Background = "#33000000"
            $AUTHN_API_KEY = $pwApiKey.Password
            $env:CONJUR_AUTHN_API_KEY = $AUTHN_API_KEY 
            $pwApiKey.Visibility = "Hidden"
            $txtBoxApiKey.Visibility = "Hidden"
            $buttonApiKey.Visibility = "Hidden"
            $buttonCyberark.Visibility = "Hidden"
            $buttonCopyUsername.Visibility = "Hidden"
            $ideList.Visibility = "Visible"
            $textBoxStatus.ToolTip = "Make a selection from the list below, to launch application with modified summon command line."
        }
    })

# Copy & Launch Button
$btnCopyAndLaunch.add_click({
        LogIt "Copy and launch button clicked.."
        LogIt "Writing IDE path to clipboard: '`"$LAUNCH_SUB_CMD`" .'"
        Set-Clipboard -Value "`"$LAUNCH_SUB_CMD`" ."
        $global:ideName = "CMD"
        $global:LAUNCH_CMD = "summon -f `'$yml`'  C:\WINDOWS\system32\cmd.exe"
        LogIt "summoning command prompt."
        Launch
    })

# Launch & Close Button
$buttonLaunchAndClose.add_click({
        if ($global:ideName -in "IntelliJ IDEA", "PyCharm") {
            $global:txtBoxError.Visibility = "Visible"
            $textBoxStatus.Background = "#fce205"
            $textBoxStatus.Text = " Warning $global:ideName selected. This launcher will not be accessible until you close that window."
        }
        Launch
        Start-Sleep -Seconds 4
        $Window.Close()
    })
# Event handler for clicking on itmes in list. Populate variables
$ideList.Add_SelectionChanged({
        $global:ideName = $($ideList.Items[$ideList.SelectedIndex.ToString()].split(':'))[0] # Everything before ':'
        $global:idePath = $($ideList.Items[$ideList.SelectedIndex.ToString()]) -replace ('^.*?: ', '') # Everything after the first ':'
        Logit "Selected $global:ideName located at: $global:idePath"

        $global:LAUNCH_SUB_CMD = $global:idePath
        Logit "Set LAUNCH_SUB_COMMAND to: $LAUNCH_SUB_CMD"
        $global:LAUNCH_CMD = "summon -f `'$yml`'  `'$LAUNCH_SUB_CMD`'"
        Logit "Set LAUNCH_COMMAND to: $LAUNCH_CMD"
        #$textBoxStatus.ToolTip = "Launch command currently set to '$LAUNCH_CMD'"
        if ($LAUNCH_CMD.Length -gt '7') {
            $buttonLaunch.Visibility = "Visible"
            $buttonLaunchAndClose.Visibility = "Visible"
        }
        if ($global:ideName -in "PowerShell", "CMD") {
            $textBoxStatus.ToolTip = "Start-Process $global:ideName -ArgumentList $global:LAUNCH_CMD -WorkingDirectory $global:PRD -PassThru"
            Logit "Command Line set to: Start-Process $global:ideName -ArgumentList $global:LAUNCH_CMD -WorkingDirectory $global:PRD -PassThru"
        }
        if ($global:ideName -in "IntelliJ IDEA", "PyCharm", "VSCode") {
            $textBoxStatus.ToolTip = "powershell summon -f `'$yml`' `'$LAUNCH_SUB_CMD`' `'$global:PRD`'"
            Logit "Command Line set to: powershell summon -f `'$yml`' `'$LAUNCH_SUB_CMD`' `'$global:PRD`'"
            <#}
        if ($global:ideName -in "IntelliJ IDEA", "PyCharm") {#>
            $textBoxStatus.Background = "#fce205"
            $textBoxStatus.Height = "80"
            $textBoxStatus.Text = "Click `"Copy & Launch`" button to copy IDE launch command to clipboard and open a cmd prompt.`nPaste the launch command from clipboard in the cmd prompt and hit enter to launch the IDE. `nOnce done, you can close the cmd prompt and this window."
            $btnCopyAndLaunch.Visibility = "Visible"
            # Hide the unrelated buttons
            $buttonLaunch.Visibility = "Hidden"
            $buttonLaunchAndClose.Visibility = "Hidden"
        }
        else {
            $textBoxStatus.Height = "25"
            $textBoxStatus.Text = " Select desired IDE below"
            $textBoxStatus.Background = "#33000000"
            $btnCopyAndLaunch.Visibility = "Hidden"
            $buttonLaunch.Visibility = "Visible"
            $buttonLaunchAndClose.Visibility = "Visible"
        }
    })


# Launch Function
function Launch {
    $launchResult = $cliResult = $cliError = $null
    $global:txtBoxError.Visibility = "Hidden"
    $global:txtBoxError.Text = $null
    Logit "Launch sequence initiated..."
    Logit "Launching $global:ideName with LAUNCH_CMD: $global:LAUNCH_CMD"
    Logit "Launch_CMD command length: $($global:LAUNCH_CMD.Length)"
    # Set-Location $global:PRD
    #  LogIt "Setting currenty directory to: $global:PRD"
    if ($global:LAUNCH_CMD.Length -gt '7') {
        if ($global:ideName -in "CMD", "PowerShell") {
            if ($global:ideName -eq "CMD") { 
                if (!($global:LAUNCH_CMD -like "/c*")) {
                    $global:LAUNCH_CMD = "/c $($global:LAUNCH_CMD.replace("`'",''))" 
                }
            }
            $launchResult = Start-Process $global:ideName -ArgumentList `"$global:LAUNCH_CMD`" -WorkingDirectory $global:PRD -PassThru 
            LogIt "Start-Process $global:ideName -ArgumentList `"$global:LAUNCH_CMD`" -WorkingDirectory $global:PRD -PassThru"
            Start-Sleep -Milliseconds 2000
            $cliResult = [pscustomobject]@{ID = $launchResult.id; Proc = $launchResult.ProcessName; Running = $launchResult.Responding }
            if ($cliResult.Running) {
                LogIt "Child CLI Process spawned successfully. ID: $($launchResult.id) Process: $($launchResult.ProcessName)"
            }
            else {
                LogIt "Toggling command line error state to true. Child CLI Process spawned?: $cliResult"
                [bool]$cliError = $true  
                [bool]$launchError = $true 
                $launchResult = powershell summon -f `'$yml`' `'cmd /C echo`' `'$global:PRD`'
                LogIt "Temporary run command: powershell summon -f `'$yml`' `'cmd /C `"echo`"`' `'$global:PRD`'" 
            }
        }
        elseif ($global:ideName -in "IntelliJ IDEA", "PyCharm", "VSCode") {
            $launchResult = powershell summon -f `'$yml`' `'$LAUNCH_SUB_CMD`' `'$global:PRD`'
            LogIt "powershell summon -f `'$yml`' `'$LAUNCH_SUB_CMD`' `'$global:PRD`'"
            if ($launchResult.Length -gt 0) {   
                [bool]$launchError = $true
                if (<#($global:ideName -in "IntelliJ IDEA", "PyCharm") -and #>(!($launchResult -like "Error*"))) {
                    [bool]$launchError = $false
                }
            }
        }
        if ($launchError) {
            Start-Sleep -Seconds 2   
            LogIt "Summon Launch Error: $launchResult"
            $global:txtBoxError.Text = "$launchResult"       
        }  
        if ($launchError -or $cliError) {
            $global:txtBoxError.Visibility = "Visible"
            $textBoxStatus.Background = "#ab2f42"
            $textBoxStatus.Text = " Error with $global:ideName"
        }
    }
}


# Launch Button
$buttonLaunch.add_click({
        Launch
    })

# Exit button
$buttonExit.add_click({
        $Window.Close()
    })

# Debug Button
$buttonDebug.add_click({
        if ($Window.Width -ge "400") {
            $Window.Width = "331"
            LogIt "Debug Exit"
        }
        else {
            $Window.Width = "650"
            LogIt "[$version] Entering debug mode..."
        }
    })
    
#OK\Reset Button
$buttonReset.add_click({
        ResetForm
    })
#Browse Button
$buttonBrowseDir.add_click({
        $startPath = $env:SystemDrive
        if ($global:txtBoxPRD.Text) {
            if (Test-Path $global:txtBoxPRD.Text -ErrorAction SilentlyContinue) {
                $startPath = $global:txtBoxPRD.Text
            }
        }
        $FolderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog -Property @{
            SelectedPath = $startPath
            Description  = "Select Folder of Project Root Directory"
        }
        $FolderBrowser.ShowDialog()
        $PRdir = $FolderBrowser.SelectedPath
        $global:txtBoxPRD.Text = $PRdir
    })
$window.add_loaded({
        ResetForm
    })
# Launch program\begin GUI
$Window.Add_Closing({
        $txtBoxDebug.text | Out-File -FilePath "C:\temp\Summon_IDE-debug.log" -Encoding unicode -Append 
    })
$Window.ShowDialog() | Out-Null
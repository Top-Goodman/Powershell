$psSR = (Get-Location).path

##################
# Connect to Server #
##################
$SiteCode = "" # Site code 
$initParams = @{}
$ProviderMachineName = "" # SMS Provider machine name
if ($null -eq (Get-Module ConfigurationManager)) {
    Import-Module "$($ENV:SMS_ADMIN_UI_PATH)\..\ConfigurationManager.psd1" @initParams 
}

if ($null -eq (Get-PSDrive -Name $SiteCode -PSProvider CMSite -ErrorAction SilentlyContinue)) {
    New-PSDrive -Name $SiteCode -PSProvider CMSite -Root $ProviderMachineName @initParams
}
Set-Location "$($SiteCode):\" @initParams

#######################
# Update this section #
#######################

$global:Product = 'M365'
$global:Version = 'SAE'

$global:DeploymentGroupA = 'All Physical Workstations - Mandatory Deployments'
#$global:DeploymentGroupA = 'Win 10 Machines - Remote Group A'
#$global:DeploymentGroupB = 'Win 10 Machines - Remote Group B'
#$global:DeploymentGroupC = 'Win 10 Machines - Remote Group C'
#$global:DeploymentGroupD = 'Win 10 Machines - Remote Group D'
$global:DeploymentGroupE = 'All Citrix Workstations - Mandatory Deployments'
$global:DeploymentGroupSAEP = 'All Microsoft 365 - SAEP - Mandatory Deployment'

####################################
# Don't Change anything under this #
####################################

# Automation for day and month to prepopulate script for you.
[int]$Offset = 0 #How many days to add to get to the first Monday based on day of the week of day one, of relative next month.
if ((Get-Date).Day -le 7) {
    [int]$MonthOffset = 0 # Used for Current month. If current day is 7 or less it is possible first Monday of the month has not yet occured, or is today. Meaning patches can go live as early as tomorrow.
}
else {
    [int]$MonthOffset = 1 # Used for Next month. If current day is 8 or more, either it's the month you were assigned patches, or you're running this script too late and can adjust manually.
}

switch (($(get-date -day 1).Addmonths($MonthOffset)).DayofWeek) {
    #Get day of week, of first day of the month.
    "Monday" { $Offset = 0 }
    "Sunday" { $Offset = 1 }
    "Saturday" { $Offset = 2 }
    "Friday" { $Offset = 3 }
    "Thursday" { $Offset = 4 }
    "Wednesday" { $Offset = 5 }
    "Tuesday" { $Offset = 6 }
} 
# Math
$monDay = ($(get-date -day 1).Addmonths($MonthOffset).AddDays($Offset)).Day # Day number of the first Monday of the Month
$monMonth = ($(get-date -day 1).Addmonths($MonthOffset)).Month # Month number of when patches were released by MS

# Automation to get SAE and SAEP names for you
$global:saepLD="Enter M365 - SEAP Update Name Here"
$global:saeLD="Enter M365 - SEA Update Name Here"
$M365Patches = (Get-CMSoftwareUpdate -Fast -DateRevisedMin $(get-date).AddMonths(-3) -DatePostedMin $(get-date).AddMonths(-3)|Where-Object{($_.LocalizedDisplayName -like "*Semi-annual Enterprise*x64*") -and ($_.LocalizedDisplayName -notlike "*extended*") -and ($_.IsSuperseded -eq $false)}).LocalizedDisplayName
if ($M365Patches.count -eq 2) {
foreach ($patch in $M365Patches) {
    if ($patch -like "*(Preview)*") {
        $global:saepLD=$patch
    }
    else {
        $global:saeLD=$patch
    }
}
}

#Create Main Window
Add-Type -assembly System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$main_form = New-Object System.Windows.Forms.Form
$main_form.Text = 'Add Deployments for M365 Update'
$main_form.Width = 560
$main_form.Height = 360
$main_form.AutoSize = $true

$UpdateSAEPNameLabel = New-Object System.Windows.Forms.Label
$UpdateSAEPNameLabel.Text = "M365 SAEP Update Name:"
$UpdateSAEPNamelabel.Location = New-Object System.Drawing.Point(10, 10)
$UpdateSAEPNameLabel.AutoSize = $true
$main_form.Controls.Add($UpdateSAEPNameLabel)

$UpdateSAEPNameTB = New-Object System.Windows.Forms.TextBox
$UpdateSAEPNameTB.Location = New-Object System.Drawing.Point(160, 5)
$UpdateSAEPNameTB.Text = "$global:saepLD"
$UpdateSAEPNameTB.Size = New-Object System.Drawing.Size(370, 20)
$main_form.Controls.Add($UpdateSAEPNameTB)

$UpdateSAENameLabel = New-Object System.Windows.Forms.Label
$UpdateSAENameLabel.Text = "M365 SAE Update Name:"
$UpdateSAENamelabel.Location = New-Object System.Drawing.Point(10, 35)
$UpdateSAENameLabel.AutoSize = $true
$main_form.Controls.Add($UpdateSAENameLabel)

$UpdateSAENameTB = New-Object System.Windows.Forms.TextBox
$UpdateSAENameTB.Location = New-Object System.Drawing.Point(160, 30)
$UpdateSAENameTB.Text = "$global:saeLD"
$UpdateSAENameTB.Size = New-Object System.Drawing.Size(370, 20)
$main_form.Controls.Add($UpdateSAENameTB)

$MonthLabel = New-Object System.Windows.Forms.Label
$MonthLabel.Text = "Month:"
$MonthLabel.Location = New-Object System.Drawing.Point(10, 60)
$MonthLabel.AutoSize = $true
$main_form.Controls.Add($MonthLabel)

$MonthLabelToolTip = New-Object System.Windows.Forms.ToolTip
$MonthLabelToolTip.SetToolTip($MonthLabel, "Enter the month that you are assigned updates, aka July patches are 7 even though they are production in August")

$MonthCB = New-Object System.Windows.Forms.ComboBox
$MonthCB.Location = New-Object System.Drawing.Point(50, 55)
1..9 | ForEach-Object { $MonthCB.Items.Add("0$_") }
10..12 | ForEach-Object { $MonthCB.Items.Add("$_") }
$MonthCB.Size = New-Object System.Drawing.Size(50, 20)
$MonthCB.SelectedIndex = $($monMonth - 1) # -1 is to adjust for index starting at 0=1
$main_form.Controls.Add($MonthCB)

$YearLabel = New-Object System.Windows.Forms.Label
$YearLabel.Text = "Year:"
$YearLabel.Location = New-Object System.Drawing.Point(105, 60)
$YearLabel.AutoSize = $true
$main_form.Controls.Add($YearLabel)

$YearLabelToolTip = New-Object System.Windows.Forms.ToolTip
$YearLabelToolTip.SetToolTip($YearLabel, "Enter the year that you are assigned updates, aka December 2021 patches are still 2021 even though they go production in January 2022")

$YearCB = New-Object System.Windows.Forms.ComboBox
$YearCB.Location = New-Object System.Drawing.Point(140, 55)
$YearCB.Items.Add($(((get-date).year) - 1))
$YearCB.Items.Add($((get-date).year))
$YearCB.Items.Add($(((get-date).year) + 1))
$YearCB.Size = New-Object System.Drawing.Size(50, 20)
$YearCB.SelectedIndex = 1
$main_form.Controls.Add($YearCB)

$StartTimeLabel = New-Object System.Windows.Forms.Label
$StartTimeLabel.Text = "Start Time:"
$StartTimeLabel.Location = New-Object System.Drawing.Point(200, 60)
$StartTimeLabel.AutoSize = $true
$main_form.Controls.Add($StartTimeLabel)

$StartTimeCB = New-Object System.Windows.Forms.ComboBox
$StartTimeCB.Location = New-Object System.Drawing.Point(260, 55)
1..12 | ForEach-Object { $StartTimeCB.Items.Add("$_`:00") }
$StartTimeCB.Size = New-Object System.Drawing.Size(50, 20)
$StartTimeCB.SelectedIndex = 3
$main_form.Controls.Add($StartTimeCB)

$StartAMPMTimeCB = New-Object System.Windows.Forms.ComboBox
$StartAMPMTimeCB.Location = New-Object System.Drawing.Point(310, 55)
$StartAMPMTimeCB.Items.Add('AM')
$StartAMPMTimeCB.Items.Add('PM')
$StartAMPMTimeCB.Size = New-Object System.Drawing.Size(40, 20)
$StartAMPMTimeCB.SelectedIndex = 1
$main_form.Controls.Add($StartAMPMTimeCB)

$EndTimeLabel = New-Object System.Windows.Forms.Label
$EndTimeLabel.Text = "End Time:"
$EndTimeLabel.Location = New-Object System.Drawing.Point(380, 60)
$EndTimeLabel.AutoSize = $true
$main_form.Controls.Add($EndTimeLabel)

$EndTimeCB = New-Object System.Windows.Forms.ComboBox
$EndTimeCB.Location = New-Object System.Drawing.Point(440, 55)
1..12 | ForEach-Object { $EndTimeCB.Items.Add("$_`:00") }
$EndTimeCB.Size = New-Object System.Drawing.Size(50, 20)
$EndTimeCB.SelectedIndex = 4
$main_form.Controls.Add($EndTimeCB)

$EndAMPMTimeCB = New-Object System.Windows.Forms.ComboBox
$EndAMPMTimeCB.Location = New-Object System.Drawing.Point(490, 55)
$EndAMPMTimeCB.Items.Add('AM')
$EndAMPMTimeCB.Items.Add('PM')
$EndAMPMTimeCB.Size = New-Object System.Drawing.Size(40, 20)
$EndAMPMTimeCB.SelectedIndex = 1
$main_form.Controls.Add($EndAMPMTimeCB)

$FirstDayLabel = New-Object System.Windows.Forms.Label
$FirstDayLabel.Text = "First Monday of Full Week:"
$FirstDayLabel.Location = New-Object System.Drawing.Point(10, 85)
$FirstDayLabel.AutoSize = $true
$main_form.Controls.Add($FirstDayLabel)

$FirstDayToolTip = New-Object System.Windows.Forms.ToolTip
$FirstDayToolTip.SetToolTip($FirstDayLabel, "Enter the day number of the first Monday of the month that updates will deploy to Production.")#, If that Monday is a Holiday, check the 'Holiday?' box below to make SAEP, GroupA, and GroupB all start that Tuesday"

$FirstDayCB = New-Object System.Windows.Forms.ComboBox
$FirstDayCB.Location = New-Object System.Drawing.Point(150, 80)
1..9 | ForEach-Object { $FirstDayCB.Items.Add("0$_") }
10..12 | ForEach-Object { $FirstDayCB.Items.Add("$_") }
$FirstDayCB.Size = New-Object System.Drawing.Size(40, 20)
$FirstDayCB.SelectedIndex = $($monDay - 1) # -1 is to adjust for index starting at 0=1
$main_form.Controls.Add($FirstDayCB)

$CitrixStartTimeLabel = New-Object System.Windows.Forms.Label
$CitrixStartTimeLabel.Text = "Citrix Start:"
$CitrixStartTimeLabel.Location = New-Object System.Drawing.Point(198, 85)
$CitrixStartTimeLabel.AutoSize = $true
$main_form.Controls.Add($CitrixStartTimeLabel)

$CitrixStartTimeCB = New-Object System.Windows.Forms.ComboBox
$CitrixStartTimeCB.Location = New-Object System.Drawing.Point(260, 80)
1..12 | ForEach-Object { $CitrixStartTimeCB.Items.Add("$_`:00") }
$CitrixStartTimeCB.Size = New-Object System.Drawing.Size(50, 20)
$CitrixStartTimeCB.SelectedIndex = 3
$main_form.Controls.Add($CitrixStartTimeCB)

$CitrixStartAMPMTimeCB = New-Object System.Windows.Forms.ComboBox
$CitrixStartAMPMTimeCB.Location = New-Object System.Drawing.Point(310, 80)
$CitrixStartAMPMTimeCB.Items.Add('AM')
$CitrixStartAMPMTimeCB.Items.Add('PM')
$CitrixStartAMPMTimeCB.Size = New-Object System.Drawing.Size(40, 20)
$CitrixStartAMPMTimeCB.SelectedIndex = 1
$main_form.Controls.Add($CitrixStartAMPMTimeCB)

$CitrixEndTimeLabel = New-Object System.Windows.Forms.Label
$CitrixEndTimeLabel.Text = "Citrix End:"
$CitrixEndTimeLabel.Location = New-Object System.Drawing.Point(378, 85)
$CitrixEndTimeLabel.AutoSize = $true
$main_form.Controls.Add($CitrixEndTimeLabel)

$CitrixEndTimeCB = New-Object System.Windows.Forms.ComboBox
$CitrixEndTimeCB.Location = New-Object System.Drawing.Point(440, 80)
1..12 | ForEach-Object { $CitrixEndTimeCB.Items.Add("$_`:00") }
$CitrixEndTimeCB.Size = New-Object System.Drawing.Size(50, 20)
$CitrixEndTimeCB.SelectedIndex = 4
$main_form.Controls.Add($CitrixEndTimeCB)

$CitrixEndAMPMTimeCB = New-Object System.Windows.Forms.ComboBox
$CitrixEndAMPMTimeCB.Location = New-Object System.Drawing.Point(490, 80)
$CitrixEndAMPMTimeCB.Items.Add('AM')
$CitrixEndAMPMTimeCB.Items.Add('PM')
$CitrixEndAMPMTimeCB.Size = New-Object System.Drawing.Size(40, 20)
$CitrixEndAMPMTimeCB.SelectedIndex = 1
$main_form.Controls.Add($CitrixEndAMPMTimeCB)

<#
$MondayHolidayChB = New-Object System.Windows.Forms.Checkbox
$MondayHolidayChB.Location = New-Object System.Drawing.Size(10,105)
$MondayHolidayChB.Size = New-Object System.Drawing.Size(70,23)
$MondayHolidayChB.Text = "Holiday?"
$main_form.Controls.Add($MondayHolidayChB)
#>

$CreateDeploymentButton = New-Object System.Windows.Forms.Button
$CreateDeploymentButton.Location = New-Object System.Drawing.Size(220, 105)
$CreateDeploymentButton.Size = New-Object System.Drawing.Size(140, 23)
$CreateDeploymentButton.Text = "Create Deployment"
$CreateDeploymentButton.Visible = $false
$main_form.Controls.Add($CreateDeploymentButton)

$CalcuteDeploymentsButton = New-Object System.Windows.Forms.Button
$CalcuteDeploymentsButton.Location = New-Object System.Drawing.Size(80, 105)
$CalcuteDeploymentsButton.Size = New-Object System.Drawing.Size(140, 23)
$CalcuteDeploymentsButton.Text = "Calculate Deployment"
$main_form.Controls.Add($CalcuteDeploymentsButton)

$DeploymentNameALabel = New-Object System.Windows.Forms.Label
$DeploymentNameALabel.Location = New-Object System.Drawing.Point(10, 130)
$DeploymentNameALabel.AutoSize = $true
$main_form.Controls.Add($DeploymentNameALabel)

$DeploymentNameATimeLabel = New-Object System.Windows.Forms.Label
$DeploymentNameATimeLabel.Location = New-Object System.Drawing.Point(20, 145)
$DeploymentNameATimeLabel.AutoSize = $true
$main_form.Controls.Add($DeploymentNameATimeLabel)
<#
$DeploymentNameBLabel = New-Object System.Windows.Forms.Label
$DeploymentNameBLabel.Location  = New-Object System.Drawing.Point(10,170)
$DeploymentNameBLabel.AutoSize = $true
$main_form.Controls.Add($DeploymentNameBLabel)

$DeploymentNameBTimeLabel = New-Object System.Windows.Forms.Label
$DeploymentNameBTimeLabel.Location  = New-Object System.Drawing.Point(20,185)
$DeploymentNameBTimeLabel.AutoSize = $true
$main_form.Controls.Add($DeploymentNameBTimeLabel)

$DeploymentNameCLabel = New-Object System.Windows.Forms.Label
$DeploymentNameCLabel.Location  = New-Object System.Drawing.Point(10,210)
$DeploymentNameCLabel.AutoSize = $true
$main_form.Controls.Add($DeploymentNameCLabel)

$DeploymentNameCTimeLabel = New-Object System.Windows.Forms.Label
$DeploymentNameCTimeLabel.Location  = New-Object System.Drawing.Point(20,225)
$DeploymentNameCTimeLabel.AutoSize = $true
$main_form.Controls.Add($DeploymentNameCTimeLabel)

$DeploymentNameDLabel = New-Object System.Windows.Forms.Label
$DeploymentNameDLabel.Location  = New-Object System.Drawing.Point(10,250)
$DeploymentNameDLabel.AutoSize = $true
$main_form.Controls.Add($DeploymentNameDLabel)

$DeploymentNameDTimeLabel = New-Object System.Windows.Forms.Label
$DeploymentNameDTimeLabel.Location  = New-Object System.Drawing.Point(20,265)
$DeploymentNameDTimeLabel.AutoSize = $true
$main_form.Controls.Add($DeploymentNameDTimeLabel)
#>
$DeploymentNameELabel = New-Object System.Windows.Forms.Label
$DeploymentNameELabel.Location = New-Object System.Drawing.Point(10, 170)
$DeploymentNameELabel.AutoSize = $true
$main_form.Controls.Add($DeploymentNameELabel)

$DeploymentNameETimeLabel = New-Object System.Windows.Forms.Label
$DeploymentNameETimeLabel.Location = New-Object System.Drawing.Point(20, 185)
$DeploymentNameETimeLabel.AutoSize = $true
$main_form.Controls.Add($DeploymentNameETimeLabel)

$DeploymentNameSAEPLabel = New-Object System.Windows.Forms.Label
$DeploymentNameSAEPLabel.Location = New-Object System.Drawing.Point(10, 210)
$DeploymentNameSAEPLabel.AutoSize = $true
$main_form.Controls.Add($DeploymentNameSAEPLabel)

$DeploymentNameSAEPTimeLabel = New-Object System.Windows.Forms.Label
$DeploymentNameSAEPTimeLabel.Location = New-Object System.Drawing.Point(20, 225)
$DeploymentNameSAEPTimeLabel.AutoSize = $true
$main_form.Controls.Add($DeploymentNameSAEPTimeLabel)

$CreateDeploymentALabel = New-Object System.Windows.Forms.Label
$CreateDeploymentALabel.Location = New-Object System.Drawing.Point(10, 250)
$CreateDeploymentALabel.AutoSize = $true
$main_form.Controls.Add($CreateDeploymentALabel)
<#
$CreateDeploymentBLabel = New-Object System.Windows.Forms.Label
$CreateDeploymentBLabel.Location  = New-Object System.Drawing.Point(10,385)
$CreateDeploymentBLabel.AutoSize = $true
$main_form.Controls.Add($CreateDeploymentBLabel)

$CreateDeploymentCLabel = New-Object System.Windows.Forms.Label
$CreateDeploymentCLabel.Location  = New-Object System.Drawing.Point(10,400)
$CreateDeploymentCLabel.AutoSize = $true
$main_form.Controls.Add($CreateDeploymentCLabel)

$CreateDeploymentDLabel = New-Object System.Windows.Forms.Label
$CreateDeploymentDLabel.Location  = New-Object System.Drawing.Point(10,415)
$CreateDeploymentDLabel.AutoSize = $true
$main_form.Controls.Add($CreateDeploymentDLabel)
#>
$CreateDeploymentELabel = New-Object System.Windows.Forms.Label
$CreateDeploymentELabel.Location = New-Object System.Drawing.Point(10, 265)
$CreateDeploymentELabel.AutoSize = $true
$main_form.Controls.Add($CreateDeploymentELabel)

$CreateDeploymentSAEPLabel = New-Object System.Windows.Forms.Label
$CreateDeploymentSAEPLabel.Location = New-Object System.Drawing.Point(10, 280)
$CreateDeploymentSAEPLabel.AutoSize = $true
$main_form.Controls.Add($CreateDeploymentSAEPLabel)


$CreateDeploymentButton.Add_Click(

    {
        $global:defaultForeColor = "Control"
        $global:SuccessForeColor = "Blue"
        $global:SuccessBackColor = "Cyan"
        $global:FailureForeColor = "Red"
        $global:FailureBackColor = "Black"
        $CreateDeploymentALabel.Text = ""
        $CreateDeploymentALabel.ForeColor = $defaultForeColor
        #  $CreateDeploymentBLabel.Text=""
        $CreateDeploymentALabel.ForeColor = $defaultForeColor
        #  $CreateDeploymentCLabel.Text=""
        $CreateDeploymentALabel.ForeColor = $defaultForeColor
        # $CreateDeploymentDLabel.Text=""
        $CreateDeploymentALabel.ForeColor = $defaultForeColor
        $CreateDeploymentELabel.Text = ""
        $CreateDeploymentALabel.ForeColor = $defaultForeColor
        $CreateDeploymentSAEPLabel.Text = ""
        $CreateDeploymentALabel.ForeColor = $defaultForeColor

        $global:SAEPname = $UpdateSAEPNameTB.Text
        $global:SAEname = $UpdateSAENameTB.Text

        $global:update = Get-CMSoftwareUpdate -Name $SAEname -Fast
        $global:SAEPupdate = Get-CMSoftwareUpdate -Name $SAEPname -Fast

        #Create Group A Deployment
        try {
            New-CMSoftwareUpdateDeployment -DeploymentName "$global:DeploymentNameA" `
                -SoftwareUpdateName $global:update.LocalizedDisplayName `
                -CollectionName $global:DeploymentGroupA `
                -DeploymentType Required `
                -VerbosityLevel AllMessages `
                -AvailableDateTime "$global:year/$global:DeploymentMonth/$global:GroupAStartDay $global:starttime$global:StartTimeAMPM" `
                -DeadlineDateTime "$global:year/$global:DeploymentMonth/$global:GroupAEndDay $global:endtime$global:EndTimeAMPM" `
                -UserNotification DisplaySoftwareCenterOnly `
                -SoftwareInstallation $True  `
                -AllowRestart $True  `
                -RestartServer $False `
                -RestartWorkstation $True `
                -PersistOnWriteFilterDevice $false `
                -RequirePostRebootFullScan $True `
                -ProtectedType RemoteDistributionPoint `
                -AcceptEula | Out-Null
            $CreateDeploymentALabel.Text = "$global:DeploymentNameA Created!"
            $CreateDeploymentALabel.ForeColor = $global:SuccessForeColor
            $CreateDeploymentALabel.BackColor = $global:SuccessBackColor
        }
        Catch {
            $CreateDeploymentALabel.Text = "$global:DeploymentNameA failed to Create!"
            $CreateDeploymentALabel.ForeColor = $global:FailureForeColor
            $CreateDeploymentALabel.BackColor = $global:FailureBackColor
        }
        <#
    #Create Group B Deployment
    try
    {
    New-CMSoftwareUpdateDeployment -DeploymentName "$global:DeploymentNameB" `
        -SoftwareUpdateName $global:update.LocalizedDisplayName `
        -CollectionName $global:DeploymentGroupB `
        -DeploymentType Required `
        -VerbosityLevel AllMessages `
        -AvailableDateTime "$global:year/$global:DeploymentMonth/$global:GroupBStartDay $global:starttime$global:StartTimeAMPM" `
        -DeadlineDateTime "$global:year/$global:DeploymentMonth/$global:GroupBEndDay $global:endtime$global:EndTimeAMPM" `
        -UserNotification DisplaySoftwareCenterOnly `
        -SoftwareInstallation $True  `
        -AllowRestart $True  `
        -RestartServer $False `
        -RestartWorkstation $True `
        -PersistOnWriteFilterDevice $false `
        -RequirePostRebootFullScan $True `
        -ProtectedType RemoteDistributionPoint `
        -AcceptEULA | Out-Null
        $CreateDeploymentBLabel.Text = "$global:DeploymentNameB Created!"
        $CreateDeploymentBLabel.ForeColor = $global:SuccessForeColor
        $CreateDeploymentBLabel.BackColor = $global:SuccessBackColor
    }
    Catch
    {
        $CreateDeploymentBLabel.Text = "$global:DeploymentNameB failed to Create!"
        $CreateDeploymentBLabel.ForeColor = $global:FailureForeColor
        $CreateDeploymentBLabel.BackColor = $global:FailureBackColor
    }

    #Create Group C Deployment

    try
    {
    New-CMSoftwareUpdateDeployment -DeploymentName "$global:DeploymentNameC" `
        -SoftwareUpdateName $global:update.LocalizedDisplayName `
        -CollectionName $global:DeploymentGroupC `
        -DeploymentType Required `
        -VerbosityLevel AllMessages `
        -AvailableDateTime "$global:year/$global:DeploymentMonth/$global:GroupCStartDay $global:starttime$global:StartTimeAMPM" `
        -DeadlineDateTime "$global:year/$global:DeploymentMonth/$global:GroupCEndDay $global:endtime$global:EndTimeAMPM" `
        -UserNotification DisplaySoftwareCenterOnly `
        -SoftwareInstallation $True  `
        -AllowRestart $True  `
        -RestartServer $False `
        -RestartWorkstation $True `
        -PersistOnWriteFilterDevice $false `
        -RequirePostRebootFullScan $True `
        -ProtectedType RemoteDistributionPoint `
        -AcceptEULA | Out-Null
        $CreateDeploymentCLabel.Text = "$global:DeploymentNameC Created!"
        $CreateDeploymentCLabel.ForeColor = $global:SuccessForeColor
        $CreateDeploymentCLabel.BackColor = $global:SuccessBackColor
    }
    Catch
    {
        $CreateDeploymentCLabel.Text = "$global:DeploymentNameC failed to Create!"
        $CreateDeploymentCLabel.ForeColor = $global:FailureForeColor
        $CreateDeploymentCLabel.BackColor = $global:FailureBackColor
    }

    #Create Group D Deployment

    try
    {
    New-CMSoftwareUpdateDeployment -DeploymentName "$global:DeploymentNameD" `
        -SoftwareUpdateName $global:update.LocalizedDisplayName `
        -CollectionName $global:DeploymentGroupD `
        -DeploymentType Required `
        -VerbosityLevel AllMessages `
        -AvailableDateTime "$global:year/$global:DeploymentMonth/$global:GroupDStartDay $global:starttime$global:StartTimeAMPM" `
        -DeadlineDateTime "$global:year/$global:DeploymentMonth/$global:GroupDEndDay $global:endtime$global:EndTimeAMPM" `
        -UserNotification DisplaySoftwareCenterOnly `
        -SoftwareInstallation $True  `
        -AllowRestart $True  `
        -RestartServer $False `
        -RestartWorkstation $True `
        -PersistOnWriteFilterDevice $False `
        -RequirePostRebootFullScan $True `
        -ProtectedType RemoteDistributionPoint `
        -AcceptEULA | Out-Null
        $CreateDeploymentDLabel.Text = "$global:DeploymentNameD Created!"
        $CreateDeploymentDLabel.ForeColor = $global:SuccessForeColor
        $CreateDeploymentDLabel.BackColor = $global:SuccessBackColor
    }
    Catch
    {
        $CreateDeploymentDLabel.Text = "$global:DeploymentNameD failed to Create!"
        $CreateDeploymentDLabel.ForeColor = $global:FailureForeColor
        $CreateDeploymentDLabel.BackColor = $global:FailureBackColor
    }
#>
        #Create Citrix Deployment

        try {
            New-CMSoftwareUpdateDeployment -DeploymentName "$global:DeploymentNameE" `
                -SoftwareUpdateName $global:update.LocalizedDisplayName `
                -CollectionName $global:DeploymentGroupE `
                -DeploymentType Required `
                -VerbosityLevel AllMessages `
                -AvailableDateTime "$global:year/$global:DeploymentMonth/$global:GroupEStartDay $global:Citrixstarttime$global:CitrixStartTimeAMPM" `
                -DeadlineDateTime "$global:year/$global:DeploymentMonth/$global:GroupEEndDay $global:Citrixendtime$global:CitrixEndTimeAMPM" `
                -UserNotification DisplaySoftwareCenterOnly `
                -SoftwareInstallation $True  `
                -AllowRestart $True  `
                -RestartServer $False `
                -RestartWorkstation $True `
                -PersistOnWriteFilterDevice $False `
                -RequirePostRebootFullScan $True `
                -ProtectedType RemoteDistributionPoint `
                -AcceptEULA | Out-Null
            $CreateDeploymentELabel.Text = "$global:DeploymentNameE Created!"
            $CreateDeploymentELabel.ForeColor = $global:SuccessForeColor
            $CreateDeploymentELabel.BackColor = $global:SuccessBackColor
        }
        Catch {
            $CreateDeploymentELabel.Text = "$global:DeploymentNameE failed to Create!"
            $CreateDeploymentELabel.ForeColor = $global:FailureForeColor
            $CreateDeploymentELabel.BackColor = $global:FailureBackColor
        }

        #Create SAEP Deployment

        try {
            New-CMSoftwareUpdateDeployment -DeploymentName "$global:DeploymentNameSAEP" `
                -SoftwareUpdateName $global:SAEPupdate.LocalizedDisplayName `
                -CollectionName $global:DeploymentGroupSAEP `
                -DeploymentType Required `
                -VerbosityLevel AllMessages `
                -AvailableDateTime "$global:year/$global:DeploymentMonth/$global:GroupAStartDay $global:starttime$global:StartTimeAMPM" `
                -DeadlineDateTime "$global:year/$global:DeploymentMonth/$global:GroupAEndDay $global:endtime$global:EndTimeAMPM" `
                -UserNotification DisplaySoftwareCenterOnly `
                -SoftwareInstallation $True  `
                -AllowRestart $True  `
                -RestartServer $False `
                -RestartWorkstation $True `
                -PersistOnWriteFilterDevice $false `
                -RequirePostRebootFullScan $True `
                -ProtectedType RemoteDistributionPoint `
                -AcceptEula | Out-Null
            $CreateDeploymentSAEPLabel.Text = "$global:DeploymentNameSAEP Created!"
            $CreateDeploymentSAEPLabel.ForeColor = $SuccessForeColor
            $CreateDeploymentSAEPLabel.BackColor = $SuccessBackColor
        }
        Catch {
            $CreateDeploymentSAEPLabel.Text = "$global:DeploymentNameSAEP failed to Create!"
            $CreateDeploymentSAEPLabel.ForeColor = $FailureForeColor
            $CreateDeploymentSAEPLabel.BackColor = $FailureBackColor
        }

    }

)

$CalcuteDeploymentsButton.Add_Click(

    {
        [int]$global:month = $MonthCB.SelectedItem
        [int]$global:year = $YearCB.SelectedItem
        [int]$global:GroupAStartDay = [int]$FirstDayCB.SelectedItem + 1
        $global:starttime = $StartTimeCB.SelectedItem
        $global:endtime = $EndTimeCB.SelectedItem
        $global:StartTimeAMPM = $StartAMPMTimeCB.SelectedItem
        $global:EndTimeAMPM = $EndAMPMTimeCB.SelectedItem
        $global:Citrixstarttime = $CitrixStartTimeCB.SelectedItem
        $global:CitrixStartTimeAMPM = $CitrixStartAMPMTimeCB.SelectedItem
        $global:Citrixendtime = $CitrixEndTimeCB.SelectedItem
        $global:CitrixEndTimeAMPM = $CitrixEndAMPMTimeCB.SelectedItem

        if ($global:month -lt 12) {
            $global:DeploymentMonth = $month + 1
        }
        else {
            $global:DeploymentMonth = 1
            $global:year = $global:year + 1
        }
        $global:GroupAEndDay = $GroupAStartDay + 2
        <# $global:GroupBStartDay = $GroupAStartDay + 1
    $global:GroupBEndDay = $GroupAEndDay + 1
    $global:GroupCStartDay = $GroupBStartDay + 1
    $global:GroupCEndDay = $GroupBEndDay + 1
    $global:GroupDStartDay = $GroupCStartDay + 1
    $global:GroupDEndDay = $GroupCEndDay + 3#>
        $global:GroupEStartDay = $GroupAStartDay #+ 1
        $global:GroupEEndDay = $GroupAEndDay #+ 1
    
        <#
    if ($MondayHolidayChB.Checked){
        $global:GroupAStartDay++
        $global:GroupAEndDay++
    }
#>

        if ($month -lt 10) {
            $global:DeploymentNameA = "Desktop - Prod - $global:Product - $global:Version - $global:year`_0$global:month - All Physical Workstations"
            <#  $global:DeploymentNameA = "Desktop - Prod - $global:Product - $global:Version - $global:year`_0$global:month - Remote Group A"
       $global:DeploymentNameB = "Desktop - Prod - $global:Product - $global:Version - $global:year`_0$global:month - Remote Group B"
        $global:DeploymentNameC = "Desktop - Prod - $global:Product - $global:Version - $global:year`_0$global:month - Remote Group C"
       $global:DeploymentNameD = "Desktop - Prod - $global:Product - $global:Version - $global:year`_0$global:month - Remote Group D"#>
            $global:DeploymentNameE = "Citrix  - Prod - $global:Product - $global:Version - $global:year`_0$global:month - All Citrix Workstations"
            $global:DeploymentNameSAEP = "Desktop - Prod - $global:Product - SAEP - $global:year`_0$global:month - All M365 SAEP Machines"
        }
        else {
            $global:DeploymentNameA = "Desktop - Prod - $global:Product - $global:Version - $global:year`_$global:month - All Physical Workstations"
            <#  $global:DeploymentNameA = "Desktop - Prod - $global:Product - $global:Version - $global:year`_$global:month - Remote Group A"
        $global:DeploymentNameB = "Desktop - Prod - $global:Product - $global:Version - $global:year`_$global:month - Remote Group B"
        $global:DeploymentNameC = "Desktop - Prod - $global:Product - $global:Version - $global:year`_$global:month - Remote Group C"
        $global:DeploymentNameD = "Desktop - Prod - $global:Product - $global:Version - $global:year`_$global:month - Remote Group D"#>
            $global:DeploymentNameE = "Citrix  - Prod - $global:Product - $global:Version - $global:year`_$global:month - All Citrix Workstations"
            $global:DeploymentNameSAEP = "Desktop - Prod - $global:Product - SAEP - $global:year`_$global:month - All M365 SAEP Machines"
        }

        $DeploymentNameALabel.Text = "Deployment Name - Group A: $DeploymentNameA"
        $DeploymentNameATimeLabel.Text = "      $DeploymentMonth`/$GroupAStartDay`/$global:year $global:starttime$global:StartTimeAMPM - $DeploymentMonth`/$GroupAEndDay`/$global:year $global:endtime$EndTimeAMPM"
        <# $DeploymentNameBLabel.Text = "Deployment Name - Group B: $DeploymentNameB"
    $DeploymentNameBTimeLabel.Text = "      $DeploymentMonth`/$GroupBStartDay`/$global:year $global:starttime$global:StartTimeAMPM - $DeploymentMonth`/$GroupBEndDay`/$global:year $global:endtime$EndTimeAMPM"
    $DeploymentNameCLabel.Text = "Deployment Name - Group C: $DeploymentNameC"
    $DeploymentNameCTimeLabel.Text = "      $DeploymentMonth`/$GroupCStartDay`/$global:year $global:starttime$global:StartTimeAMPM - $DeploymentMonth`/$GroupCEndDay`/$global:year $global:endtime$EndTimeAMPM"
    $DeploymentNameDLabel.Text = "Deployment Name - Group D: $DeploymentNameD"
    $DeploymentNameDTimeLabel.Text = "      $DeploymentMonth`/$GroupDStartDay`/$global:year $global:starttime$global:StartTimeAMPM - $DeploymentMonth`/$GroupDEndDay`/$global:year $global:endtime$EndTimeAMPM"#>
        $DeploymentNameELabel.Text = "Deployment Name - Citrix : $DeploymentNameE"
        $DeploymentNameETimeLabel.Text = "      $DeploymentMonth`/$GroupEStartDay`/$global:year $global:Citrixstarttime$global:CitrixStartTimeAMPM - $DeploymentMonth`/$GroupEEndDay`/$global:year $global:Citrixendtime$global:CitrixEndTimeAMPM"
        $DeploymentNameSAEPLabel.Text = "Deployment Name - SAEP   : $DeploymentNameSAEP"
        $DeploymentNameSAEPTimeLabel.Text = "      $DeploymentMonth`/$GroupAStartDay`/$global:year $global:starttime$global:StartTimeAMPM - $DeploymentMonth`/$GroupAEndDay`/$global:year $global:endtime$global:EndTimeAMPM"

        $CreateDeploymentButton.Visible = $true
        $CalcuteDeploymentsButton.Text = "Re-Calulate Deployments"
    }

)

$main_form.ShowDialog()
Set-Location $psSR

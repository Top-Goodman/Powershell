# Includes - Must modify this module based on host where ccm is installed.
#Import-Module $env:SMS_ADMIN_UI_PATH.Replace('i386', 'ConfigurationManager.psd1') -Force -ErrorAction Stop
#Import-Module "\\server\share$\Program Files\Microsoft Configuration Manager\AdminConsole\bin\ConfigurationManager\ConfigurationManager.psd1"  -Force -ErrorAction Stop
Import-Module "\\server\ConfigurationManager_bin\ConfigurationManager.psd1" -Force -ErrorAction Stop # Might need to change this based on location and permission.
# Custom definition for file metadata extraction -> image
$icoCode = @"
using System;
using System.Drawing;
using System.Runtime.InteropServices;

namespace System
{
   [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
        public struct SHFILEINFO
        {
            public IntPtr hIcon;
            public int iIcon;
            public uint dwAttributes;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 260)]
            public string szDisplayName;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 80)]
            public string szTypeName;
};
    public class SHGETFILEINFO
    {
     [DllImport("shell32.dll", CharSet = CharSet.Unicode)]
     public static extern IntPtr SHGetFileInfo(string pszPath, uint     dwFileAttributes,ref SHFILEINFO psfi, uint cbSizeFileInfo, uint uFlags);
    }
}
"@
Add-Type -TypeDefinition $icoCode
#Load forms for sendkeys (For interactive email stuff, reporting)
#[System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")  
##*===============================================
##* VARIABLE DECLARATION
##*===============================================
[string]$cmSiteCode = "" #Site Code
[string]$fileRepo = "\\Server\Share\"
[string]$aColName = "Test Collection" 
[string]$gName = "Dev DPs" # Can maybe do multiple like this '$("Desktop - DP Group - Dev","Desktop - DP Group - Prod")'
[string]$appVendor = "Microsoft" 
[string]$appName = "Power BI Desktop"
[string]$appVersion = ((($((Invoke-WebRequest -Uri "https://aka.ms/pbiSingleInstaller" -UseBasicParsing).Content) -split ('Version:'))[1] -split ('p'))[1]) -replace '[</>]', ''
[string]$app = "PBIDesktopSetup-$appVersion-x64.exe"
[version]$vNewVer = $appVersion
[string]$outFolder = "$FileRepo\$appname\$appVersion"
[string]$outFile = "$outFolder\$app"
[string]$outSource = "$FileRepo\$appname\Install Script\Files\$app"
[string]$outIco = "$outFolder\$appName.ico"
[String]$cmName = "$appVendor - $appName - $appVersion"
[string]$cmDescirption = "FakeApplication - created by Powershell, brought to you in part by Spencer."
[bool]$cmAutoInstall = $false
[bool]$cmFeatured = $false
[string]$cmOwner = "K28228"
[string]$cmSupportContact = "R28228"
[string]$cmOptRef = ""
[string]$lDescription = "***Requires Microsoft Power BI account creation after install. Sign in is not required for use.***

Microsoft Power BI Desktop is built for the analyst. It combines state-of-the-art interactive visualizations, with industry-leading data query and modeling built-in. Create and publish your reports to Power BI. Power BI Desktop helps you empower others with timely critical insights, anytime, anywhere. Create rich, interactive reports with visual analytics at your fingertips."
[string]$lAppName = "$appName"
[string]$lPublisher = "$appVendor"
[string]$lKeywords = "PBI, Microsoft, Graphs, Charts,stuff, things"
[string]$lUserDoc = "https://docs.microsoft.com/en-us/power-bi/"
[string]$lLinkText = "More Information"
[string]$lPrivacyURL = ""
[string]$lIconPath = "$outIco"
[string]$lCatalog = "" # Administrative Categories maybe? We don't use this setting If I reclal correctly
[string]$pDistPntSet = 'AutoDownload' # Must choose one of AutoDownload | DeltaCopy |NoDownload
[string]$pUserCat = 'Analytical'#User Categories, these we use, sort for Software Center. Can be a comma seperated list
[string]$dContLoc = "$FileRepo\$appname\Install Script"
[string]$dName = "PSAD $AppVendor $appName"
[string]$dComment = ""
[string]$dInstCommand = 'Deploy-Application.exe -DeploymentType Install'
[string]$dInstWorkLoc = ""
[string]$dUninstCommand = 'Deploy-Application.exe -DeploymentType Uninstall'
[string]$dUninstWorkLoc = ""
[int32]$dEstRunTime = 5
[int32]$dMaxRunTime = 120
[string]$dInstBehavior = "InstallForSystem" # Must choose one of InstallForUser | InstallForSystem | InstallForSystemIfResourceIsDeviceOtherwiseInstallForUser
[string]$dLogonReq = "OnlyWhenUserLoggedOn" # Must choose one of OnlyWhenUserLoggedOn | WhereOrNotUserLoggedOn | WhetherOrNotUserLoggedOn | OnlyWhenNoUserLoggedOn
[string]$dRebootBhavior = "NoAction" # Must choose one of BasedOnExitCode | NoAction | ForceReboot | ProgramReboot
[bool]$dReqUserInt = $false
[string]$dUserIntMode = "Normal" # Must choose one of Normal | Minimized | Maximized | Hidden
[string]$dUninstOpt = "SameAsInstall" # Must choose one of SameAsInstall | NoneRequired | Different
[string]$dUninstContLoc = "" # Only use this IF $dUninstOpt is st to "Different"
[bool]$dUseDPBoundaryGroup = $true
[string]$dDownRunLocal = "Download" # Choose one of DoNothing | Download
[string]$dDetecScriptType = "PowerShell" # Must choose one of PowerShell | JavaScript | VBScript
[string]$mPath = 'C:\Program Files\Microsoft Power BI Desktop\bin'
[string]$mFile = 'PBIDesktop.exe'
[string]$mLocaiton = "`"$mPath\$mFile`""
[string]$dDetecScript = "If (test-path $mLocaiton) { If (([version]`$((get-item -ErrorAction SilentlyContinue $mLocaiton).VersionInfo.FileVersion)).CompareTo([version]`"$vNewVer`") -ge 0) {`$true} else {}}"
# These variables are for experimetning with detection caluses (enhanced detection methods)..... 
<#[string]$mPropType = "Version" # Choose one of DateCreated | DateModified | Version | Size
[version]$mVersion = $appVersion
[string]$mExpressOp = "GreaterEquals" # Choose one of IsEquals | NotEquals | GreaterThan | GreaterEquals | LessThan | LessEquals | Between | OneOf | NoneOf#>
[string]$aDepAction = "Install" # Choose one of Install | Uninstall
[string]$aDepPurp = "Available" # Choose one of Available | Required
[string]$aUserNotify = "DisplaySoftwareCenterOnly" # Choose one of DisplayAll | DisplaySoftwareCenterOnly | HideAll
[string]$aTType = "LocalTime" # Choose one of LocalTime | Utc
[DateTime]
[bool]$aApprReq = $false
[bool]$aPreDep = $false
# PowerBi is an annoying exception that I need these ids to create the Uri.... I mean technically I don't, I don't want a giant command for the uri with NESTED Invoke-Webrequests
[string]$id1 = (($((invoke-webrequest https://aka.ms/pbiSingleInstaller -UseBasicParsing).content) -split ('data-bi-dlnm="Microsoft Power BI Desktop" data-bi-dlid="'))[1] -split ('"'))[0]
[string]$id2 = (($((invoke-webrequest https://aka.ms/pbiSingleInstaller -UseBasicParsing).content) -split ("id=$id1&amp;"))[1] -split ('"'))[0]
[string]$Uri = (($((Invoke-WebRequest -Uri "https://www.microsoft.com/en-us/download/confirmation.aspx?id=$($id1)&$($id2)" -UseBasicParsing).content) -split ('downloadData={base_0:{url:"'))[1] -split ('"'))[0]
[string]$cmAppFolder = "$cmSiteCode\Application\Development\Spencer"
[string]$logRoot = "$env:SystemRoot\Logs\cmAutoUpdates1"
[string]$logFileName = "$app-cmAutoUpdate-$(get-date -Format MM.dd.yyyy).log"
[string]$global:cmAppArgs = $null
[string]$global:cmSetArgs = $null
[string]$global:cmDepArgs = $null
[string]$global:cmApdArgs = $null
# End Varibale decleration **** DO Not Modify Below this point *****

##*===============================================
##* FUNCTIONS/PREREQUISITES
##*===============================================
function log ($string) {
    if (!(Test-Path $logRoot)) { New-Item -ItemType Directory $logRoot }
    "$(Get-Date) - $($string)" | Out-File "$logRoot\$logfileName" -Append
    Write-host $string -BackgroundColor Yellow -ForegroundColor Black
}

function cmArguments ($arg, $var, [REF]$finalArgs) {
    if ($null -notlike $var) {
        $finalargs.Value += " $arg `"$var`""
    }
}
function cmBool ($arg, [bool]$var, [REF]$finalArgs) {
    if ($var -eq $true) {
        $finalargs.Value += " $arg"
    }
}

##*===============================================
##* REPOSITORY
##*===============================================
log ("Checking For Output Folder")
# Create Folder in repo for this app\app_verison
if (!(test-path "FileSystem::$outFolder")) {
    New-Item -ItemType Directory "FileSystem::$outFolder"
    log ("Created Output Folder $outFolder")
}
else { log("Folder already exists") }
# Download the app
log ("Checking for Output File")
if (!(Test-Path "FileSystem::$outFile")) {
    log ("File does not exist. Downloading $outFile now")
    # Downloads it, if not there already
    Invoke-WebRequest -Uri $uri -OutFile "FileSystem::$outFile"
    log ("$outFile Download complete")

    # Move downloaded installer to SCCM (PSADT) resource folder
    try {
        Copy-Item "$outFile" -Destination $outSource -Confirm:$false
    }
    catch { break 1 }
}
else { log("File already exists") }
# Borrow app icon
# [System.Drawing.Icon]::ExtractAssociatedIcon("$outFile").toBitmap().Save("$outIco") # this does not accept UNC paths
# So Instead lets grab it from metadata and copy it to an empty filestream, then use that to write the .ico
log ("Checking for icon")
if (!(Test-Path "FileSystem::$outIco")) {
    log ("Icon not found, extracting Icon from $outFile")
    [System.SHFILEINFO]$icoStruct = New-Object System.SHFILEINFO
    $icoSize = [System.Runtime.InteropServices.Marshal]::SizeOf($icoStruct)
    [System.SHGETFILEINFO]::SHGetFileInfo($outFile, 0, [ref]$icoStruct, $icoSize, 0x000000100)
    $icoIcon = [System.Drawing.Icon]::FromHandle($icoStruct.hIcon)
    $icoFile = New-Object System.IO.FileStream($outIco, 'OpenorCreate')
    $icoIcon.Save($icoFile)
    $icoFile.Close()
    $icoIcon.Dispose()
    log ("$outFile icon extracted to $outIco")
}
else { log("Icon already exists") }
##*===============================================
##* CONFIGURATION MANAGER
##*===============================================

# Enter CM Site
log ("Entering configuration Manager Site")
Set-Location $cmSiteCode
log ("Successfully entered $cmSiteCode")

# Create the SCCM application/object
if (get-cmapplication $cmName) {
    log ("An SCCM object with the name $cmName already exists. Please change this and try again. Exiting")
    exit
} 
log ("Building arguments for SCCM object")
cmArguments "-Name" $cmName ([REF]$global:cmAppArgs)
cmArguments "-Description" $cmDescirption ([REF]$global:cmAppArgs)
cmArguments "-LocalizedName" $lAppName ([REF]$global:cmAppArgs)
cmArguments "-Owner" $cmOwner ([REF]$global:cmAppArgs)
cmArguments "-OptionalReference" $cmOptRef ([REF]$global:cmAppArgs)
cmArguments "-LocalizedDescription" $lDescription ([REF]$global:cmAppArgs)
cmArguments "-Publisher" $lPublisher ([REF]$global:cmAppArgs)
cmArguments "-SoftwareVersion" $appVersion ([REF]$global:cmAppArgs)
cmArguments "-ReleaseDate" $(get-date -Format MM/dd/yyyy) ([REF]$global:cmAppArgs)
cmArguments "-Keyword" $lKeywords ([REF]$global:cmAppArgs)
cmArguments "-SupportContact" $cmSupportContact ([REF]$global:cmAppArgs)
cmArguments "-UserDocumentation" $lUserDoc ([REF]$global:cmAppArgs)
cmArguments "-LinkText" $lLinkText ([REF]$global:cmAppArgs)
cmArguments "-IconLocationFile" $lIconPath ([REF]$global:cmAppArgs)
cmArguments "-PrivacyUrl" $lPrivacyUrl ([REF]$global:cmAppArgs)
cmArguments "-OptionalReference" $cmOptRef ([REF]$global:cmAppArgs)
cmArguments "-AppCatalog" $lCatalog ([REF]$global:cmAppArgs)
cmBool "-IsFeatured" $cmFeatured ([REF]$global:cmAppArgs)
cmBool "-AutoInstall" $cmAutoInstall ([REF]$global:cmAppArgs)
# True or False items need to be added manually, to prevent them from resolving to True, I could add a check in the function. Maybe will do this later.
$global:cmAppArgs += " -IsFeatured `$$cmFeatured" + " -AutoInstall `$$cmAutoInstall"
log ("Argument build complete.")
log ("Creating application with the following arguments:
 $global:cmAppArgs")
Invoke-Expression -Command "New-CmApplication $global:cmAppArgs"
log ("Successfully created SCCM object in Applicaiton root")
# Move the SCCM Object
Move-CMObject -FolderPath "$cmAppFolder" -InputObject $(Get-CMApplication -Name "$cmName")
log ("Successfully moved SCCM object $cmName from Application root to $cmAppFolder")
# Set Properties - you can also use this area to edit stuff from existing CMApplicaitons
log ("Building arguments for SCCM object properties.")
cmArguments "-DistributionPointSetting" $pDistPntSet ([REF]$global:cmSetArgs)
cmArguments "-UserCategory" $pUserCat ([REF]$global:cmSetArgs)
log ("Argument build complete.")
log ("Modifying application with the following arguments:
$global:cmSetArgs")
Invoke-Expression -Command "Set-CMApplication -Name `"$cmName`" $global:cmSetArgs"
log ("Successfuly applied new properties")

# Create the Detection Method
<#log ("Creating enhanced detection method.... this is Microsoft's terminology not mine. OK")
$cmDetecClause = New-CMDetectionClauseFile -Value -Path $mPath -FileName $mFile -PropertyType $mPropType -ExpectedValue $mVersion -ExpressionOperator $mExpressOp
log ("Detection method successfully stored with the following parameters:
$cmDetecClause")#>
# Create a Deployment-type (assuming psadt for now - might get messy). For now I will make this seperate and point to it.
log ("Building arguments for SCCM object deployment type")
cmArguments "-ContentLocation" $dContLoc ([REF]$global:cmDepArgs)
cmArguments "-DeploymentTypeName" $dName ([REF]$global:cmDepArgs)
cmArguments "-InstallCommand" $dInstCommand ([REF]$global:cmDepArgs)
cmArguments "-UninstallCommand" $dUninstCommand ([REF]$global:cmDepArgs)
cmArguments "-EstimatedRuntimeMins" $dEstRunTime ([REF]$global:cmDepArgs)
cmArguments "-LogonRequirementType" $dLogonReq ([REF]$global:cmDepArgs)
cmArguments "-MaximumRuntimeMins" $dMaxRunTime ([REF]$global:cmDepArgs)
cmArguments "-InstallWorkingDirectory" $dInstWorkLoc ([REF]$global:cmDepArgs)
cmArguments "-UninstallWorkingDirectory" $dUninstWorkLoc ([REF]$global:cmDepArgs)
cmArguments "-UserInteractionMode" $dUserIntMode ([REF]$global:cmDepArgs)
cmArguments "-InstallationBehaviorType" $dInstBehavior ([REF]$global:cmDepArgs)
cmArguments "-RebootBehavior" $dRebootBhavior ([REF]$global:cmDepArgs)
cmArguments "-UninstallContentLocation" $dUninstContLoc ([REF]$global:cmDepArgs)
cmArguments "-Comment" $dComment ([REF]$global:cmDepArgs)
cmArguments "-UninstallOption" $dUninstOpt ([REF]$global:cmDepArgs)
cmArguments "-SlowNetworkDeploymentMode" $dDownRunLocal ([REF]$global:cmDepArgs)
#cmArguments "-ScriptLanguage" $dDetecScriptType ([REF]$global:cmDepArgs)
#cmArguments "-ScriptText" `"$dDetecScript`" ([REF]$global:cmDepArgs)
cmBool "-RequireUserInteraction" $dReqUserInt ([REF]$global:cmDepArgs)
cmBool "-ContentFallback" $dUseDPBoundaryGroup ([REF]$global:cmDepArgs)
#$global:cmDepArgs += " -AddDetectionClause $cmDetecClause" # Attach the enhanced detection method created above
$global:cmDepArgs += " -ScriptLanguage $dDetecScriptType -ScriptText `'$dDetecScript`'"
log ("Argument build complete")
log ("Creating deployment type with the following arguments:
$global:cmDepArgs")
Invoke-Expression -Command "Get-CMApplication -Name `"$cmName`"|Add-CMScriptDeploymentType  $global:cmDepArgs"
log ("Successfuly Created Deployment Type")
# Distribute Content 
log ("Distributing Content to:
$gName")
Invoke-Expression -Command "Get-CMApplication -Name `"$cmName`"| Start-CMContentDistribution  -DistributionPointGroupName `"$gname`""
log ("Content Distribution Complete")
log ("Taking a one minute break, because I've earned it!")
start-sleep -Seconds 60
log ("OK back to work, what's next?")
# Create deployment (including collection deployed to)
log ("Building arguments for SCCM application deployment")
if ($aDepAction -eq "Uninstall") { $aDepPurp = "Required" } # Error Checking - Available & Uninstall not supported
cmArguments "-CollectionName" $aColName ([REF]$global:cmApdArgs)
cmArguments "-DeployAction" $aDepAction ([REF]$global:cmApdArgs)
cmArguments "-DeployPurpose" $aDepPurp ([REF]$global:cmApdArgs)
cmArguments "-UserNotification" $aUserNotify ([REF]$global:cmApdArgs)
cmArguments "-TimeBaseOn" $aTType ([REF]$global:cmApdArgs)
#cmArguments "-AvailableDateTime" $aAvDT ([REF]$global:cmApdArgs)
cmBool "-ApprovalRequired" $aApprReq ([REF]$global:cmApdArgs)
cmBool "-PreDeploy" $aPreDep ([REF]$global:cmApdArgs)
log ("Argument build complete")
log ("Deploying with the following arguments:
$global:cmApdArgs")
Invoke-Expression -Command "Get-CMApplication -Name `"$cmName`"|New-CMApplicationDeployment $global:cmApdArgs"
log("Application Dpeloyment Complete")

# Remove anything greater than n-1 deployment?

<##*===============================================
##* REPORTING/NOTIFICATION
##*===============================================

# Send email with status 
#Email addresses
$To = "admin@consoso.com"
#Email CCs
#$CC = "User@contoso.com"
#Email subject
$Subject = "OurPCPatcher deployment notification - $appName"
#ASCII hexadecimal equivalents as variables for punctuation characters ( ) %20, (,) %2C, (?) %3F, (.) %2E, (!) %21, (:) %3A, (;) %3B, and (new-line) %0A. Your message can be multiline.
$Body = $("Hello,
Your deployment of $app is now complete. It has been deployed to $aColName
For More information please refer to \\$env:COMPUTERNAME\$($logRoot.replace(':','$'))\$logfileName").Replace(',','%2C').Replace(' ','%20').Replace('?','%3F').Replace('.','%2E').Replace('!','%21').Replace(':','%3A').Replace(';','%3B').Replace("`n",'%0A')
#Message to send via mail app. Tested working using outlook. Notice the ? is escaped with backtick (`)
$Email = "$To&subject=$Subject&body=$Body"
Start-sleep 2
#cmd /c "outlook.exe /a $attachment /m $Email"
#Otlook can't be found in path environment variables.
#if ("$?" -eq "False") {
    #Regex match for Outlook in Program Files and Program Files (x86). Possible, common, likely, default locations for outlook.exe
    $MSO = Get-ChildItem 'C:\Program Files*\Microsoft Office*\*' -Recurse | Where-Object { $_.FullName -match "^[a-zA-Z]:\\Program Files( \(x86\)){0,1}\\Microsoft Office( 1\d\\ClientX(64|86)){0,1}\\((?i)root\\){0,1}Office1\d\\((?i)outlook\.exe)$" }
    #The regex returns exactly one entry
    if ($MSO.Count -eq 1) {
        &$MSO[0].FullName /c ipm.note /m $Email
    }
    if ($MSO.count -ne 1) {
    #The regex returns more than one entry
    if ($MSO.Count -gt 1) {
    #Trim the array into just a list of paths
    $MSO = $MSO | ForEach-Object { $_.FullName } 
    #Show all the paths found for outlook plus the location of the attachment
    Write-Output "    Multiple Outlooks Detected:"$MSO"
    Use: "$archivePath" 
    To send email yourself"| msg * /W /Time:9999999
    }
    #Regex matched nothing on computer, no outlooks found. 
    if ($MSO.Count -eq 0) {
        #Show location of attachment
    Write-Output "Cannot find Outlook." "Navigate to  $archivePath  and send email yourself" | msg * /W /Time:9999999 
    }
Exit
}
Do {Start-Sleep -Seconds 5}
Until (Get-Process -Name Outlook -ErrorAction Ignore)
#Outlook as active window, 
(New-Object -ComObject WScript.Shell).AppActivate((get-process outlook).Id).MainWindowTitle
Start-Sleep 5
#Control Enter to send message. Make sure this is enabled in settings
[Windows.Forms.Sendkeys]::SendWait("^~")
#If above is not enabled this will press yes on the message box
[Windows.Forms.Sendkeys]::SendWait("^~")
Start-Sleep 2
#Used for Secure send checkboxes, if you don't know what secure send is. Ignore this part
[Windows.Forms.Sendkeys]::SendWait("{TAB}")
[Windows.Forms.Sendkeys]::SendWait(" ")
[Windows.Forms.Sendkeys]::SendWait("+{TAB}")
Start-sleep -Seconds 2
#Will send message, Leave this commented out if you wish to review
(New-Object -ComObject Wscript.Shell).Sendkeys("~")
#>
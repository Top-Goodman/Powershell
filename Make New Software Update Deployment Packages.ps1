[CmdletBinding()]
Param (
	[Parameter(Mandatory=$true)]
	[ValidateSet('Phases','Prod')]
	[string]$Deployment = 'Phases',

    [Parameter(Mandatory=$true)]
	[ValidateSet('January','February','March','April','May','June','July','August','September','October','November','December')]
	[string]$Month,

    [Parameter(Mandatory=$true)]
    [ValidatePattern('^20(2[3-9]|[3-9]\d)')] #https://regex101.com/r/FgeMMW/1
	[string]$year
)

switch ($Month)
{
    'January' {$MonthNum="01"}
    'February'{$MonthNum="02"}
    'March'{$MonthNum="03"}
    'April'{$MonthNum="04"}
    'May'{$MonthNum="05"}
    'June'{$MonthNum="06"}
    'July'{$MonthNum="07"}
    'August'{$MonthNum="08"}
    'September'{$MonthNum="09"}
    'October'{$MonthNum="10"}
    'November'{$MonthNum="11"}
    'December' {$MonthNum="12"} 
}
$psSR=(Get-Location).path


###########################
#    Connect to Server    #
#Run this before each step#
###########################

$SiteCode = "" # Site code 
$ProviderMachineName = "" # SMS Provider machine name

$initParams = @{}
if($null -eq (Get-Module ConfigurationManager)) {
    Import-Module "$($ENV:SMS_ADMIN_UI_PATH)\..\ConfigurationManager.psd1" @initParams 
}
if($null -eq (Get-PSDrive -Name $SiteCode -PSProvider CMSite -ErrorAction SilentlyContinue)) {
    New-PSDrive -Name $SiteCode -PSProvider CMSite -Root $ProviderMachineName @initParams
}
Set-Location "$($SiteCode):\" @initParams

if ($Deployment -ine "Prod") {
########################
#Phase 1 - 3 ADR change#
#  Only run on Server  #
########################

<#$year="2022"
$MonthNum="05"
$Month="May"#>
$path="filesystem::\\server\share$\SoftwareUpdateDeploymentPackages\$year"
       
New-Item -path "$path" -Name "$MonthNum-$Month" -ItemType Directory 

$path="$path\$MonthNum-$Month"

$packages= @()
$packages+= new-object -typename psobject -property @{Name="M365_SAE"; Path="\\server\share$SoftwareUpdateDeploymentPackages\$year\$MonthNum-$Month\M365_SAE"}
$packages+= new-object -typename psobject -property @{Name="M365_SAEP"; Path="\\server\share$\SoftwareUpdateDeploymentPackages\$year\$MonthNum-$Month\M365_SAEP"}
$packages+= new-object -typename psobject -property @{Name="Windows_10"; Path="\\server\share$\SoftwareUpdateDeploymentPackages\$year\$MonthNum-$Month\Windows_10"}
$packages+= new-object -typename psobject -property @{Name="Windows_11"; Path="\\server\share$\SoftwareUpdateDeploymentPackages\$year\$MonthNum-$Month\Windows_11"}
$packages+= new-object -typename psobject -property @{Name="DotNet_MSRT"; Path="\\server\share$\SoftwareUpdateDeploymentPackages\$year\$MonthNum-$Month\DotNet_MSRT"}

foreach($package in $packages){
    $Name=$package.Name
    $PackagePath=$package.Path
    new-item -path "$path" -name "$Name" -ItemType Directory | Out-Null
    New-CMSoftwareUpdateDeploymentPackage -Name "Desktop-$Year-$Month-$Name" -Path "$PackagePath" | Out-Null
    Write-Host "Desktop-$Year-$Month-$Name Created!" -BackgroundColor Blue -ForegroundColor Cyan
    Start-CMContentDistribution -DeploymentPackageName "Desktop-$Year-$Month-$Name" -DistributionPointGroupName "On-Premises DPs"
    Write-Host "Desktop-$Year-$Month-$Name Distributed!" -BackgroundColor Blue -ForegroundColor Yellow
}

$ADRs = Get-CMSoftwareUpdateAutoDeploymentRule -Fast

foreach($ADR in $ADRs){
    if($ADR.Name -like "Desktop - Phases*"){
        if($ADR.Name -like "*Windows 10"){
            Set-CMSoftwareUpdateAutoDeploymentRule -Name $ADR.Name -DeploymentPackageName "Desktop-$Year-$Month-Windows_10"
            Write-host "Desktop-Phases-Windows-10 Changed"#Write-host "Desktop-Phase-1-Windows-10 Changed"
        }
        if($ADR.Name -like "*Windows 11"){
            Set-CMSoftwareUpdateAutoDeploymentRule -Name $ADR.Name -DeploymentPackageName "Desktop-$Year-$Month-Windows_11"
            Write-host "Desktop-Phases-Windows-11 Changed"#Write-host "Desktop-Phase-1-Windows-10 Changed"
        }
        elseif($ADR.Name -like "*dotNET MSRT"){
            Set-CMSoftwareUpdateAutoDeploymentRule -Name $ADR.Name -DeploymentPackageName "Desktop-$Year-$Month-DotNet_MSRT"
            Write-host "Desktop-Phases-DotNet-MSRT Changed" #Write-host "Desktop-Phase-1-DotNet-MSRT Changed"
        }
        elseif($ADR.Name -like "*M365 SAEP"){
            Set-CMSoftwareUpdateAutoDeploymentRule -Name $ADR.Name -DeploymentPackageName "Desktop-$Year-$Month-M365_SAEP"
            Write-host "Desktop-Phases-M365-SAEP Changed"#Write-host "Desktop-Phase-1-M365-SAEP Changed"
        }
        elseif($ADR.Name -like "*M365 SAE"){
            Set-CMSoftwareUpdateAutoDeploymentRule -Name $ADR.Name -DeploymentPackageName "Desktop-$Year-$Month-M365_SAE"
            Write-host "Desktop-Phases-M365-SAE Changed"#Write-host "Desktop-Phase-1-M365-SAE Changed"
        }
    }
    <#
    elseif($ADR.Name -like "Desktop - P1*"){
        if($ADR.Name -like "*Windows 10"){
            Set-CMSoftwareUpdateAutoDeploymentRule -Name $ADR.Name -DeploymentPackageName "Desktop-$Year-$Month-Windows_10"
            Write-host "Desktop-Phase-1-Windows-10 Changed"
        }
        elseif($ADR.Name -like "*dotNET MSRT"){
            Set-CMSoftwareUpdateAutoDeploymentRule -Name $ADR.Name -DeploymentPackageName "Desktop-$Year-$Month-DotNet_MSRT"
            Write-host "Desktop-Phase-1-DotNet-MSRT Changed"
        }
        elseif($ADR.Name -like "*M365 SAEP"){
            Set-CMSoftwareUpdateAutoDeploymentRule -Name $ADR.Name -DeploymentPackageName "Desktop-$Year-$Month-M365_SAEP"
            Write-host "Desktop-Phase-1-M365-SAEP Changed"
        }
        elseif($ADR.Name -like "*M365 SAE"){
            Set-CMSoftwareUpdateAutoDeploymentRule -Name $ADR.Name -DeploymentPackageName "Desktop-$Year-$Month-M365_SAE"
            Write-host "Desktop-Phase-1-M365-SAE Changed"
        }
        
    }
    elseif($ADR.Name -like "Desktop - P2*"){
        if($ADR.Name -like "*Windows 10"){
            Set-CMSoftwareUpdateAutoDeploymentRule -Name $ADR.Name -DeploymentPackageName "Desktop-$Year-$Month-Windows_10"
            Write-host "Desktop-Phase-2-Windows-10 Changed"
        }
        elseif($ADR.Name -like "*dotNET MSRT"){
            Set-CMSoftwareUpdateAutoDeploymentRule -Name $ADR.Name -DeploymentPackageName "Desktop-$Year-$Month-DotNet_MSRT"
            Write-host "Desktop-Phase-2-DotNet-MSRT Changed"
        }
        elseif($ADR.Name -like "*M365 SAEP"){
            Set-CMSoftwareUpdateAutoDeploymentRule -Name $ADR.Name -DeploymentPackageName "Desktop-$Year-$Month-M365_SAEP"
            Write-host "Desktop-Phase-2-M365-SAEP Changed"
        }
        elseif($ADR.Name -like "*M365 SAE"){
            Set-CMSoftwareUpdateAutoDeploymentRule -Name $ADR.Name -DeploymentPackageName "Desktop-$Year-$Month-M365_SAE"
            Write-host "Desktop-Phase-2-M365-SAE Changed"
        }
        
    }
    elseif($ADR.Name -like "Desktop - P3*"){
        if($ADR.Name -like "*Windows 10"){
            Set-CMSoftwareUpdateAutoDeploymentRule -Name $ADR.Name -DeploymentPackageName "Desktop-$Year-$Month-Windows_10"
            Write-host "Desktop-Phase-3-Windows-10 Changed"
        }
        elseif($ADR.Name -like "*dotNET MSRT"){
            Set-CMSoftwareUpdateAutoDeploymentRule -Name $ADR.Name -DeploymentPackageName "Desktop-$Year-$Month-DotNet_MSRT"
            Write-host "Desktop-Phase-3-DotNet-MSRT Changed"
        }
        elseif($ADR.Name -like "*M365 SAEP"){
            Set-CMSoftwareUpdateAutoDeploymentRule -Name $ADR.Name -DeploymentPackageName "Desktop-$Year-$Month-M365_SAEP"
            Write-host "Desktop-Phase-3-M365-SAEP Changed"
        }
        elseif($ADR.Name -like "*M365 SAE"){
            Set-CMSoftwareUpdateAutoDeploymentRule -Name $ADR.Name -DeploymentPackageName "Desktop-$Year-$Month-M365_SAE"
            Write-host "Desktop-Phase-3-M365-SAE Changed"
        }
        
    }#>
}
}
elseif ($Deployment -eq "Prod") {
#######################
#Production ADR change#
# Only run on Server  #
#######################

#$year="2022"
#$Month="May"
$ADRs = Get-CMSoftwareUpdateAutoDeploymentRule -Fast

foreach($ADR in $ADRs){
    if($ADR.Name -like "Desktop - Production*"){
        
        if($ADR.Name -like "*Windows 10"){
            Set-CMSoftwareUpdateAutoDeploymentRule -Name $ADR.Name -DeploymentPackageName "Desktop-$Year-$Month-Windows_10"
            Write-host "Desktop-Production-Windows-10 Changed"
        }
        if($ADR.Name -like "*Windows 11"){
            Set-CMSoftwareUpdateAutoDeploymentRule -Name $ADR.Name -DeploymentPackageName "Desktop-$Year-$Month-Windows_11"
            Write-host "Desktop-Production-Windows-11 Changed"
        }
        elseif($ADR.Name -like "*dotNET MSRT"){
            Set-CMSoftwareUpdateAutoDeploymentRule -Name $ADR.Name -DeploymentPackageName "Desktop-$Year-$Month-DotNet_MSRT"
            Write-host "Desktop-Production-DotNet-MSRTd Changed"
        }
        
    }
}
}
Set-Location $psSR
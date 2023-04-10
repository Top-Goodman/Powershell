<## Instructions
1. Get AppID and Secret.
    AppID is from created application. Make sure to grant it needed permissions if it is a custom app.
    Secret is generated from Certificates & Secrets section of App Registrations. Make sure permissions are approved while here.
2. Generate Secrert Hash (can modify key if desired, just make sure whatever is used to generate, matches below to decrypt).
    Each time Hash is generated with below PowerShell commands it will differ, key will need to be remembered as it is needed to decrypt later
    Regenerating the hash does NOT mean that it needs to be replaced. As long as same key is used, will decrypt as expected.
    If the secret changes (for example KeyVault expire) Hash will need to be regenerated.
    $AccessSecret = "" # Secret Value property from Azure App Registration (or Keyvault if youhave access to that in your subscription).
    $ClientSecretHash = ConvertTo-SecureString -String "$AccessSecret" -AsPlainText -Force|ConvertFrom-SecureString -Key $Key # Key must match below for decrypt
    This multi line commented section should be deleted before using\saving\deploying this script. Paste the value below as $ClientSecretHash, it will be a long string, that is expected
3. Use Function to send API request (Method is Get if not specified and version is v1.0 when not specified)
    For more information on query parameters: https://learn.microsoft.com/en-us/graph/use-the-api
    Be cautious when using beta, these components can change or be removed without warning via MS
    Make sure to run 'Get' method first to make sure other methods are targetting what you expect them to (also 'get' might be needed for attributes to use with other methods)
    Examples below. First command will output selected properties of 2fk6dh2-devmXX to a variable, then second command uses that data to delete that computer from Intune
    $IntuneDevice = Invoke-REST -Resource "deviceManagement/managedDevices" -parm1 "`$filter=deviceName eq '2fk6dh2-devmXX'" -parm2 "`$select=id,deviceName,serialNumber"
    Invoke-REST -Method "Delete" -resource "/deviceManagement/managedDevices/$($IntuneDevice.id)" 
4. Test that API call is functioning as expected.
    If making multiple API calls, you should not need to run Get-MsalToken each time. I beleive the tokens are good for somewhere between 8-24 hours.
    If commands work and you start to get error 401 it could be that the token expired. Rerunning the Get-MsalToken command will alleviate this 
    If you instantly get 401, permissions of ent app could be bad, secret could be wrong or maybe some other azure related issue.
5. Delete these instructions.
#>

## Functions
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

Function Invoke-REST {
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory = $false)]
        [ValidateSet("Get", "Post", "Put", "Patch", "Delete")]
        [string]$Method = "Get",
        [Parameter(Mandatory = $false)]
        [ValidateSet("v1.0", "beta")]
        [string]$Version = "v1.0",
        [Parameter(Mandatory = $true)]
        [ValidateNotNullorEmpty()]
        [string]$Resource,
        [Parameter(Mandatory = $true)]
        [string]$HTTP=https://graph.microsoft.com,
        [Parameter(Mandatory = $false)]
        [string]$parm1,
        [Parameter(Mandatory = $false)]
        [string]$parm2,
        [Parameter(Mandatory = $false)]
        [string]$parm3,
        [Parameter(Mandatory = $false)]
        [string]$parm4
    )
    $URI = "$HTTP/$Version/$Resource"
    $Query = "?"
    ($parm1, $parm2, $parm3, $parm4)|ForEach-Object {
    if ($null -ne $_) {
        $Query += $_ 
    }
    if ($Query -ne "?") {$Query += "&"}
    }
    $Query = $($Query -replace "&{0,}$")
    if ($Query -ne "?") {
        $URI += "$Query"
    }
    (Invoke-RestMethod -Method $Method -Uri $uri -Headers @{Authorization = "Bearer $($MsalToken.AccessToken)" }).value
}

## Variables
$AppId = "<Get from Azure>" 
$TenantId = "<Get from Azure>"
# 256-bit encryption key (32bytes)
$Key =  # You can change this if you want, just make sure it is either 128 bits, 192 bits, or 256 bits. 16,24, or 32 bytes. I like using an array of 32 numbers. For example (3..34).
$ClientSecretHash = "" #Paste string here after generating. Make sure value being used to Decrypt ($Key) is the same


## Dependencies
if (Get-Module -ListAvailable -Name msal.ps) {
}
else { 
    Install-Package msal.ps -Scope CurrentUser -Force 
}
# Get Authentication Token using tanant, app and decrypted secret.
$MsalToken = Get-MsalToken -TenantId $TenantId -ClientId $AppId -ClientSecret $([System.Runtime.InteropServices.Marshal]::PtrToStringAuto($([System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($(ConvertTo-SecureString -String "$ClientSecretHash" -Key $Key)))) | ConvertTo-SecureString -AsPlainText -Force)

# Method "Get" and Version "v1.0" are assumed unless otherwise specified.
Invoke-REST -Resource "" -parm1 ""

[CmdletBinding()]
Param (
	[Parameter(Mandatory = $false)]
	[string]$Database = 'CM_',
	[Parameter(Mandatory = $false)]
	[string]$Server = 'server' ,
	[Parameter(Mandatory = $false)]
	[string]$Query = "SELECT DISTINCT 
    dpgr.NAME [DP Group],
    pk.NAME [Package Name],
    dgp.pkgid [Package ID],
    dpcn.targeteddpcount as Targeted,
    dpcn.numberinstalled as [Distributed],
    dpcn.numberinprogress as [In Progress],
    dpcn.numbererrors as [Error(s)],
    CASE
    WHEN pk.packagetype = 0 THEN 'Software Distribution Package'
    WHEN pk.packagetype = 3 THEN 'Driver Package'
    WHEN pk.packagetype = 4 THEN 'Task Sequence Package'
    WHEN pk.packagetype = 5 THEN 'Software Update Package'
    WHEN pk.packagetype = 6 THEN 'Device Setting Package'
    WHEN pk.packagetype = 7 THEN 'Virtual Package'
    WHEN pk.packagetype = 8 THEN 'Application'
    WHEN pk.packagetype = 257 THEN 'Image Package'
    WHEN pk.packagetype = 258 THEN 'Boot Image Package'
    WHEN pk.packagetype = 259 THEN 'Operating System Install Package'
    ELSE 'Unknown'
    END AS 'Package Type'
    FROM vsms_dpGroupInfo dpgr
    INNER JOIN v_dpgrouppackages dgp ON dgp.groupid = dpgr.groupid
    LEFT JOIN v_package pk ON pk.packageid = dgp.pkgid
    LEFT JOIN v_dpgroupcontentdetails dpcn ON dpcn.groupid = dpgr.groupid
    AND dpcn.pkgid = pk.packageid
    where NOT dpcn.targeteddpcount = dpcn.numberinstalled"
)
## Dependencies
if (Get-Module -ListAvailable -Name sqlserver) {
Install-Module sqlserver
}
## Prerequisites
if ("sqlserver" -in (Get-Module).Name) {
    Import-Module "sqlserver"
}

## Begin
Invoke-Sqlcmd -ServerInstance "$Server" -Database "$Database" -Query "$Query"|Format-Table

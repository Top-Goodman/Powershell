# If not specified, script will assume path to iso will be provided using InstallationISO argument.
[CmdletBinding(DefaultParameterSetName = 'Iso')]
Param (  
    # Main mandatory selection. Starting point of VM. Either another VHDx or an Iso
    [Parameter(ParameterSetName = 'Iso', Mandatory = $true, ValueFromPipelineByPropertyName = $true, ValueFromPipeline = $true)]
    [Parameter(ParameterSetName = 'VHD', Mandatory = $true, ValueFromPipelineByPropertyName = $true, ValueFromPipeline = $true)]
    [Parameter(ParameterSetName = 'Clone', Mandatory = $true, ValueFromPipelineByPropertyName = $true, ValueFromPipeline = $true)]

    # Comma seperated list of VMs to make. If using Iso, will add same iso to all. If using VHD, will use same original vhdx
    [Parameter(Mandatory = $true)]
    [array]$VMNames,

    # Path to virtual hard disk to use. Regex just checks that path is mapped to a letter, and file type is vhdx
    [Parameter(ParameterSetName = 'VHD', Mandatory = $true)]
    [ValidatePattern('^[a-zA-Z]\:\\.*\.vhdx$')]
    [string]$BaseVHDX,

    [Parameter(ParameterSetName = 'VHD')]
    [switch]$VHD = $false,

    # Path to disk image to use. Regex just checks that path is mapped to a letter, and file type is iso
    [Parameter(ParameterSetName = 'Iso', Mandatory = $true)]
    [ValidatePattern('^[a-zA-Z]\:\\.*\.iso$')]
    [string]$InstallationISO,

    [Parameter(ParameterSetName = 'Iso')]
    [switch]$Iso = $false,

    [Parameter(ParameterSetName = 'Clone')]
    [switch]$Clone = $false,

    # Name of Virtual Switch to apply to created VMs, default value will be overridden if specified during script call
    [Parameter(Mandatory = $false)]
    [string]$SwitchName = 'Default Switch',

    # Path of where to save created VMs, default value will be overridden if specified during script call
    [Parameter(Mandatory = $false)]
    [string]$MachinePath = 'C:\Users\Public\Documents\Hyper-V\Virtual Machines\',

    # Path of where to save created Virtual Disk, default value will be overridden if specified during script call
    [Parameter(Mandatory = $false)]
    [string]$DiskPath = 'C:\Users\Public\Documents\Hyper-V\Virtual Hard Disks',
    
    # Smart page file is only used at boot to try and mitigate over allocation of RAM. Not same as Pagefile.sys
    [Parameter(Mandatory = $false)]
    [string]$SmartPagingFilePath = 'C:\Users\Public\Documents\Hyper-V\',

    # Number to use for RAM. Can be in "Human Readable" file size notation 4096MB or 4GB. Regex is checking that a number is entered and may or may not include m,g followed by a b case insensitive.
    [Parameter(Mandatory = $false)]
    [ValidatePattern('\d{1,}(((G|g)|(M|m))(B|b)){0,1}')]
    [int64]$MemoryStartupGB = 4GB,

    # Checking that at least one number is entered
    [Parameter(Mandatory = $false)]
    [ValidatePattern('\d{1,}')]
    [string]$ProcessorCount = '4'
)

# Loop through array
foreach ($VMName in $VMNames) {
    write-host "Preparing "$VMName 
    if ($VHD) {
        # Copy specified base
        Copy-Item $BaseVHDX -Destination "$diskpath\$vmname.vhdx"
        # Make new VM using that base (Not sure why, this doens't actually work as cloning. Yet it does seem to at least apply an OS that was used in original)
        New-VM -Name $vmname -Path $machinepath -SwitchName $SwitchName -Generation 2 -VHDPath "$diskpath\$vmname.vhdx" -MemoryStartupBytes $MemoryStartupGB | Out-Null
    } 
    elseif ($Iso) {
        # New machine with speicifed parameters
        New-VM -Name $vmname -Path $machinepath -SwitchName $SwitchName -Generation 2 -MemoryStartupBytes $MemoryStartupGB | Out-Null
        # Add the iso to it
        Add-VMDvdDrive -VMName $vmname -Path $InstallationISO
        # Get Firmware boot information
        $bootorder = (Get-VMFirmware -VMName $vmname).bootorder | Sort-Object -Property Device
        # Refresh list to make sure applied iso is available
        Get-VM -VMName $vmname | Set-VMFirmware -BootOrder $bootorder
    }
    elseif ($clone) {
        ## Future stuff planned that I didn't get to yet. Will use import\export commands and checks to shutdown if running and optionally make snapshot before doing all this.
    }
    # Prestage Virtual TPM
    Set-VMKeyProtector -VMName $VMName -NewLocalKeyProtector | Out-Null
    # Enable Virtual TPM
    Enable-VMTPM -VMName $VMName | Out-Null
    # Turn off Auto Checkpoints and apply other hardware settings
    Set-VM -Name $VMName -AutomaticCheckpointsEnabled $false -ProcessorCount $ProcessorCount -SmartPagingFilePath $SmartPagingFilePath -staticmemory | Out-Null
    # Make a checkpoint becuase it never hurts to be able to revert
    Checkpoint-VM -Name $VMName -SnapshotName 'Begin' | Out-Null
    # Power on the VM
    Start-VM -Name $vmname | Out-Null
    write-host "Starting  "$VMName 
}
# Output all local VM info because why not
Get-VM
# Network Path
$Path2Map = "\\pathTo\Network\Share<replace me>"
# Get "Drive Letters C: - Z: (I know A and B exist, yet don't want to bother since this could list them left of C: drive)
$Letters = 67..90 | ForEach-Object { [char]$_ + ":" }
# Get Currently Mapped Drive Letters
$Mapped = Get-WmiObject win32_logicaldisk | Select-Object DeviceID, ProviderName 
$InUse = $Mapped.DeviceID
$NetProvider = $Mapped.ProviderName
# Make sure desired Network Path is not already mapped
forEach ($nPath in $NetProvider) {
    if ($npath -eq $Path2Map) {
        Write-Output "Share already mapped... Exiting"
        #exit 0
        Return
    }
}
# Network Path not found already mapped to system
Write-Output "$path2Map not found. Continue to Mapping"
# Test each letter
ForEach ($Letter in $Letters) {
    # Check is letter is mapped
    if ($InUse -contains $Letter) {
        # Letter is in use
        Write-Output "$Letter is in use"
    } 
    Else {
        # Found an unused letter
        Write-Output "$Letter is available"
        # Save it for use later
        $FreeLetter = $Letter
        # Exit Loop
        break
    }
}
# Map network share to found free letter
net use $freeletter $Path2Map
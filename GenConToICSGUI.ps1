Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName System.Web
Add-Type -AssemblyName System.Windows.Forms

# ---------------- GEN CON DATE CALCULATION ----------------
$year = (Get-Date).Year
$aug4 = Get-Date -Year $year -Month 8 -Day 4
$daysBack = ([int]$aug4.DayOfWeek - 3 + 7) % 7
$genConStart = $aug4.AddDays(-$daysBack)

$GenConDays = @{
    "Wednesday" = $genConStart
    "Thursday"  = $genConStart.AddDays(1)
    "Friday"    = $genConStart.AddDays(2)
    "Saturday"  = $genConStart.AddDays(3)
    "Sunday"    = $genConStart.AddDays(4)
}

# ---------------- VENUE MAP ----------------

$VenueMap = @{
    "JW" = "JW Marriott Indianapolis, 10 S West St, Indianapolis, IN 46204, USA"
    "Hyatt" = "Hyatt Regency Indianapolis, 1 S Capitol Ave, Indianapolis, IN 46204, USA"
    "Stadium" = "Lucas Oil Stadium, 500 S Capitol Ave, Indianapolis, IN 46225, USA"
    "Westin" = "The Westin Indianapolis, 241 W Washington St, Indianapolis, IN 46204, USA"
    "Hilton" = "Hilton Indianapolis Hotel & Suites, 120 W Market St, Indianapolis, IN 46204, USA"
    "Embassy" = "Embassy Suites by Hilton Indianapolis Downtown, 110 W Washington St, Indianapolis, IN 46204"
    "Crowne" = "Crowne Plaza Indianapolis Downtown Union Station, 123 W Louisiana St, Indianapolis, IN 46225"
    "ICC" = "Indianapolis Convention Center, 100 S Capitol Ave, Indianapolis, IN 46225"
    "Omni" = "Omni Severin Hotel, 40 W Jackson Pl, Indianapolis, IN 46225"

}

$EventCache = @{}
# ---------------- Helper: Build DateTime ----------------

function Build-DateTime($string) {

    if ($string -match '^(?<day>\w+),\s+(?<time>[\d:]+\s+[APM]+)') {

        $dayName = $Matches['day']
        $timePart = $Matches['time']

        if (-not $GenConDays.ContainsKey($dayName)) { return $null }

        $date = $GenConDays[$dayName]
        $combined = "$($date.ToString('yyyy-MM-dd')) $timePart"

        $eastern = [System.TimeZoneInfo]::FindSystemTimeZoneById("Eastern Standard Time")
        $local = [datetime]::Parse($combined)
        $utc = [System.TimeZoneInfo]::ConvertTimeToUtc($local, $eastern)

        return $utc.ToString("yyyyMMddTHHmmssZ")
    }

    return $null
}

# ---------------- Parse Function ----------------

function Parse-Event($url) {

    if ($url -match '(?<EventCode>\d+)$') {$url = "https://www.gencon.com/events/$($Matches.EventCode)"}

    try {
        $response = Invoke-WebRequest -Uri $url -UseBasicParsing -Headers @{ "User-Agent" = "Mozilla/5.0 (Windows NT 10.0; Win64; x64)" }

        $html = $response.Content

        $attrBlocks = [regex]::Matches(
            $html,
            '<div class=''attr''>(.*?)</div>\s*</div>',
            'Singleline'
        )

        $EventData = @{}

        foreach ($block in $attrBlocks) {

            $nameMatch  = [regex]::Match($block.Value,"<div class='name'>(.*?)</div>",'Singleline')
            $valueMatch = [regex]::Match($block.Value,"<div class='value-full'>(.*?)</div>",'Singleline')
            
            if ($nameMatch.Success -and $valueMatch.Success) {

                $rawName  = [System.Web.HttpUtility]::HtmlDecode($nameMatch.Groups[1].Value)
                $rawValue = [System.Web.HttpUtility]::HtmlDecode($valueMatch.Groups[1].Value)

                $name = ($rawName -replace '<.*?>','' -replace '\s+', ' ').Trim().TrimEnd(':')
                $value = ($rawValue -replace '<.*?>','' -replace '\s+', ' ').Trim()

                $EventData[$name] = $value
            }
        }
        $EventData["URL"] = $url
        $EventData["Event Code"] = $Matches.EventCode
        return $EventData
    }
    catch {
        [System.Windows.MessageBox]::Show("Failed to fetch: $($_.Exception.Message)")
        return $null
    }
}


function Get-ImageFromBase64([string]$base64) {
    $bytes = [Convert]::FromBase64String($base64)
    $ms = New-Object System.IO.MemoryStream(,$bytes)
    $img = New-Object System.Windows.Media.Imaging.BitmapImage
    $img.BeginInit()
    $img.StreamSource = $ms
    $img.CacheOption = [System.Windows.Media.Imaging.BitmapCacheOption]::OnLoad
    $img.EndInit()
    $img
}

# ---------------- GUI ----------------

$XAML = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        Title="GenCon ICS Builder"
        Height="650"
        Width="1000"
        WindowStartupLocation="CenterScreen"
        Background="Black"
        Foreground="White">

    <Grid Margin="10">
        <Grid.ColumnDefinitions>
            <ColumnDefinition Width="350"/>
            <ColumnDefinition Width="*"/>
        </Grid.ColumnDefinitions>

        <!-- Left panel -->
        <Grid Grid.Column="0" Margin="0,0,10,0">
            <Grid.RowDefinitions>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="*"/>
                <RowDefinition Height="Auto"/>
            </Grid.RowDefinitions>

            <!-- URL input -->
            <StackPanel Orientation="Horizontal">
                <TextBox Name="UrlInputBox" Width="260" Background="#111111" Foreground="White" BorderBrush="#E9FE60" BorderThickness="1" Padding="4" ToolTip="Enter a GenCon Event URL (ie: https://www.gencon.com/events/327174), GameID (ie: RPG26ND327174), or Event Code (ie: 327174)" Text="URL, GameID, or Event Code"/>
                <Button Name="AddUrlBtn" Width="60" Margin="5,0,0,0" Background="#E9FE60" Foreground="Black" FontWeight="Bold" Padding="6,4" BorderThickness="0">Add</Button>
            </StackPanel>

            <Button Name="RemoveUrlBtn" Grid.Row="1" Margin="0,10,0,5" Background="#E9FE60" Foreground="Black" FontWeight="Bold" Padding="6,4" BorderThickness="0">Remove Selected</Button>

            <ListBox Name="UrlListBox" Grid.Row="2" Margin="0,5,0,10" Background="#111111" Foreground="White" BorderBrush="#E9FE60" BorderThickness="1"/>

            <Button Name="GenerateBtn" Grid.Row="3" Height="32" Background="#E9FE60" Foreground="Black" FontWeight="Bold" Padding="6,4" BorderThickness="0">Generate ICS</Button>
        </Grid>

        <!-- Preview -->
<GroupBox Header="Event Preview" Grid.Column="1" Foreground="#E9FE60" BorderBrush="#E9FE60" BorderThickness="1" Padding="5" FontSize="14">
    <Grid Name="PreviewGrid">
        <!-- Background Image -->
        <Image Name="PreviewBg"
               Stretch="UniformToFill"
               Opacity="0.25" />
        
        <!-- TextBox sits on top -->
        <TextBox Name="PreviewBox"
                 IsReadOnly="True"
                 TextWrapping="Wrap"
                 AcceptsReturn="True"
                 Background="Transparent"
                 Foreground="White"
                 BorderThickness="0"
                 FontSize="14"
                 Padding="5"/>
    </Grid>
</GroupBox>
    </Grid>
</Window>
"@

$Window = [Windows.Markup.XamlReader]::Parse($XAML)

$UrlInputBox = $Window.FindName("UrlInputBox")
$AddUrlBtn = $Window.FindName("AddUrlBtn")
$RemoveUrlBtn = $Window.FindName("RemoveUrlBtn")
$UrlListBox = $Window.FindName("UrlListBox")
$PreviewBox = $Window.FindName("PreviewBox")
$GenerateBtn = $Window.FindName("GenerateBtn")

#-------------- Icon Base64 --------------
$base64Icon = "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"

$iconImg = Get-ImageFromBase64 $base64Icon
$Window.Icon = $iconImg


# Find the Image control
$PreviewBg = $Window.FindName("PreviewBg")

# Assign your Base64 icon image
$PreviewBg.Source = $iconImg

<# Optional: add a slight blur so it's even less distracting
$blur = New-Object System.Windows.Media.Effects.BlurEffect
$blur.Radius = 4
$PreviewBg.Effect = $blur
#>
# ---------------- Events ----------------

$AddUrlBtn.Add_Click({
    $url = $UrlInputBox.Text.Trim()
    if (-not $url) { return }
    if ($EventCache.ContainsKey($url)) { return }

    $data = Parse-Event $url
    if ($data) {
        $EventCache[$url] = $data
        $UrlListBox.Items.Add($url)
        $UrlInputBox.Clear()
    }
})

$RemoveUrlBtn.Add_Click({
    if ($UrlListBox.SelectedItem) {
        $EventCache.Remove($UrlListBox.SelectedItem)
        $UrlListBox.Items.Remove($UrlListBox.SelectedItem)
        $PreviewBox.Clear()
    }
})

$UrlListBox.Add_SelectionChanged({
    if ($UrlListBox.SelectedItem) {

        $data = $EventCache[$UrlListBox.SelectedItem]
        $previewLines = @()

        foreach ($key in $data.Keys | Sort-Object) {

            $value = $data[$key]

            $previewLines += "${key}:"
            $previewLines += "$value"
            $previewLines += ""
        }

        $PreviewBox.Text = $previewLines -join "`r`n"
    }
})

$GenerateBtn.Add_Click({

    if ($EventCache.Count -eq 0) { return }

    $dlg = New-Object System.Windows.Forms.SaveFileDialog
    $dlg.Filter = "ICS files (*.ics)|*.ics"

    if ($dlg.ShowDialog() -ne "OK") { return }

    $OutputPath = $dlg.FileName

    $Lines = @()
    $Lines += "BEGIN:VCALENDAR"
    $Lines += "VERSION:2.0"
    $Lines += "PRODID:-//GenCon ICS Generator//EN"
    $Lines += "CALSCALE:GREGORIAN"
    $Lines += "METHOD:PUBLISH"

    foreach ($url in $EventCache.Keys) {

        $EventData = $EventCache[$url]

        if (-not $EventData["Title"]) { continue }

        $startUTC = Build-DateTime $EventData["Start Date & Time"]
        $endUTC   = Build-DateTime $EventData["End Date & Time"]

        if (-not $startUTC -or -not $endUTC) { continue }

        $title = $EventData["Title"]
        $locationRaw = $EventData["Location"]
        $mapAddress = $locationRaw.Trim()

        foreach ($venue in $VenueMap.Keys) {
            if ($locationRaw -like "*$venue*") {
                $mapAddress = $VenueMap[$venue]
                break
            }
        }

        $mapAddress = $mapAddress -replace ',', '\,'

        $descParts = @()

        if ($EventData["Game System"]) {
            $descParts += "Game System: $($EventData["Game System"].Trim())"
        }

        if ($locationRaw) {
            $roomText = ($locationRaw -replace "`r?`n", ", " -replace '\s+', ' ').Trim()
            $descParts += "Room: $roomText"
        }

        $descParts += "URL: $($EventData["URL"])"

        foreach ($field in @("Description","Short Description","Long Description")) {
            if ($EventData[$field] -and $EventData[$field].Trim() -ne "") {
                $descParts += "${field}:`n$($EventData[$field].Trim())"
            }
        }

        $description = ($descParts -join "`n`n")
        $description = $description -replace "`r?`n", "\n"
        $description = $description -replace ',', '\,'

        $Lines += "BEGIN:VEVENT"
        $Lines += "UID=" + [guid]::NewGuid()
        $Lines += "DTSTAMP=" + (Get-Date -Format "yyyyMMddTHHmmssZ")
        $Lines += "DTSTART:$startUTC"
        $Lines += "DTEND:$endUTC"
        $Lines += "SUMMARY:$title"
        $Lines += "LOCATION:$mapAddress"
        $Lines += "DESCRIPTION:$description"
        $Lines += "END:VEVENT"
    }

    $Lines += "END:VCALENDAR"

    [System.IO.File]::WriteAllLines($OutputPath, $Lines)

    [System.Windows.MessageBox]::Show("ICS created successfully.")
})

$Window.ShowDialog() | Out-Null

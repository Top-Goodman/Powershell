Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName System.Windows.Forms

# XML for GUI
[xml]$XAMLWindow = @"
<Window
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        Title="Lunch" Height="450" Width="800" Background="#FF491573">
        <Grid>
        <Label Name="lblTitle" Content="What's For Lunch?" HorizontalAlignment="Left" Margin="104,10,0,0" VerticalAlignment="Top" FontSize="72" Foreground="White"/>
        <ListBox Name="lstGenre" Margin="79,121,421,125"/>
        <ListBox Name="lstItem" Margin="421,121,79,125"/>
        <Button Name="btnRGenre" Content="Randomize" HorizontalAlignment="Left" Margin="197,341,0,0" VerticalAlignment="Top" RenderTransformOrigin="-0.05,0.303"/>
        <Button Name="btnRItem" Content="Randomize" HorizontalAlignment="Left" Margin="539,341,0,0" VerticalAlignment="Top" RenderTransformOrigin="-0.05,0.303"/>
        <TextBlock Name="tbSelection" HorizontalAlignment="Center" Margin="0,335,0,0" TextWrapping="Wrap" Text="" VerticalAlignment="Top" Foreground="White" FontSize="24"/>
        <Button Name="btnRandom" Content="Button" Opacity="0" HorizontalAlignment="Center" Margin="0,335,0,0" VerticalAlignment="Top" Height="26" Width="100" Background="#00DDDDDD" BorderBrush="#00707070" Foreground="#00000000"/>
    </Grid>
</Window>
"@

# Create the Window Object
$Reader = (New-Object System.Xml.XmlNodeReader $XAMLWindow)
$Window = [Windows.Markup.XamlReader]::Load( $Reader )

# Get Form Items

# Button
$btnG = $window.FindName("btnRGenre")
$btnI = $window.FindName("btnRItem")
$btnR = $window.FindName("btnRandom")

# List
$lstG = $window.FindName("lstGenre")
$lstI = $window.FindName("lstItem")

# Label
$lblT = $window.FindName("lblTitle")

#Text Block
$tbS = $Window.FindName("tbSelection")

[array]$Genre = @("American", "Japanese", "Chinese", "Italian", "Mediteranian", "Spannish", "Other Asian")
ForEach ($item in $Genre) {
    $lstG.Items.Add($item)
}
# Event handler for clicking on itmes in list. Populate variables
$lstG.Add_SelectionChanged( {
        $lstI.Items.Clear()
        $selG = $lstG.SelectedItem

        switch ($selG) {
            "American" {
                $lstI.Items.Add("Hot Dog/Burger")
                $lstI.Items.Add("Wings\Fried Chicken")
                $lstI.Items.Add("Sandwich")
                $lstI.Items.Add("Steak")
                $lstI.Items.Add("American BBQ")
            }
            "Japanese" { 
                $lstI.Items.Add("Noodles")
                $lstI.Items.Add("Ramen")
                $lstI.Items.Add("Sushi")
            }
            "Chinese" {
                $lstI.Items.Add{ Take-Out }
                $lstI.Items.Add{ Peking Duck }
                $lstI.Items.Add{ Dim Sum }
            }
            "Italian" {
                $lstI.Items.Add("Pizza")
                $lstI.Items.Add("Pasta")
            }
            "Mediteranian" {
                $lstI.Items.Add("Greek")
                $lstI.Items.Add("Seafood")
            }
            "Spannish" { 
                $lstI.Items.Add("Burrito\Taco")
                $lstI.Items.Add("Brazillian\Randizzio")
                $lstI.Items.Add("Cajun")
            }
            "Other Asian" { 
                $lstI.Items.Add("Pad Thai")
                $lstI.Items.Add("Pho")
                $lstI.Items.Add("Korean BBQ")
                $lstI.Items.Add("Indian")
            }
        }

        $lstI.Add_SelectionChanged({
                $tbS.Text = $lstI.Items[$lstI.SelectedIndex.ToString()]
            })

    })

    #Random Genre
$btnG.add_click({
  $ranVal = get-random -Maximum $lstG.Items.Count
  $lstG.SelectedIndex = $ranVal
})
    # Random Item
$btnI.add_click({
    $ranVal = get-random -Maximum $lstI.Items.Count
    $lstI.SelectedIndex = $ranVal
  })

$btnR.add_click({
    $ranVal = get-random -Maximum $lstG.Items.Count
    $lstG.SelectedIndex = $ranVal
    $ranVal = get-random -Maximum $lstI.Items.Count
    $lstI.SelectedIndex = $ranVal
})

$Window.ShowDialog() | Out-Null
param (
    [Parameter(Mandatory=$true)]
    [ValidateSet("Anime", "TV")]
    [string]$mediaType,

    [Parameter(Mandatory=$true)]
    [string]$showTitle,

    [int]$Seasons
)

# Determine the target path based on mediaType
if ($mediaType -eq "Anime") {
    $targetPath = Join-Path -Path "P:\AwesomeAnimeFolder\" -ChildPath $showTitle
} elseif ($mediaType -eq "TV") {
    $targetPath = Join-Path -Path "P:\AwesomeTVFolder\" -ChildPath $showTitle
} else {
    Write-Host "Invalid mediaType. Please specify 'Anime' or 'TV'."
    exit
}

# Check if the main folder for the show exists
if (Test-Path -Path $targetPath) {
    # Find the highest existing season number
    $existingSeasons = Get-ChildItem -Path $targetPath -Directory | Where-Object { $_.Name -match "^Season \d+$" } | ForEach-Object { [int]($_.Name -replace "Season ", "") }
    $maxSeason = if ($existingSeasons) { $existingSeasons | Measure-Object -Maximum | Select-Object -ExpandProperty Maximum } else { 0 }

    # Create subfolders for each season starting from the next season number up to the specified number
    if ($Seasons) {
        for ($i = $maxSeason + 1; $i -le $Seasons; $i++) {
            $seasonFolder = "Season $i"
            New-Item -ItemType Directory -Path (Join-Path -Path $targetPath -ChildPath $seasonFolder) -Force
        }
    } else {
        # Create the next season folder
        $nextSeason = $maxSeason + 1
        $seasonFolder = "Season $nextSeason"
        New-Item -ItemType Directory -Path (Join-Path -Path $targetPath -ChildPath $seasonFolder) -Force
    }
} else {
    # Create the main folder for the show
    New-Item -ItemType Directory -Path $targetPath -Force

    # Create subfolders for each season if Seasons parameter is provided
    if ($Seasons) {
        for ($i = 1; $i -le $Seasons; $i++) {
            $seasonFolder = "Season $i"
            New-Item -ItemType Directory -Path (Join-Path -Path $targetPath -ChildPath $seasonFolder) -Force
        }
    } else {
        # Create Season 1 if no Seasons parameter is provided
        $seasonFolder = "Season 1"
        New-Item -ItemType Directory -Path (Join-Path -Path $targetPath -ChildPath $seasonFolder) -Force
    }
}

Write-Host "Folders created successfully!"

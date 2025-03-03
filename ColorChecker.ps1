# Color Checker

# Get all ConsoleColor options
$Colors = [Enum]::GetValues([System.ConsoleColor])

# Print the header row with background color names
Write-Host -NoNewline "Foreground\Background"  # Header for the foreground row
foreach ($Background in $Colors) {
    $Host.UI.RawUI.BackgroundColor = $Background
    $Host.UI.RawUI.ForegroundColor = "White"
    Write-Host -NoNewline " $Background " -Width 15
}
Write-Host ""

# Reset colors for the table body
$Host.UI.RawUI.BackgroundColor = "Black"
$Host.UI.RawUI.ForegroundColor = "Gray"

# Print each row for foreground colors
foreach ($Foreground in $Colors) {
    # Print the foreground color as the row header
    $Host.UI.RawUI.ForegroundColor = $Foreground
    Write-Host -NoNewline "$Foreground".PadRight(20)  # Adjust padding for alignment

    # Print each cell in the row for the foreground/background combination
    foreach ($Background in $Colors) {
        $Host.UI.RawUI.ForegroundColor = $Foreground
        $Host.UI.RawUI.BackgroundColor = $Background
        Write-Host -NoNewline "  FG/BG  "  # Cell content
    }
    Write-Host ""  # Newline after each row
}

# Reset to default colors after the script finishes
$Host.UI.RawUI.ForegroundColor = "Gray"
$Host.UI.RawUI.BackgroundColor = "Black"


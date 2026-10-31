# Each profile is authored independently, rather than mutating the baseline.
function Get-NativeParagraphTopologyProfiles {
    $texts = @('First native paragraph', 'Second native paragraph', 'Third native paragraph', 'Final native paragraph')
    $profiles = @(
        @{ Name = 'Baseline'; Order = @(0, 1, 2, 3) },
        @{ Name = 'DropFirst'; Order = @(1, 2, 3) },
        @{ Name = 'DropMiddle'; Order = @(0, 2, 3) },
        @{ Name = 'DropFinal'; Order = @(0, 1, 2) },
        @{ Name = 'Reorder'; Order = @(3, 0, 2) },
        @{ Name = 'Duplicate'; Order = @(0, 1, 1, 2, 3) },
        @{ Name = 'AmbiguousLanguage'; Order = @(0, 0, 3); FrenchParagraph = 1 }
    )
    foreach ($profile in $profiles) {
        $profile.Text = ($profile.Order | ForEach-Object { $texts[$_] }) -join "`r"
        $profile
    }
}

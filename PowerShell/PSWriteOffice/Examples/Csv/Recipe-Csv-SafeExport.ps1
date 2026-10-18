$path = '.\Review-Queue.csv'
$rows = @(
    [pscustomobject]@{ Id = 1001; Owner = 'Platform'; Comment = '=HYPERLINK("https://example.org")' }
    [pscustomobject]@{ Id = 1002; Owner = 'Security'; Comment = 'Ready for review' }
)

$rows | Export-OfficeCsv -Path $path -FormulaInjectionPolicy Escape -UseQuotes AsNeeded
Import-OfficeCsv -Path $path -InferSchema | Select-Object Id, Owner, Comment

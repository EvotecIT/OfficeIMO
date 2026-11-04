# One line: a title slide, then table slides. Long tables continue on the next slide with the header repeated.
Get-Process | Sort-Object CPU -Descending | Select-Object -First 40 Name, Id, CPU, WorkingSet64 |
    Export-OfficePowerPoint -Path .\Top-Processes.pptx -Title 'Busiest processes'

# Add -Category and -Value for a chart slide after the table.
$sales = @(
    [pscustomobject]@{ Region = 'North'; Revenue = 420; Target = 400 }
    [pscustomobject]@{ Region = 'South'; Revenue = 365; Target = 390 }
    [pscustomobject]@{ Region = 'East'; Revenue = 510; Target = 450 }
    [pscustomobject]@{ Region = 'West'; Revenue = 298; Target = 320 }
)
$sales | Export-OfficePowerPoint -Path .\Regional-Sales.pptx -Title 'Regional sales' -Subtitle 'Q3 2026, USD thousands' `
    -Category Region -Value Revenue, Target -ChartTitle 'Revenue against target' -Theme Boardroom -Show

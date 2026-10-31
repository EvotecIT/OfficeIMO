param([string] $OutputDirectory = (Join-Path $PSScriptRoot '../Documents'))
Import-Module PSWriteOffice -ErrorAction Stop
$null = New-Item -Path $OutputDirectory -ItemType Directory -Force
$path = Join-Path $OutputDirectory 'PowerPoint-Charts.pptx'
$rows = @(
    [pscustomobject]@{ Month = 'Jan'; MonthNumber = 1; Sales = 10; Profit = 4 }
    [pscustomobject]@{ Month = 'Feb'; MonthNumber = 2; Sales = 14; Profit = 6 }
    [pscustomobject]@{ Month = 'Mar'; MonthNumber = 3; Sales = 18; Profit = 8 }
)
New-OfficePowerPoint -Path $path {
    PowerPointSlide -Title 'Sales and profit' {
        $rows | PowerPointChart -Type Column -Category Month -Value Sales, Profit -Title 'Sales vs profit'
    }
    PowerPointSlide -Title 'Sales mix' {
        PowerPointChart -Type Pie -InputObject $rows -Category Month -Value Sales -Title 'Sales mix'
    }
    PowerPointSlide -Title 'Numeric trends' {
        PowerPointChart -Type Scatter -InputObject $rows -Category MonthNumber -Value Sales, Profit -Title 'Trend scatter'
    }
}
Write-Host "Presentation saved to $path"

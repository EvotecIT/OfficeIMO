$path = '.\Quarterly-Business-Review.pptx'
$trend = @(
    [pscustomobject]@{ Quarter = 'Q1'; Revenue = 4.2; Margin = 31 }
    [pscustomobject]@{ Quarter = 'Q2'; Revenue = 4.8; Margin = 33 }
    [pscustomobject]@{ Quarter = 'Q3'; Revenue = 5.1; Margin = 35 }
    [pscustomobject]@{ Quarter = 'Q4'; Revenue = 5.7; Margin = 36 }
)
$priorities = @(
    [pscustomobject]@{ Priority = 'Customer onboarding'; Owner = 'Product'; Target = 'Reduce setup time by 25%' }
    [pscustomobject]@{ Priority = 'Renewal risk'; Owner = 'Sales'; Target = 'Review top 20 accounts' }
    [pscustomobject]@{ Priority = 'Delivery capacity'; Owner = 'Operations'; Target = 'Add two automation lanes' }
)
$revenueGrowth = [math]::Round(($trend[-1].Revenue / $trend[0].Revenue - 1) * 100, 1)
$marginChange = $trend[-1].Margin - $trend[0].Margin

New-OfficePowerPoint -Path $path {
    Set-OfficePowerPointSlideSize -Preset Screen16x9

    $cover = PowerPointSlide -Background '#0F172A' -NoFooter -NoSlideNumber -PassThru
    Add-OfficePowerPointTextBox -Slide $cover -Content { PowerPointText 'NORTHWIND | YEAR-END REVIEW' -Size 18 -Color '#5EEAD4' } -X 60 -Y 90 -Width 840 -Height 40
    Add-OfficePowerPointTextBox -Slide $cover -Content { PowerPointText 'Quarterly Business Review' -Size 40 -Color '#FFFFFF' -Bold } -X 60 -Y 150 -Width 840 -Height 80
    Add-OfficePowerPointTextBox -Slide $cover -Content { PowerPointText 'Performance, decisions, and next-quarter priorities' -Size 22 -Color '#CBD5E1' } -X 60 -Y 250 -Width 840 -Height 70
    PowerPointNotes -Slide $cover -Text 'Open with the outcome: growth continued and the team needs three decisions.'

    $performance = PowerPointSlide -Title 'Performance trend' -PassThru
    PowerPointChart -Slide $performance -InputObject $trend -Category Quarter -Value Revenue -Type Column -Title 'Revenue (USD millions)' -X 55 -Y 130 -Width 400 -Height 280
    PowerPointChart -Slide $performance -InputObject $trend -Category Quarter -Value Margin -Type Line -Title 'Operating margin (%)' -X 500 -Y 130 -Width 400 -Height 280
    Add-OfficePowerPointTextBox -Slide $performance -Content { PowerPointText "Revenue grew $revenueGrowth% from Q1; margin improved $marginChange percentage points." -Size 18 -Color '#0F172A' } -X 60 -Y 440 -Width 840 -Height 55
    PowerPointNotes -Slide $performance -Text 'Explain the margin improvement before discussing revenue.'

    PowerPointSlide -Title 'Next-quarter priorities' {
        PowerPointTable -InputObject $priorities
        PowerPointList 'Approve owners', 'Confirm targets', 'Review progress monthly'
        PowerPointNotes 'Close by assigning each decision to a named owner.'
    }
}

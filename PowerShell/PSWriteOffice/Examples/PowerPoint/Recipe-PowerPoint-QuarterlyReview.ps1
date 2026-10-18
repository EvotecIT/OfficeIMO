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

PptNew -Path $path {
    PptSlideSize -Preset Screen16x9

    PptSlide {
        PptBackground -Color '#0F172A'
        PptTextBox -Run @{ Text = 'NORTHWIND | YEAR-END REVIEW'; Color = '#5EEAD4'; FontSize = 18 } -X 60 -Y 90 -Width 840 -Height 40
        PptTextBox -Run @{ Text = 'Quarterly Business Review'; Color = '#FFFFFF'; FontSize = 40; Bold = $true } -X 60 -Y 150 -Width 840 -Height 80
        PptTextBox -Run @{ Text = 'Performance, decisions, and next-quarter priorities'; Color = '#CBD5E1'; FontSize = 22 } -X 60 -Y 250 -Width 840 -Height 70
        PptNotes -Text 'Open with the outcome: growth continued and the team needs three decisions.'
    }

    PptSlide {
        PptTitle -Title 'Performance trend'
        PptChart -Data $trend -CategoryProperty Quarter -SeriesProperty Revenue -Type ClusteredColumn -Title 'Revenue (USD millions)' -X 55 -Y 130 -Width 400 -Height 280
        PptChart -Data $trend -CategoryProperty Quarter -SeriesProperty Margin -Type Line -Title 'Operating margin (%)' -X 500 -Y 130 -Width 400 -Height 280
        PptTextBox -Run @{ Text = "Revenue grew $revenueGrowth% from Q1; margin improved $marginChange percentage points."; FontSize = 18; Color = '#0F172A' } -X 60 -Y 440 -Width 840 -Height 55
        PptNotes -Text 'Explain the margin improvement before discussing revenue.'
    }

    PptSlide {
        PptTitle -Title 'Next-quarter priorities'
        PptTable -Data $priorities -X 55 -Y 130 -Width 840 -Height 200
        PptBullets -Bullets 'Approve owners', 'Confirm targets', 'Review progress monthly' -X 80 -Y 365 -Width 800 -Height 130
        PptNotes -Text 'Close by assigning each decision to a named owner.'
    }
}

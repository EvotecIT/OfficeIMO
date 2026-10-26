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

# Layouts come from the content: a cover first, two charts side by side, a table slide.
New-OfficePowerPoint -Path $path -Theme Boardroom -Force {
    PptSlide -Eyebrow 'Northwind | Year-end review' -Title 'Quarterly business review' -Subtitle 'Performance, decisions and next-quarter priorities' {
        PptNotes 'Open with the outcome: growth continued and the team needs three decisions.'
    }

    PptSlide -Title 'Performance trend' -Subtitle "Revenue grew $revenueGrowth% from Q1; margin improved $marginChange percentage points." {
        PptChart -InputObject $trend -Category Quarter -Value Revenue -Type Column -Title 'Revenue (USD millions)'
        PptChart -InputObject $trend -Category Quarter -Value Margin -Type Line -Title 'Operating margin (%)'
        PptNotes 'Explain the margin improvement before discussing revenue.'
    }

    PptSlide -Title 'Next-quarter priorities' {
        PptTable $priorities
        PptList 'Approve owners', 'Confirm targets', 'Review progress monthly'
        PptNotes 'Close by assigning each decision to a named owner.'
    }
}

function Assert-VisioShowcaseVisualQualityProof {
    param(
        [Parameter(Mandatory = $true)]
        [string] $RelativePath,

        [Parameter(Mandatory = $true)]
        [string] $FullPath
    )

    $visualQualityText = Get-Content -LiteralPath $FullPath -Raw
    Assert-Condition -Condition ($visualQualityText -match '(?m)^quality\.isClean=(true|false)\r?$') -Message "Visual-quality proof '$RelativePath' is missing quality.isClean."
    Assert-Condition -Condition ($visualQualityText -match '(?m)^quality\.issueCount=\d+\r?$') -Message "Visual-quality proof '$RelativePath' is missing quality.issueCount."
    Assert-Condition -Condition ($visualQualityText -match '(?m)^quality\.warningCount=\d+\r?$') -Message "Visual-quality proof '$RelativePath' is missing quality.warningCount."
    Assert-Condition -Condition ($visualQualityText -match '(?m)^quality\.errorCount=\d+\r?$') -Message "Visual-quality proof '$RelativePath' is missing quality.errorCount."
}

function Get-VisioShowcaseVisualQualityProofSummary {
    param(
        [Parameter(Mandatory = $true)]
        [string] $RelativePath,

        [Parameter(Mandatory = $true)]
        [string] $FullPath
    )

    Assert-VisioShowcaseVisualQualityProof -RelativePath $RelativePath -FullPath $FullPath
    $values = @{}
    foreach ($line in (Get-Content -LiteralPath $FullPath)) {
        if ($line -match '^([^=]+)=(.*)$') {
            $values[$Matches[1]] = $Matches[2]
        }
    }

    $issueKinds = @()
    if ($values.ContainsKey('quality.issueKinds') -and -not [string]::IsNullOrWhiteSpace($values['quality.issueKinds'])) {
        $issueKinds = @($values['quality.issueKinds'].Split(',') | ForEach-Object { $_.Trim() } | Where-Object { $_ } | Sort-Object)
    }

    [pscustomobject]@{
        hasProof = $true
        isClean = ($values['quality.isClean'] -eq 'true')
        issueCount = [long] $values['quality.issueCount']
        errorCount = [long] $values['quality.errorCount']
        warningCount = [long] $values['quality.warningCount']
        informationCount = [long] $values['quality.informationCount']
        issueKinds = $issueKinds
    }
}

param([string] $TestRoot = (Join-Path ([IO.Path]::GetTempPath()) ('officeimo-showcase-provenance-' + [guid]::NewGuid().ToString('N'))), [string] $ModuleVersion)
$ErrorActionPreference = 'Stop'
$ownsTestRoot = -not $PSBoundParameters.ContainsKey('TestRoot')
. (Join-Path $PSScriptRoot 'ShowcaseEvidence.Helpers.ps1')
. (Join-Path $PSScriptRoot 'ShowcaseEvidence.Provenance.ps1')
if (Test-Path -LiteralPath $TestRoot) { throw 'Use a fresh task-owned test directory.' }
New-Item -ItemType Directory -Path $TestRoot | Out-Null
try {
New-Item -ItemType Directory -Path (Join-Path $TestRoot 'Website/scripts') | Out-Null
foreach ($name in 'Invoke-ShowcasePowerShellExample.ps1', 'ShowcaseEvidence.Helpers.ps1', 'ShowcaseEvidence.Provenance.ps1') {
    'producer fixture' | Set-Content -LiteralPath (Join-Path $TestRoot ('Website/scripts/' + $name))
}
'sample script' | Set-Content -LiteralPath (Join-Path $TestRoot 'example.ps1')
'sample input' | Set-Content -LiteralPath (Join-Path $TestRoot 'input.csv')
$card = [pscustomobject]@{ id = 'example'; source_language = 'powershell'; source_path = 'example.ps1'; source_url = '/downloads/showcase/example.ps1'; downloads = @([pscustomobject]@{ url = '/downloads/showcase/input.csv' }, [pscustomobject]@{ url = '/downloads/showcase/report.pdf' }); previews = @([pscustomobject]@{ url = '/downloads/showcase/preview.png' }) }
$catalog = [pscustomobject]@{ cards = @($card); artifacts = @([pscustomobject]@{ sourceRoot = 'repository'; source = 'input.csv'; destination = 'input.csv' }) }
$card.downloads += [pscustomobject]@{ url = '/downloads/showcase/report.docx' }
$catalog.artifacts += @([pscustomobject]@{ sourceRoot = 'documents'; source = 'report.pdf'; destination = 'report.pdf' }, [pscustomobject]@{ sourceRoot = 'documents'; source = 'report.docx'; destination = 'report.docx' })
$catalog.artifacts += [pscustomobject]@{ sourceRoot = 'documents'; source = 'preview.png'; destination = 'preview.png'; previewSource = 'report.pdf'; previewPage = 1 }
$artifacts = @(
    [pscustomobject]@{ path = '/downloads/showcase/example.ps1'; sha256 = (Get-FileHash -LiteralPath (Join-Path $TestRoot 'example.ps1')).Hash.ToLowerInvariant() }
    [pscustomobject]@{ path = '/downloads/showcase/input.csv'; sha256 = (Get-FileHash -LiteralPath (Join-Path $TestRoot 'input.csv')).Hash.ToLowerInvariant() }
    [pscustomobject]@{ path = '/downloads/showcase/report.pdf'; sha256 = 'output-hash' }
    [pscustomobject]@{ path = '/downloads/showcase/preview.png'; sha256 = 'preview-hash' }
    [pscustomobject]@{ path = '/downloads/showcase/report.docx'; sha256 = 'docx-hash' }
)
$receipt = [pscustomobject]@{ exampleId = 'example'; state = 'generated'; completedUtc = '2026-01-01T00:00:00Z'; sources = @(Get-ShowcaseSourceBindings $TestRoot $catalog $card); package = [pscustomobject]@{ version = '3.0.8' }; outputs = @([pscustomobject]@{ path = '/downloads/showcase/report.pdf'; sha256 = 'output-hash' }) }
$receipt.outputs += [pscustomobject]@{ path = '/downloads/showcase/report.docx'; sha256 = 'docx-hash' }
$receipt | Add-Member -NotePropertyName startedUtc -NotePropertyValue '2026-01-01T00:00:00Z'
$receipt | Add-Member -NotePropertyName host -NotePropertyValue ([pscustomobject]@{ powershell = '7.6.6'; os = 'fixture'; architecture = 'X64' })
$receipt.package | Add-Member -NotePropertyName name -NotePropertyValue 'PSWriteOffice'
$receipt.package | Add-Member -NotePropertyName files -NotePropertyValue @([pscustomobject]@{ path = 'PSWriteOffice.psd1'; sha256 = 'a' * 64 }, [pscustomobject]@{ path = 'PSWriteOffice.psm1'; sha256 = 'b' * 64 })
$receipt.package | Add-Member -NotePropertyName fileCount -NotePropertyValue 2
$receipt.outputs[0] | Add-Member -NotePropertyName diagnostics -NotePropertyValue ([pscustomobject]@{ pageCount = 1; fonts = @() })
$receipt | Add-Member -NotePropertyName conversions -NotePropertyValue @([pscustomobject]@{ outputPath = '/downloads/showcase/report.pdf'; outputSha256 = 'output-hash'; routeId = 'docx-pdf'; command = 'Export-OfficeDocumentPdf'; routeEvidence = 'fixture'; report = [pscustomobject]@{ Warnings = @() }; warnings = @() })
$receipt | Add-Member -NotePropertyName previews -NotePropertyValue @([pscustomobject]@{ path = '/downloads/showcase/preview.png'; rendering = [ordered]@{ tool = 'pdftocairo'; version = 'fixture'; executableSha256 = 'c' * 64; page = 1; inputSha256 = 'output-hash'; sha256 = 'preview-hash'; diagnostics = @() } })
$receipt | Add-Member -NotePropertyName evidenceProducer -NotePropertyValue (Get-ShowcaseEvidenceProducerBinding $TestRoot)
$receipt | Add-Member -NotePropertyName receiptSha256 -NotePropertyValue (Get-ShowcaseGenerationReceiptHash $receipt)
$arguments = @{ RepositoryRoot = $TestRoot; Catalog = $catalog; Artifacts = $artifacts; ModuleVersion = '3.0.8'; Generated = @{}; Previous = @() }
function Assert-Contract([bool] $Condition, [string] $Message) { if (-not $Condition) { throw $Message } }
$unknown = @(Complete-ShowcaseGenerationEvidence @arguments)
Assert-Contract ($unknown[0].state -ceq 'not-recorded') 'Existing downloads must not acquire invented producer evidence.'
$arguments.Generated = @{ example = $receipt }
$fresh = @(Complete-ShowcaseGenerationEvidence @arguments)
Assert-Contract ($fresh[0].state -ceq 'generated' -and $fresh[0].previewReview.state -ceq 'pending') 'Fresh output must await visual review.'
$blocked = $false
try { Complete-ShowcaseGenerationEvidence @arguments -ReviewedExampleId example | Out-Null } catch { $blocked = $_.Exception.Message -like '*separate manifest-only*' }
Assert-Contract $blocked 'A new generation must reject same-invocation review attestation.'
$blocked = $false
try { & (Join-Path $PSScriptRoot 'Build-ShowcaseEvidence.ps1') -ExampleId example -ReviewedExampleId example | Out-Null } catch { $blocked = $_.Exception.Message -like '*separate manifest-only*' }
Assert-Contract $blocked 'The public builder must reject combined generation/review before reading or writing artifacts.'
$arguments.Generated = @{}
$arguments.Previous = $fresh
$reviewed = @(Complete-ShowcaseGenerationEvidence @arguments -ReviewedExampleId example)
Assert-Contract ($reviewed[0].previewReview.state -ceq 'reviewed' -and $reviewed[0].previewReview.artifacts[0].sha256 -ceq 'preview-hash') 'Review must bind the actual preview hash.'
$arguments.Previous = $reviewed
$retained = @(Complete-ShowcaseGenerationEvidence @arguments)
Assert-Contract (($retained | ConvertTo-Json -Depth 12 -Compress) -ceq ($reviewed | ConvertTo-Json -Depth 12 -Compress)) 'Manifest-only refresh must retain matching producer and review evidence unchanged.'
$healthy = $reviewed[0] | ConvertTo-Json -Depth 30
$arguments.Previous = @($reviewed[0], $reviewed[0])
Assert-Contract ((@(Complete-ShowcaseGenerationEvidence @arguments))[0].state -ceq 'invalidated') 'Duplicate retained receipts must not silently select an attestation.'
$arguments.Previous = $reviewed
foreach ($mutation in 'missing-review', 'changed-review', 'duplicate-review', 'wrong-review-path', 'old-review-time') {
    $corrupt = $healthy | ConvertFrom-Json -Depth 30
    switch ($mutation) {
        'missing-review' { $corrupt.previewReview.artifacts = @() }
        'changed-review' { $corrupt.previewReview.artifacts[0].sha256 = 'changed' }
        'duplicate-review' { $corrupt.previewReview.artifacts = @($corrupt.previewReview.artifacts[0], $corrupt.previewReview.artifacts[0]) }
        'wrong-review-path' { $corrupt.previewReview.artifacts[0].path = '/downloads/showcase/unrelated.png' }
        'old-review-time' { $corrupt.previewReview.reviewedUtc = '2025-01-01T00:00:00Z' }
    }
    $arguments.Previous = @($corrupt)
    $result = @(Complete-ShowcaseGenerationEvidence @arguments)[0]
    Assert-Contract ($result.state -ceq 'generated' -and $result.previewReview.state -ceq 'pending') 'A damaged preview attestation must return to pending without inventing a new generation.'
}
foreach ($mutation in 'missing-files', 'partial-files', 'duplicate-files', 'invalid-file-hash', 'missing-conversions', 'missing-report', 'missing-warnings', 'missing-previews', 'missing-renderer', 'changed-render-input', 'missing-font-diagnostics', 'missing-host', 'changed-receipt-hash') {
    $corrupt = $healthy | ConvertFrom-Json -Depth 30
    switch ($mutation) {
        'missing-files' { $corrupt.package.files = @() }
        'partial-files' { $corrupt.package.files = @($corrupt.package.files[0]) }
        'duplicate-files' { $corrupt.package.files = @($corrupt.package.files[0], $corrupt.package.files[0]) }
        'invalid-file-hash' { $corrupt.package.files[0].sha256 = 'not-a-hash' }
        'missing-conversions' { $corrupt.conversions = @() }
        'missing-report' { $corrupt.conversions[0].report = $null }
        'missing-warnings' { $corrupt.conversions[0].PSObject.Properties.Remove('warnings') }
        'missing-previews' { $corrupt.previews = @() }
        'missing-renderer' { $corrupt.previews[0].rendering = $null }
        'changed-render-input' { $corrupt.previews[0].rendering.inputSha256 = 'changed' }
        'missing-font-diagnostics' { $corrupt.outputs[0].diagnostics = $null }
        'missing-host' { $corrupt.host = $null }
        'changed-receipt-hash' { $corrupt.receiptSha256 = 'd' * 64 }
    }
    $arguments.Previous = @($corrupt)
    Assert-Contract ((@(Complete-ShowcaseGenerationEvidence @arguments))[0].state -ceq 'invalidated') 'A truncated or changed producer claim must invalidate the receipt.'
}
$arguments.Previous = $reviewed
$completeOutputs = $receipt.outputs
foreach ($incompleteOutputs in @(@(), @($completeOutputs[0]), @($completeOutputs[0], $completeOutputs[0]), @($completeOutputs[0], [pscustomobject]@{ path = '/downloads/showcase/unrelated.pdf'; sha256 = 'output-hash' }))) {
    $receipt.outputs = $incompleteOutputs
    Assert-Contract ((@(Complete-ShowcaseGenerationEvidence @arguments))[0].state -ceq 'invalidated') 'Missing, duplicate or unrelated native outputs must invalidate the producer receipt.'
}
$receipt.outputs = $completeOutputs
$completeProducer = $receipt.evidenceProducer
$receipt.evidenceProducer = $null
Assert-Contract ((@(Complete-ShowcaseGenerationEvidence @arguments))[0].state -ceq 'invalidated') 'A missing producer binding must invalidate retained evidence.'
$receipt.evidenceProducer = $completeProducer
$completeDependencies = $completeProducer.dependencies
foreach ($incompleteDependencies in @(@(), @($completeDependencies[0]), @($completeDependencies[0], $completeDependencies[0]))) {
    $receipt.evidenceProducer.dependencies = $incompleteDependencies
    Assert-Contract ((@(Complete-ShowcaseGenerationEvidence @arguments))[0].state -ceq 'invalidated') 'Missing, partial or duplicate producer dependencies must invalidate retained evidence.'
}
$receipt.evidenceProducer.dependencies = $completeDependencies
foreach ($producerFile in @($completeProducer) + @($completeProducer.dependencies)) {
    $path = Join-Path $TestRoot $producerFile.path
    $original = [IO.File]::ReadAllBytes($path)
    'changed producer' | Set-Content -LiteralPath $path
    Assert-Contract ((@(Complete-ShowcaseGenerationEvidence @arguments))[0].state -ceq 'invalidated') 'Changing the worker or a loaded helper must invalidate its retained evidence.'
    [IO.File]::WriteAllBytes($path, $original)
}
$artifacts[3].sha256 = 'changed-preview'
$stale = @(Complete-ShowcaseGenerationEvidence @arguments)
Assert-Contract ($stale[0].state -ceq 'invalidated' -and $stale[0].previewReview.state -ceq 'pending') 'Changed preview must invalidate previous review.'
$artifacts[3].sha256 = 'preview-hash'
$artifacts[2].sha256 = 'changed-output'
Assert-Contract ((@(Complete-ShowcaseGenerationEvidence @arguments))[0].state -ceq 'invalidated') 'Changed output must invalidate producer evidence.'
$artifacts[2].sha256 = 'output-hash'
'changed input' | Set-Content -LiteralPath (Join-Path $TestRoot 'input.csv')
Assert-Contract ((@(Complete-ShowcaseGenerationEvidence @arguments))[0].state -ceq 'invalidated') 'Changed input must invalidate producer evidence.'
$blocked = $false
try { Complete-ShowcaseGenerationEvidence @arguments -ReviewedExampleId example | Out-Null } catch { $blocked = $_.Exception.Message -like '*missing or stale*' }
Assert-Contract $blocked 'Stale generation must reject visual-review attestation.'
'sample input' | Set-Content -LiteralPath (Join-Path $TestRoot 'input.csv')
$arguments.ModuleVersion = '3.0.9'
Assert-Contract ((@(Complete-ShowcaseGenerationEvidence @arguments))[0].state -ceq 'invalidated') 'A changed documented module version must invalidate old producer evidence.'
if ($ModuleVersion) {
    # Exercise the real child runner against an available module. A successful recipe
    # that omits an export must never relabel an old destination file as fresh output.
    $runnerRoot = Join-Path $TestRoot 'runner'
    $runnerDocuments = Join-Path $runnerRoot 'documents'
    New-Item -ItemType Directory -Path (Join-Path $runnerRoot 'Website/data'), (Join-Path $runnerDocuments 'PowerShell') -Force | Out-Null
    @'
param([string] $OutputDirectory, [string] $ModuleVersion)
New-Item -ItemType Directory -Path $OutputDirectory -Force | Out-Null
'@ | Set-Content -LiteralPath (Join-Path $runnerRoot 'example.ps1')
    $runnerCard = [ordered]@{ id = 'example'; source_language = 'powershell'; source_path = 'example.ps1'; source_url = '/downloads/showcase/example.ps1'; downloads = @([ordered]@{ url = '/downloads/showcase/report.docx' }); previews = @() }
    [ordered]@{ cards = @($runnerCard); artifacts = @([ordered]@{ sourceRoot = 'documents'; source = 'PowerShell/report.docx'; destination = 'report.docx' }) } |
        ConvertTo-Json -Depth 8 | Set-Content -LiteralPath (Join-Path $runnerRoot 'Website/data/showcase.json')
    $oldOutput = Join-Path $runnerDocuments 'PowerShell/report.docx'
    'old-output' | Set-Content -LiteralPath $oldOutput
    $receiptPath = Join-Path $runnerRoot 'receipt.json'
    $logPath = Join-Path $runnerRoot 'runner.log'
    & pwsh -NoProfile -File (Join-Path $PSScriptRoot 'Invoke-ShowcasePowerShellExample.ps1') -RepositoryRoot $runnerRoot -DocumentsRoot $runnerDocuments -ExampleId example -ModuleVersion $ModuleVersion -ReceiptPath $receiptPath *> $logPath
    Assert-Contract ($LASTEXITCODE -ne 0 -and (Get-Content -LiteralPath $logPath -Raw).Contains('Current showcase generation did not produce')) 'The runner must reject the omitted export, even with an older destination file.'
    Assert-Contract (-not (Test-Path -LiteralPath $receiptPath)) 'A failed generation must not produce a success receipt.'
    Assert-Contract ((Get-Content -LiteralPath $oldOutput -Raw).Trim() -ceq 'old-output') 'An incomplete generation must preserve the previous destination.'
    Assert-Contract (@(Get-ChildItem -LiteralPath $runnerDocuments -Directory -Filter 'ShowcaseGeneration-*').Count -eq 0) 'The runner must remove its failed temporary generation.'
}
Write-Host 'Showcase generation provenance contracts passed.'
} finally {
    if ($ownsTestRoot) { Remove-Item -LiteralPath $TestRoot -Recurse -ErrorAction Stop }
}

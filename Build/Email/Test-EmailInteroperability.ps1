<#
.SYNOPSIS
Runs selected email qualification lanes and rejects missing or skipped evidence.
.DESCRIPTION
Managed runs the ordinary email tests. External lanes use the existing explicit
test prerequisites; this script does not enable Outlook, choose a profile, or
download corpora. Results and potentially sensitive test diagnostics stay in the
caller-selected output directory. A requested lane must execute every named test.
#>
[CmdletBinding()]
param(
    [ValidateSet('Managed', 'MsgReader', 'MimeKit', 'LibPff', 'Outlook', 'Smime',
        'PrivateStores', 'PrivateStoreConversion', 'ApplePartial', 'NativeKeychain')]
    [string[]] $Lane = @('Managed'),
    [ValidateSet('net8.0', 'net10.0', 'net472')]
    [string] $Framework = 'net8.0',
    [string] $Configuration = 'Release',
    [string] $OutputPath = '',
    [string] $ArtifactsPath,
    [switch] $NoRestore,
    [switch] $NoBuild
)

$ErrorActionPreference = 'Stop'
# Inspect native exit codes explicitly so the report also records failed lanes.
$PSNativeCommandUseErrorActionPreference = $false
$repoRoot = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '../..'))
$project = Join-Path $repoRoot 'OfficeIMO.Email.Tests/OfficeIMO.Email.Tests.csproj'
if ($Lane.Count -eq 0) { throw 'Request at least one qualification lane.' }
if ([string]::IsNullOrWhiteSpace($OutputPath)) {
    $OutputPath = Join-Path ([IO.Path]::GetTempPath()) ('officeimo-email-qualification-' + [Guid]::NewGuid().ToString('N'))
}
$outputProvider = $null
$outputDrive = $null
$output = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath(
    $OutputPath, [ref] $outputProvider, [ref] $outputDrive)
if ($outputProvider.Name -ne 'FileSystem') { throw 'OutputPath must use the filesystem provider.' }
if (Test-Path -LiteralPath $output) {
    if (-not (Test-Path -LiteralPath $output -PathType Container) -or
        @(Get-ChildItem -LiteralPath $output -Force).Count -ne 0) {
        throw 'OutputPath must be new or empty; previous evidence is never overwritten.'
    }
} else {
    [void] [IO.Directory]::CreateDirectory($output)
}

$tests = 'OfficeIMO.Email.Tests.'
$stores = 'OfficeIMO.Email.Store.Tests.'
$lanes = [ordered] @{
    Managed = @()
    ApplePartial = @($stores + 'ExternalApplePartialEmlxCorpusTests.RecoversPinnedIndependentSiblingAndPreservesDecodedBytesThroughEmlExport')
    NativeKeychain = @($tests + 'NativeMacKeychainTests.CallerSelectedNonExtractableIdentitySignsDecryptsAndHonorsKeychainAuthorization')
    MsgReader = @($tests + 'ExternalEmailCorpusTests.ProcessesAllMsgReaderSamplesWhenCorpusIsAvailable')
    MimeKit = @(
        $tests + 'ExternalEmailCorpusTests.ProcessesMimeKitMimeTnefAndMboxCorporaWhenAvailable'
        $tests + 'MimeKitAdversarialInteropTests.PinnedProducerCorpus_covers_the_adversarial_scenario_matrix'
        $tests + 'MimeKitAdversarialInteropTests.Producer_fixture_preserves_duplicate_header_order'
    )
    LibPff = @($stores + 'LibPffPstWriterInteropTests.GeneratedUnicodePstCanBeInspectedAndSemanticallyExportedByLibPff')
    Outlook = @(
        $tests + 'OutlookInteropTests.ExchangesMailAppointmentContactAndTaskMsgFilesWithInstalledOutlookWhenEnabled'
        $stores + 'OutlookPstWriterInteropTests.Generated_unicode_pst_passes_scanpst_without_repair_or_byte_changes'
        $stores + 'OutlookPstWriterInteropTests.Generated_unicode_pst_can_be_mounted_read_and_removed_by_classic_outlook'
    )
    Smime = @($tests + 'ExternalOutlookSmimeCorpusTests.VerifiesAndDecryptsRealOutlookEmlAndMsgCorpusWhenAvailable')
    PrivateStores = @($stores + 'ExternalEmailStoreCorpusTests.ReadsBoundedPrivatePstAndOstCorpusWithoutTableTraversalFailures')
    PrivateStoreConversion = @($stores + 'ExternalEmailStoreCorpusTests.ConvertsBoundedPrivateStoresToTemporaryPstWithoutRetainingCorpusContent')
}
$externalMethods = @($lanes.GetEnumerator() | Where-Object Key -ne 'Managed' |
    ForEach-Object { $_.Value })
$managedFilter = 'Category!=Performance&' + (($externalMethods | ForEach-Object { 'FullyQualifiedName!=' + $_ }) -join '&')
$started = [DateTime]::UtcNow
$head = & git -C $repoRoot rev-parse HEAD
if ($LASTEXITCODE -ne 0) { throw 'Cannot record the source commit.' }
$dirty = & git -C $repoRoot status --porcelain
if ($LASTEXITCODE -ne 0) { throw 'Cannot record source modification status.' }
$sdk = & dotnet --version
if ($LASTEXITCODE -ne 0) { throw 'The dotnet SDK is unavailable.' }
$results = [Collections.Generic.List[object]]::new()
$failed = $false

foreach ($entry in $lanes.GetEnumerator()) {
    $name = [string] $entry.Key
    if ($Lane -notcontains $name) {
        $results.Add([ordered] @{
            lane = $name; status = 'disabled'; reason = 'Not requested.'
        })
        continue
    }
    $filter = if ($name -eq 'Managed') { $managedFilter } else {
        ($entry.Value | ForEach-Object { 'FullyQualifiedName=' + $_ }) -join '|'
    }
    $row = [ordered] @{
        lane = $name; status = 'failed'; filter = $filter
        requiredTests = @($entry.Value); exitCode = $null
        total = 0; passed = 0; failed = 0; skipped = 0
        trx = $name + '/' + $name + '.trx'; reason = $null
    }
    $laneOutput = Join-Path $output $name
    [void] [IO.Directory]::CreateDirectory($laneOutput)
    $arguments = @('test', $project, '--configuration', $Configuration,
        '--framework', $Framework, '--filter', $filter, '--nologo', '-v', 'quiet',
        '--logger', ('trx;LogFileName=' + $name + '.trx'), '--results-directory', $laneOutput)
    if ($NoRestore) { $arguments += '--no-restore' }
    if ($NoBuild) { $arguments += '--no-build' }
    if ($ArtifactsPath) { $arguments += @('--artifacts-path', $ArtifactsPath) }
    Write-Host "Email qualification: $name ($Framework)"
    try {
        # Use the repository cwd for restore/configuration and relative test assets.
        Push-Location $repoRoot
        try {
            & dotnet @arguments
            $row.exitCode = $LASTEXITCODE
        } finally {
            Pop-Location
        }
        $trxPath = Join-Path $output $row.trx
        if (-not (Test-Path -LiteralPath $trxPath -PathType Leaf)) {
            throw 'The test host did not produce a TRX result.'
        }
        # Evidence is local test-host output; refuse DTD/external entity processing.
        $settings = [Xml.XmlReaderSettings]::new()
        $settings.DtdProcessing = [Xml.DtdProcessing]::Prohibit
        $settings.XmlResolver = $null
        $reader = [Xml.XmlReader]::Create($trxPath, $settings)
        try {
            $trx = [Xml.XmlDocument]::new()
            $trx.XmlResolver = $null
            $trx.Load($reader)
        } finally { $reader.Dispose() }
        $counters = $trx.TestRun.ResultSummary.Counters
        if ($null -eq $counters) { throw 'TRX result counters are missing.' }
        $row.total = [int] $counters.total
        $row.passed = [int] $counters.passed
        $row.failed = [int] $counters.failed
        $records = @($trx.TestRun.Results.UnitTestResult)
        # VSTest can report notExecuted=0 while individual records are skipped.
        $row.skipped = [Math]::Max([int] $counters.notExecuted,
            @($records | Where-Object outcome -eq 'NotExecuted').Count)
        if ($row.exitCode -ne 0 -or $row.total -le 0 -or
            $row.passed -ne $row.total -or $row.failed -ne 0 -or
            $row.skipped -ne 0 -or $records.Count -ne $row.total -or
            @($records | Where-Object outcome -ne 'Passed').Count -ne 0) {
            throw 'Requested evidence must contain executed, passing tests with no skips or failures.'
        }
        foreach ($required in $entry.Value) {
            $matched = @($records | Where-Object {
                $_.testName -eq $required -or
                ([string] $_.testName).StartsWith($required + '(', [StringComparison]::Ordinal)
            })
            if ($matched.Count -eq 0) { throw "Required test was not executed: $required" }
        }
        $row.status = 'passed'
    } catch {
        $row.reason = $_.Exception.Message
        $failed = $true
        Write-Warning "Email qualification failed: $name. See its local TRX and report."
    }
    $results.Add($row)
}

$report = [ordered] @{
    schemaVersion = 1; sourceCommit = ([string] $head).Trim()
    sourceModified = (@($dirty).Count -gt 0)
    framework = $Framework; configuration = $Configuration; sdk = ([string] $sdk).Trim()
    startedUtc = $started.ToString('o'); completedUtc = [DateTime]::UtcNow.ToString('o')
    status = $(if ($failed) { 'failed' } else { 'passed' })
    lanes = @($results.ToArray())
}
$json = $report | ConvertTo-Json -Depth 8
[IO.File]::WriteAllText((Join-Path $output 'qualification.json'), $json,
    [Text.UTF8Encoding]::new($false))
Write-Host "Email qualification report: $(Join-Path $output 'qualification.json')"
if ($failed) { throw 'One or more requested email qualification lanes failed.' }

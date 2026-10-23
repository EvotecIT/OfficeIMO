$ErrorActionPreference = 'Stop'
$entrypoint = Join-Path $PSScriptRoot 'Get-NpmReleaseEligibility.ps1'
$root = if ($env:RUNNER_TEMP) { $env:RUNNER_TEMP } elseif ($env:EVOTEC_SCRATCH_ROOT) { $env:EVOTEC_SCRATCH_ROOT } else { [IO.Path]::GetTempPath() }
$fixture = Join-Path $root ('officeimo-npm-eligibility-' + [Guid]::NewGuid().ToString('N'))
$null = New-Item -ItemType Directory -Path (Join-Path $fixture 'OfficeIMO.JavaScript/src') -Force
$names = @('GITHUB_EVENT_NAME', 'SOURCE_SHA', 'REQUESTED_QUALITY_RUN', 'REQUESTED_PUBLISH', 'REQUESTED_BOOTSTRAP', 'GITHUB_STEP_SUMMARY', 'GITHUB_OUTPUT')
$global:OfficeNpmEligibilityFixtureState = [pscustomobject]@{ RegistryStatus = 0; RegistryBody = ''; QualityResponse = '' }
$saved = @{}
foreach ($name in $names) { $saved[$name] = [Environment]::GetEnvironmentVariable($name, 'Process') }

function Invoke-FixtureGit {
    param([string[]] $Arguments)
    $value = & git -C $fixture @Arguments 2>&1
    if ($LASTEXITCODE -ne 0) { throw "Fixture Git failed: $value" }
    return $value
}

# Only the registry and workflow lookup boundaries are substituted; Git history
# and the production eligibility/path-gate scripts execute normally.
function Invoke-WebRequest { param($Uri, [switch] $SkipHttpErrorCheck)
    return [pscustomobject]@{ StatusCode = $global:OfficeNpmEligibilityFixtureState.RegistryStatus; Content = $global:OfficeNpmEligibilityFixtureState.RegistryBody }
}
function gh {
    $global:LASTEXITCODE = 0
    return $global:OfficeNpmEligibilityFixtureState.QualityResponse
}
function Invoke-Case {
    param([string] $Name, [string] $Event = 'workflow_run', [int] $Status = 200,
        [bool] $Publish = $false, [bool] $Bootstrap = $false, [bool] $Eligible = $true,
        [string] $ErrorPattern, [string] $Source = $script:JavaScriptCommit,
        [string] $PackageName = '@evotecit/officeimo', [string] $Quality = '123')
    $global:OfficeNpmEligibilityFixtureState.RegistryStatus = $Status
    $global:OfficeNpmEligibilityFixtureState.RegistryBody = @{ name = $PackageName; versions = @{ '0.1.0' = @{} } } | ConvertTo-Json -Depth 4
    $env:GITHUB_EVENT_NAME = $Event
    $env:SOURCE_SHA = $Source
    $env:REQUESTED_QUALITY_RUN = $Quality
    $env:REQUESTED_PUBLISH = $Publish.ToString().ToLowerInvariant()
    $env:REQUESTED_BOOTSTRAP = $Bootstrap.ToString().ToLowerInvariant()
    $env:GITHUB_OUTPUT = Join-Path $fixture 'output.txt'
    $env:GITHUB_STEP_SUMMARY = Join-Path $fixture 'summary.txt'
    [IO.File]::WriteAllText($env:GITHUB_OUTPUT, '')
    try {
        & $entrypoint
        if ($ErrorPattern) { throw "Expected rejection: $Name" }
        $output = Get-Content -LiteralPath $env:GITHUB_OUTPUT -Raw
        if ($output -notmatch "(?m)^eligible=$($Eligible.ToString().ToLowerInvariant())\r?$" -or
            $output -notmatch "(?m)^source-sha=$Source\r?$" -or $output -notmatch '(?m)^quality-run-id=123\r?$') {
            throw "Unexpected publication eligibility: $Name"
        }
    } catch {
        if (-not $ErrorPattern -or $_.Exception.Message -notmatch $ErrorPattern) { throw }
    }
    $script:Passed++
    Write-Output "PASS: $Name"
}

Push-Location $fixture
try {
    $null = Invoke-FixtureGit -Arguments @('init', '-q')
    $null = Invoke-FixtureGit -Arguments @('config', 'user.name', 'Synthetic fixture')
    $null = Invoke-FixtureGit -Arguments @('config', 'user.email', 'fixture@example.invalid')
    $null = Invoke-FixtureGit -Arguments @('config', 'commit.gpgsign', 'false')
    $manifest = Join-Path $fixture 'OfficeIMO.JavaScript/package.json'
    [IO.File]::WriteAllText($manifest, '{"name":"@evotecit/officeimo","version":"0.1.0"}')
    [IO.File]::WriteAllText((Join-Path $fixture 'OfficeIMO.JavaScript/src/index.ts'), 'export const value = 1;')
    $null = Invoke-FixtureGit -Arguments @('add', '.')
    $null = Invoke-FixtureGit -Arguments @('commit', '-qm', 'Published package')
    $null = Invoke-FixtureGit -Arguments @('tag', 'npm-v0.1.0')
    [IO.File]::WriteAllText((Join-Path $fixture 'engine.cs'), '// unrelated engine change')
    $null = Invoke-FixtureGit -Arguments @('add', '.')
    $null = Invoke-FixtureGit -Arguments @('commit', '-qm', 'Native only')
    $nativeCommit = [string](Invoke-FixtureGit -Arguments @('rev-parse', 'HEAD'))
    [IO.File]::WriteAllText((Join-Path $fixture 'OfficeIMO.JavaScript/src/index.ts'), 'export const value = 2;')
    $null = Invoke-FixtureGit -Arguments @('add', '.')
    $null = Invoke-FixtureGit -Arguments @('commit', '-qm', 'JavaScript update')
    $script:JavaScriptCommit = [string](Invoke-FixtureGit -Arguments @('rev-parse', 'HEAD'))
    $global:OfficeNpmEligibilityFixtureState.QualityResponse = '[{"databaseId":123}]'
    $script:Passed = 0
    Invoke-Case -Name 'First automatic run skips before environment publication' -Status 404 -Eligible $false
    Invoke-Case -Name 'First manual dry run qualifies without bootstrap' -Event workflow_dispatch -Status 404
    Invoke-Case -Name 'First manual upload requires bootstrap' -Event workflow_dispatch -Status 404 -Publish $true -ErrorPattern 'first publication requires bootstrap'
    Invoke-Case -Name 'Explicit first bootstrap qualifies' -Event workflow_dispatch -Status 404 -Publish $true -Bootstrap $true
    Invoke-Case -Name 'Bootstrap cannot target an existing package' -Event workflow_dispatch -Publish $true -Bootstrap $true -ErrorPattern 'only for a package that does not exist'
    Invoke-Case -Name 'Bootstrap cannot be a dry run' -Event workflow_dispatch -Bootstrap $true -ErrorPattern 'explicit manual publication'
    Invoke-Case -Name 'Registry outage stops selection' -Status 503 -ErrorPattern 'Cannot determine'
    Invoke-Case -Name 'Wrong registry package identity stops selection' -PackageName '@example/wrong' -ErrorPattern 'Invalid npm package metadata'
    Invoke-Case -Name 'Native-only source skips automatic publication' -Source $nativeCommit -Eligible $false
    Invoke-Case -Name 'JavaScript source qualifies after an intervening native merge'
    Invoke-Case -Name 'Manual dispatch locates successful source quality' -Event workflow_dispatch -Quality ''
    $global:OfficeNpmEligibilityFixtureState.QualityResponse = '[]'
    Invoke-Case -Name 'Missing successful quality stops manual dispatch' -Event workflow_dispatch -Quality '' -ErrorPattern 'No successful master push'
    Invoke-Case -Name 'Malformed source identity stops selection' -Source 'not-a-commit' -ErrorPattern 'Invalid npm source commit'
    [IO.File]::WriteAllText($manifest, '{"name":"@evotecit/officeimo","version":"0.2.0"}')
    Invoke-Case -Name 'Automatic minor-line change requires manual dispatch' -ErrorPattern 'major or minor npm line requires a manual'
    Invoke-Case -Name 'Manual minor-line dispatch remains available' -Event workflow_dispatch
    Write-Output "$script:Passed npm release eligibility contracts passed. Fixture: $fixture"
} finally {
    Pop-Location
    Remove-Variable -Name OfficeNpmEligibilityFixtureState -Scope Global -ErrorAction SilentlyContinue
    foreach ($name in $names) { [Environment]::SetEnvironmentVariable($name, $saved[$name], 'Process') }
}

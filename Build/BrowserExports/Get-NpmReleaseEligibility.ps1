$ErrorActionPreference = 'Stop'
$automatic = $env:GITHUB_EVENT_NAME -ceq 'workflow_run'
if ($env:SOURCE_SHA -notmatch '^[a-fA-F0-9]{40}$') { throw 'Invalid npm source commit.' }
if ($env:REQUESTED_BOOTSTRAP -ceq 'true' -and $env:REQUESTED_PUBLISH -cne 'true') {
  throw 'Bootstrap requires an explicit manual publication.'
}
$quality = $env:REQUESTED_QUALITY_RUN
if (-not $quality) {
  $json = & gh run list --workflow dotnet-tests.yml --commit $env:SOURCE_SHA --branch master --event push --status success --limit 100 --json databaseId
  if ($LASTEXITCODE -ne 0) { throw 'Cannot read npm source quality checks.' }
  $runs = @($json | ConvertFrom-Json)
  if (-not $runs.Count) { throw 'No successful master push quality run qualifies this source.' }
  $quality = [string]$runs[0].databaseId
}
$manifest = Get-Content OfficeIMO.JavaScript/package.json -Raw | ConvertFrom-Json
$response = Invoke-WebRequest ('https://registry.npmjs.org/' + [Uri]::EscapeDataString($manifest.name)) -SkipHttpErrorCheck
$eligible = $true
if ($response.StatusCode -eq 404) {
  if ($automatic) {
    $eligible = $false
    'First npm publication requires the owner-created npm-release environment and an explicit bootstrap dispatch. No version was reserved.' >> $env:GITHUB_STEP_SUMMARY
  } elseif ($env:REQUESTED_PUBLISH -ceq 'true' -and $env:REQUESTED_BOOTSTRAP -cne 'true') {
    throw 'The first publication requires bootstrap=true.'
  }
} elseif ($response.StatusCode -eq 200) {
  if ($env:REQUESTED_BOOTSTRAP -ceq 'true') { throw 'Bootstrap is only for a package that does not exist.' }
  $published = $response.Content | ConvertFrom-Json
  if ($published.name -cne $manifest.name -or $null -eq $published.versions) { throw 'Invalid npm package metadata.' }
  $latest = @($published.versions.PSObject.Properties.Name | Where-Object {
    $_ -cmatch '^(0|[1-9]\d*)\.(0|[1-9]\d*)\.(0|[1-9]\d*)$'
  } | ForEach-Object { [Version]::Parse($_) } | Sort-Object -Descending) | Select-Object -First 1
  if (-not $latest) { throw 'No stable published npm source is available for comparison.' }
  if ($automatic) {
    $minimum = [Version]::Parse($manifest.version)
    if ($minimum.Major -gt $latest.Major -or ($minimum.Major -eq $latest.Major -and $minimum.Minor -gt $latest.Minor)) {
      throw 'A new major or minor npm line requires a manual version dispatch.'
    }
    $changes = & "$PSScriptRoot/Get-NpmPackageChanges.ps1" -RepositoryRoot (Get-Location).Path -BaseCommit "refs/tags/npm-v$latest" -SourceCommit $env:SOURCE_SHA
    $eligible = $changes.Changed
    if (-not $eligible) { 'No npm package or runtime input changed since the last public npm version; skipping publication.' >> $env:GITHUB_STEP_SUMMARY }
  }
} else { throw 'Cannot determine the public npm package state.' }
"eligible=$($eligible.ToString().ToLowerInvariant())" >> $env:GITHUB_OUTPUT
"source-sha=$env:SOURCE_SHA" >> $env:GITHUB_OUTPUT
"quality-run-id=$quality" >> $env:GITHUB_OUTPUT

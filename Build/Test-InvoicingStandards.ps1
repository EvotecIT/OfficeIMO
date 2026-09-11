[CmdletBinding()]
param(
    [string] $OutputPath = (Join-Path $PSScriptRoot '../artifacts/invoicing-standards'),
    [string] $Framework = 'net10.0',
    [string] $JavaExecutable = 'java'
)
$ErrorActionPreference = 'Stop'
$repository = Split-Path $PSScriptRoot -Parent
$OutputPath = [IO.Path]::GetFullPath($OutputPath)
if (!(Test-Path -LiteralPath $OutputPath)) { $null = New-Item -ItemType Directory -Path $OutputPath }
# Read download identities from the owning validator rather than maintaining duplicate pins.
$bundleSource = Get-Content (Join-Path $repository 'OfficeIMO.Invoicing.Validation/InvoiceRuleBundle.cs') -Raw
$runnerSource = Get-Content (Join-Path $repository 'OfficeIMO.Invoicing.Validation/SaxonInvoiceRulesRunner.cs') -Raw
function Get-Constant([string] $Source, [string] $Name) {
    $match = [regex]::Match($Source, 'public const string ' + [regex]::Escape($Name) + ' = "([^"]+)";')
    if (!$match.Success) { throw "Missing validator constant: $Name" }
    $match.Groups[1].Value
}
function Get-PinnedArtifact([string] $Url, [string] $Hash, [string] $Name) {
    $target = Join-Path $OutputPath $Name
    if (!(Test-Path -LiteralPath $target)) { Invoke-WebRequest -Uri $Url -OutFile $target -TimeoutSec 120 }
    if ((Get-FileHash -LiteralPath $target -Algorithm SHA256).Hash -ne $Hash) { throw "Artifact hash mismatch: $target" }
    $target
}
$bundle = Get-PinnedArtifact (Get-Constant $bundleSource 'XRechnungDownloadUrl') (Get-Constant $bundleSource 'XRechnungSha256') 'xrechnung.zip'
$peppol = Get-PinnedArtifact (Get-Constant $bundleSource 'PeppolDownloadUrl') (Get-Constant $bundleSource 'PeppolSha256') 'peppol.sch'
$saxon = Get-PinnedArtifact (Get-Constant $runnerSource 'DownloadUrl') (Get-Constant $runnerSource 'ArchiveSha256') 'saxon.zip'
$saxonPath = Join-Path $OutputPath 'saxon'
if (!(Test-Path -LiteralPath $saxonPath)) { Expand-Archive -LiteralPath $saxon -DestinationPath $saxonPath }
$variables = @{
    OFFICEIMO_INVOICE_STANDARDS_TESTS = '1'
    OFFICEIMO_INVOICE_RULE_BUNDLE = $bundle
    OFFICEIMO_INVOICE_PEPPOL_RULES = $peppol
    OFFICEIMO_INVOICE_SAXON_JAR = (Join-Path $saxonPath 'saxon-he-12.10.jar')
    OFFICEIMO_INVOICE_JAVA = $JavaExecutable
    OFFICEIMO_INVOICE_EVIDENCE = (Join-Path $OutputPath 'evidence')
}
$previous = @{}
try {
    foreach ($name in $variables.Keys) {
        $previous[$name] = [Environment]::GetEnvironmentVariable($name, 'Process')
        [Environment]::SetEnvironmentVariable($name, $variables[$name], 'Process')
    }
    dotnet test (Join-Path $repository 'OfficeIMO.Invoicing.Validation.Tests/OfficeIMO.Invoicing.Validation.Tests.csproj') -c Release -f $Framework --logger 'console;verbosity=normal'
    if ($LASTEXITCODE -ne 0) { throw "Invoice standards tests failed ($LASTEXITCODE)." }
} finally {
    foreach ($name in $previous.Keys) { [Environment]::SetEnvironmentVariable($name, $previous[$name], 'Process') }
}

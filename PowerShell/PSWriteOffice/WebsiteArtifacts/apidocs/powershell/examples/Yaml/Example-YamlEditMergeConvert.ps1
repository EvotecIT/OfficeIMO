Import-Module PSWriteOffice -ErrorAction Stop
$documents = Join-Path $PSScriptRoot '..\Documents'
$null = New-Item -Path $documents -ItemType Directory -Force
$values = Join-Path $documents 'Yaml-Values.yaml'
$production = Join-Path $documents 'Yaml-Values.prod.yaml'
$merged = Join-Path $documents 'Yaml-Values.merged.yaml'

# A chart's defaults, with comments and quoting a hand-maintained file would have.
@'
# Defaults for the web chart
image:
  repository: nginx   # pinned by CI
  tag: '1.25'
replicas: 1

# Ports the service exposes
ports: [80, 443]
'@ | Set-Content -LiteralPath $values -NoNewline

# Edit in place: only the tag's text changes when the block saves.
Open-OfficeYaml -Path $values {
    Set-OfficeYamlNode -Select 'image.tag' -Value '1.26'
    Set-OfficeYamlNode -Select 'resources.limits' -Value ([ordered]@{ cpu = '500m'; memory = '256Mi' })
}
Get-Content -LiteralPath $values -Raw

# Layer production values over the defaults, the way Helm combines values files.
[ordered]@{ image = [ordered]@{ tag = '1.27' }; replicas = 3 } | Export-OfficeYaml -Path $production
Merge-OfficeDocument -Path $values, $production -OutputPath $merged -PassThruReport |
    Select-Object -Property Added, Replaced, Removed

# Read the result as data and check it loads under the strict rules.
$settings = Import-OfficeYaml -Path $merged
'{0}:{1} x{2}' -f $settings.image.repository, $settings.image.tag, $settings.replicas
"Loads under the strict rules: $(Test-OfficeYaml -Path $merged -Strict)"

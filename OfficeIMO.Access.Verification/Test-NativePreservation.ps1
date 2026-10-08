param([Parameter(Mandatory)][string] $OutputDirectory)
$ErrorActionPreference = 'Stop'
if ($env:OS -ne 'Windows_NT') { throw 'The independent application oracle requires Windows Access.' }
$root = [IO.Path]::GetFullPath($OutputDirectory)
$records = Get-Content -LiteralPath (Join-Path $root 'preservation.json') -Raw | ConvertFrom-Json
$application = $null; $results = @()
try {
    $application = New-Object -ComObject Access.Application
    $application.Visible = $false; $application.AutomationSecurity = 3
    if ([int]$application.AutomationSecurity -ne 3) { throw 'Access did not accept ForceDisable.' }
    foreach ($record in $records) {
        $path = Join-Path $root $record.file
        if ((Get-FileHash -LiteralPath $path).Hash.ToLowerInvariant() -ne $record.sha256) { throw 'The saved native snapshot changed before independent application inspection.' }
        $application.OpenCurrentDatabase($path,$false)
        $observed=@()
        try {
            foreach ($kind in @(@{Names=$record.forms;Type=2;Owner='form'},@{Names=$record.reports;Type=3;Owner='report'},@{Names=$record.macros;Type=4;Owner='macro'},@{Names=$record.vbaModules;Type=5;Owner='module'})) {
                foreach ($name in $kind.Names) {
                    $export = Join-Path $root ($record.file+'.'+$name+'.'+$kind.Owner+'.txt')
                    $application.SaveAsText($kind.Type,$name,$export)
                    $text=[IO.File]::ReadAllText($export)
                    if ($kind.Type -eq 5) { $text = [IO.File]::ReadAllText($export,[Text.Encoding]::GetEncoding(1250)) }
                    if ($name -eq 'FixtureModule' -and -not $text.Contains('FixtureValue = 42')) { throw 'The preserved VBA module differs from the independent source export.' }
                    if ($name -eq 'DesignerModule' -and -not $text.Contains('Zażółć gęślą jaźń')) { throw 'The preserved VBA source lost its code-page characters.' }
                    if ($kind.Type -eq 4 -and -not $text.Contains('Action ="StopMacro"')) { throw 'The preserved action macro differs from its independent definition.' }
                    if ($name -like 'BoundForm*' -and (-not $text.Contains('RecordSource ="Contacts"') -or -not $text.Contains('ControlSource ="DisplayName"'))) { throw 'The preserved bound designer differs from its independent definition.' }
                    if ($name -eq 'BoundReport' -and -not $text.Contains('RecordSource ="Contacts"')) { throw 'The preserved report lost its record source.' }
                    $observed += [ordered]@{name=$name;kind=$kind.Owner;export=[IO.Path]::GetFileName($export);sha256=(Get-FileHash -LiteralPath $export).Hash.ToLowerInvariant()}
                }
            }
            if ($record.dataMacros.Count -gt 0) {
                $export=Join-Path $root ($record.file+'.Contacts.datamacro.xml')
                $application.SaveAsText(12,'Contacts',$export)
                if (-not [IO.File]::ReadAllText($export).Contains('Inert synthetic data macro')) { throw 'The preserved table data macro differs from the independent definition.' }
            }
        } finally { $application.CloseCurrentDatabase() }
        $results += [ordered]@{file=$record.file;wholeFileIdentityBeforeOpen=$true;independentAccessReopen=$true;objects=$observed}
    }
} finally { if ($application) { $application.Quit(2); [Runtime.InteropServices.Marshal]::FinalReleaseComObject($application) | Out-Null } }
$results | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath (Join-Path $root 'independent-preservation.json') -Encoding utf8
$results | ForEach-Object { [pscustomobject]$_ } | Select-Object file,wholeFileIdentityBeforeOpen,independentAccessReopen

param([Parameter(Mandatory)][string] $OutputDirectory)
$ErrorActionPreference = 'Stop'
if ($env:OS -ne 'Windows_NT') { throw 'This independent oracle requires Windows and an installed DAO engine.' }
$root = [IO.Path]::GetFullPath($OutputDirectory)
if (Test-Path -LiteralPath $root) { throw 'Use a fresh output directory; this oracle never replaces existing databases.' }
New-Item -ItemType Directory -Path $root | Out-Null
$appPath = (Get-ItemProperty -LiteralPath 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\MSACCESS.EXE' -ErrorAction SilentlyContinue).'(default)'
$producerBuild = if ($appPath -and (Test-Path -LiteralPath $appPath)) { (Get-Item -LiteralPath $appPath).VersionInfo.FileVersion } else { 'Unavailable' }
$manifest = [ordered]@{schemaVersion=1;producer='Microsoft DAO';producerBuild=$producerBuild;engineVersion=$null;bitness=[IntPtr]::Size*8;license='MIT (synthetic OfficeIMO fixtures)';startup='DAO only; no application, active content, external links or trust-policy changes';files=@();failures=@()}
$engine = $null
try {
    $engine = New-Object -ComObject DAO.DBEngine.120
    $manifest.engineVersion = $engine.Version
    foreach ($probe in @(
        @{Name='jet3.mdb';Version=32;Type=4;Password=$false;Value=42},
        @{Name='password-jet4.mdb';Version=66;Type=4;Password=$true;Value=42},
        @{Name='password-ace.accdb';Version=130;Type=4;Password=$true;Value=42},
        @{Name='large-number.accdb';Version=128;Type=16;Password=$false;Value=[long]5000000000},
        @{Name='extended-date.accdb';Version=128;Type=26;Password=$false;Value=$null}
    )) {
        $path = Join-Path $root $probe.Name
        $db = $null
        try {
            $locale = ';LANGID=0x0409;CP=1252;COUNTRY=0'
            # Public, synthetic fixture password; no user credential is discovered or used.
            if ($probe.Password) { $locale += ';PWD=Fixture123' }
            $db = $engine.CreateDatabase($path, $locale, $probe.Version)
            $table = $db.CreateTableDef('Probe')
            $field = $table.CreateField('Value', $probe.Type)
            $table.Fields.Append($field); $db.TableDefs.Append($table)
            [Runtime.InteropServices.Marshal]::FinalReleaseComObject($field) | Out-Null
            [Runtime.InteropServices.Marshal]::FinalReleaseComObject($table) | Out-Null
            if ($null -ne $probe.Value) {
                # DAO's late-bound Value setter coerces BigInt to Int32 in this host; SQL preserves the supplied synthetic integer.
                $integer = ([long]$probe.Value).ToString([Globalization.CultureInfo]::InvariantCulture)
                $db.Execute('INSERT INTO Probe ([Value]) VALUES (' + $integer + ')', 128)
            }
            $db.Close(); [Runtime.InteropServices.Marshal]::FinalReleaseComObject($db) | Out-Null; $db = $null
            $passwordRequired = $null
            if ($probe.Password) {
                $unprotected = $null
                try { $unprotected = $engine.OpenDatabase($path, $false, $true); $passwordRequired = $false }
                catch { $passwordRequired = $true }
                finally { if ($unprotected) { $unprotected.Close(); [Runtime.InteropServices.Marshal]::FinalReleaseComObject($unprotected) | Out-Null } }
                if (-not $passwordRequired) { throw 'Synthetic protected fixture opened without its password.' }
            }
            $connect = if ($probe.Password) { ';PWD=Fixture123' } else { '' }
            $db = $engine.OpenDatabase($path, $false, $true, $connect)
            $observedTable = $db.TableDefs.Item('Probe')
            try {
                $observedField = $observedTable.Fields.Item('Value')
                try { $observedType = [int]$observedField.Type }
                finally { [Runtime.InteropServices.Marshal]::FinalReleaseComObject($observedField) | Out-Null }
            } finally { [Runtime.InteropServices.Marshal]::FinalReleaseComObject($observedTable) | Out-Null }
            $rows = $db.OpenRecordset('SELECT * FROM Probe', 4)
            try { $values = @(); while (-not $rows.EOF) { $values += $rows.Fields.Item('Value').Value; $rows.MoveNext() } }
            finally { $rows.Close(); [Runtime.InteropServices.Marshal]::FinalReleaseComObject($rows) | Out-Null }
            $db.Close(); [Runtime.InteropServices.Marshal]::FinalReleaseComObject($db) | Out-Null; $db = $null
            $bytes = [IO.File]::ReadAllBytes($path)
            $manifest.files += [ordered]@{path=$probe.Name;headerVersion=[int]$bytes[20];headerSubVersion=[int]$bytes[21];length=$bytes.Length;sha256=(Get-FileHash -LiteralPath $path).Hash.ToLowerInvariant();daoFieldType=$observedType;observedValues=$values;passwordRequired=$passwordRequired;syntheticPassword=if ($probe.Password) {'Fixture123'} else {$null};encryptionRequested=[bool]$probe.Password;limitations=@('Header evidence only in OfficeIMO; native catalogs and protection not decoded', 'Extended date/time field definition only; no precision-value qualification')}
        } catch {
            $manifest.failures += [ordered]@{path=$probe.Name;requestedVersion=$probe.Version;requestedType=$probe.Type;hresult=$_.Exception.HResult;error=$_.Exception.Message}
        } finally { if ($db) { $db.Close(); [Runtime.InteropServices.Marshal]::FinalReleaseComObject($db) | Out-Null } }
    }
} finally { if ($engine) { [Runtime.InteropServices.Marshal]::FinalReleaseComObject($engine) | Out-Null } }
$manifest | ConvertTo-Json -Depth 12 | Set-Content -LiteralPath (Join-Path $root 'manifest.json') -Encoding utf8
$manifest | ConvertTo-Json -Depth 12

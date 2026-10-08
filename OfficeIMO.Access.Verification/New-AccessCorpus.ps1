param(
    [Parameter(Mandatory)][string] $OutputDirectory
)
$ErrorActionPreference = 'Stop'
if ($env:OS -ne 'Windows_NT') { throw 'This independent oracle requires Windows and an installed DAO engine.' }
$root = [IO.Path]::GetFullPath($OutputDirectory)
if (Test-Path -LiteralPath $root) { throw 'Use a fresh output directory; this oracle never replaces existing databases.' }
New-Item -ItemType Directory -Path $root | Out-Null
$engine = $null
$appPath = (Get-ItemProperty -LiteralPath 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\MSACCESS.EXE' -ErrorAction SilentlyContinue).'(default)'
$producerBuild = if ($appPath -and (Test-Path -LiteralPath $appPath)) { (Get-Item -LiteralPath $appPath).VersionInfo.FileVersion } else { 'Access application build not available; DAO version recorded separately' }
$manifest = [ordered]@{ schemaVersion=1; producer='Microsoft DAO'; producerBuild=$producerBuild; engineVersion=$null; bitness=[IntPtr]::Size*8; license='MIT (synthetic OfficeIMO fixtures)'; startup='DAO only; no Access application, macros, modules or external links executed'; files=@() }
try {
    $engine = New-Object -ComObject DAO.DBEngine.120
    $manifest.engineVersion = $engine.Version
    foreach ($profile in @(@{Name='jet4';Extension='mdb';Version=64}, @{Name='ace12';Extension='accdb';Version=128})) {
        $path = Join-Path $root ($profile.Name + '.' + $profile.Extension)
        $db = $null
        try {
            $db = $engine.CreateDatabase($path, ';LANGID=0x0409;CP=1252;COUNTRY=0', $profile.Version)
            $db.Execute('CREATE TABLE Groups (Id LONG CONSTRAINT PK_Groups PRIMARY KEY, Label TEXT(80))', 128)
            $db.Execute('CREATE TABLE Contacts (Id COUNTER CONSTRAINT PK_Contacts PRIMARY KEY, GroupId LONG, DisplayName TEXT(120), Amount CURRENCY, CreatedAt DATETIME, Active YESNO, Notes MEMO)', 128)
            $db.Execute('ALTER TABLE Contacts ADD CONSTRAINT FK_Contacts_Groups FOREIGN KEY (GroupId) REFERENCES Groups (Id)', 128)
            $db.Execute("INSERT INTO Groups (Id, Label) VALUES (1, 'Synthetic group')", 128)
            $db.Execute("INSERT INTO Contacts (GroupId, DisplayName, Amount, CreatedAt, Active, Notes) VALUES (1, 'Ada', 12.3456, #2026-01-02 03:04:05#, True, 'Synthetic text')", 128)
            $db.Execute("INSERT INTO Contacts (GroupId, DisplayName, Amount, CreatedAt, Active) VALUES (1, '', -1.25, NULL, False)", 128)
            $query = $db.CreateQueryDef('ContactsByGroup', 'PARAMETERS selectedGroup Long; SELECT Id, DisplayName FROM Contacts WHERE GroupId = selectedGroup ORDER BY Id;')
            [Runtime.InteropServices.Marshal]::FinalReleaseComObject($query) | Out-Null
            $db.Close(); [Runtime.InteropServices.Marshal]::FinalReleaseComObject($db) | Out-Null
            $db = $engine.OpenDatabase($path, $false, $true)
            # Observe persisted definitions after an independent read-only reopen, not authored expectations.
            $tables = @()
            foreach ($table in $db.TableDefs) {
                if ($table.Name.StartsWith('MSys')) { continue }
                $columns = @($table.Fields | ForEach-Object { [ordered]@{name=$_.Name;daoType=[int]$_.Type;size=[int]$_.Size;required=[bool]$_.Required;attributes=[int]$_.Attributes} })
                $rows = @()
                $rs = $db.OpenRecordset('SELECT * FROM [' + $table.Name + '] ORDER BY Id', 4)
                try {
                    while (-not $rs.EOF) {
                        $row = [ordered]@{}
                        foreach ($field in $rs.Fields) {
                            $value = $field.Value
                            if ($value -is [DBNull]) { $value = $null }
                            if ($value -is [DateTime]) { $value = $value.ToString('yyyy-MM-ddTHH:mm:ss', [Globalization.CultureInfo]::InvariantCulture) }
                            $row[$field.Name] = $value
                        }
                        $rows += $row
                        $rs.MoveNext()
                    }
                } finally { $rs.Close(); [Runtime.InteropServices.Marshal]::FinalReleaseComObject($rs) | Out-Null }
                $indexes = @($table.Indexes | ForEach-Object {
                    [ordered]@{name=$_.Name;primary=[bool]$_.Primary;unique=[bool]$_.Unique;foreign=[bool]$_.Foreign;ignoreNulls=[bool]$_.IgnoreNulls;required=[bool]$_.Required;fields=@($_.Fields | ForEach-Object { [ordered]@{name=$_.Name;attributes=[int]$_.Attributes} })}
                })
                $tables += [ordered]@{name=$table.Name;columns=$columns;indexes=$indexes;rows=$rows}
            }
            $relationships = @($db.Relations | Where-Object { -not $_.Name.StartsWith('MSys') } | ForEach-Object {
                [ordered]@{name=$_.Name;parentTable=$_.Table;childTable=$_.ForeignTable;attributes=[int]$_.Attributes;fields=@($_.Fields | ForEach-Object { [ordered]@{parentColumn=$_.Name;childColumn=$_.ForeignName} })}
            })
            $query = $db.QueryDefs.Item('ContactsByGroup')
            try { $export = [ordered]@{observation='DAO read-only reopen';tables=$tables;query=[ordered]@{name=$query.Name;sql=$query.SQL};relationships=$relationships} }
            finally { [Runtime.InteropServices.Marshal]::FinalReleaseComObject($query) | Out-Null }
            $exportPath = Join-Path $root ($profile.Name + '.expected.json')
            $export | ConvertTo-Json -Depth 12 | Set-Content -LiteralPath $exportPath -Encoding utf8
        } finally { if ($db) { $db.Close(); [Runtime.InteropServices.Marshal]::FinalReleaseComObject($db) | Out-Null } }
        # Reopen through DAO read-only to prove that fixture creation completed and released locks.
        $reopened = $engine.OpenDatabase($path, $false, $true)
        try { $count = $reopened.OpenRecordset('SELECT Count(*) AS N FROM Contacts', 4); $rowCount = [int]$count.Fields.Item('N').Value; $count.Close(); [Runtime.InteropServices.Marshal]::FinalReleaseComObject($count) | Out-Null }
        finally { $reopened.Close(); [Runtime.InteropServices.Marshal]::FinalReleaseComObject($reopened) | Out-Null }
        $bytes = [IO.File]::ReadAllBytes($path)
        $manifest.files += [ordered]@{path=[IO.Path]::GetFileName($path);profile=$profile.Name;headerVersion=[int]$bytes[20];length=$bytes.Length;sha256=(Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash.ToLowerInvariant();expected=[IO.Path]::GetFileName($exportPath);expectedSha256=(Get-FileHash -LiteralPath $exportPath -Algorithm SHA256).Hash.ToLowerInvariant();independentRowCount=$rowCount;limitations=@('No application objects', 'No password or encryption', 'Modern ACE feature upgrades not exercised')}
    }
} finally { if ($engine) { [Runtime.InteropServices.Marshal]::FinalReleaseComObject($engine) | Out-Null } }
$manifest | ConvertTo-Json -Depth 12 | Set-Content -LiteralPath (Join-Path $root 'manifest.json') -Encoding utf8
$manifest | ConvertTo-Json -Depth 12

param(
    [Parameter(Mandatory)][string] $OutputDirectory,
    [Parameter(Mandatory)][string] $ArtifactsDirectory
)
$ErrorActionPreference = 'Stop'
if ($env:OS -ne 'Windows_NT') { throw 'Independent native creation qualification requires installed Windows DAO.' }
$root = [IO.Path]::GetFullPath($OutputDirectory)
if (Test-Path -LiteralPath $root) { throw 'Use a fresh output directory. Existing databases are never opened by this qualification route.' }
$project = Join-Path $PSScriptRoot 'OfficeIMO.Access.Verification.csproj'
dotnet run --project $project -c Release --artifacts-path ([IO.Path]::GetFullPath($ArtifactsDirectory)) -- --bootstrap $root
if ($LASTEXITCODE -ne 0) { throw 'The managed bootstrap generator failed.' }
$engine = $null
$appPath = (Get-ItemProperty -LiteralPath 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\MSACCESS.EXE' -ErrorAction SilentlyContinue).'(default)'
$build = if ($appPath -and (Test-Path -LiteralPath $appPath)) { (Get-Item -LiteralPath $appPath).VersionInfo.FileVersion } else { 'Unavailable' }
$manifest = [ordered]@{schemaVersion=1;producer='OfficeIMO.Access.Verification native feasibility probe';seedUsed=$false;consumer='Microsoft DAO';consumerBuild=$build;engineVersion=$null;bitness=[IntPtr]::Size*8;license='MIT (synthetic OfficeIMO fixtures)';proof='Fresh managed page generation followed by independent DAO schema/index/relationship/value consumption';files=@();limitations=@('Fixed unprotected Jet4/ACE12 scaffold and synthetic schema only; not a production Save codec', 'Restricted observed ASCII General legacy index keys', 'No modern complex system catalog, application-object carrier, encryption, signature or general edit qualification')}
try {
    $engine = New-Object -ComObject DAO.DBEngine.120
    $manifest.engineVersion = $engine.Version
    foreach ($name in @('catalog-scaffold.mdb','catalog-scaffold.accdb')) {
        $path = Join-Path $root $name
        $db = $null
        try {
            $db = $engine.OpenDatabase($path, $false, $true)
            $tables = @()
            foreach ($tableName in @('Groups','Contacts')) {
                $table = $db.TableDefs.Item($tableName)
                try {
                    $columns = @($table.Fields | ForEach-Object { [ordered]@{name=$_.Name;daoType=[int]$_.Type;size=[int]$_.Size;attributes=[int]$_.Attributes} })
                    $indexes = @($table.Indexes | ForEach-Object { [ordered]@{name=$_.Name;primary=[bool]$_.Primary;unique=[bool]$_.Unique;foreign=[bool]$_.Foreign;fields=@($_.Fields | ForEach-Object {$_.Name})} })
                } finally { [Runtime.InteropServices.Marshal]::FinalReleaseComObject($table) | Out-Null }
                $rows = @(); $recordset = $db.OpenRecordset('SELECT * FROM [' + $tableName + '] ORDER BY Id', 4)
                try {
                    while (-not $recordset.EOF) {
                        $row = [ordered]@{}
                        foreach ($field in $recordset.Fields) {
                            $value = $field.Value
                            if ($value -is [DBNull]) { $value = $null }
                            if ($value -is [DateTime]) { $value = $value.ToString('yyyy-MM-ddTHH:mm:ss',[Globalization.CultureInfo]::InvariantCulture) }
                            $row[$field.Name] = $value
                        }
                        $rows += $row; $recordset.MoveNext()
                    }
                } finally { $recordset.Close(); [Runtime.InteropServices.Marshal]::FinalReleaseComObject($recordset) | Out-Null }
                $tables += [ordered]@{name=$tableName;columns=$columns;indexes=$indexes;rows=$rows}
            }
            $contacts = $tables | Where-Object {$_.name -eq 'Contacts'}
            $groups = $tables | Where-Object {$_.name -eq 'Groups'}
            if ($groups.rows.Count -ne 1 -or $contacts.rows.Count -ne 2) { throw 'Native row counts differ from the authored logical schema.' }
            $expectedTypes=@{Id=4;GroupId=4;DisplayName=10;Amount=5;CreatedAt=8;Active=1}
            if($contacts.columns.Count -ne $expectedTypes.Count){throw 'Independent native column count differs from the logical schema.'}
            foreach($column in $contacts.columns){if(-not $expectedTypes.ContainsKey($column.name) -or $column.daoType -ne $expectedTypes[$column.name]){throw 'Independent native column type differs from the logical schema.'}}
            if ($contacts.rows[0].Id -ne 1 -or $contacts.rows[0].DisplayName -ne 'Ada' -or $contacts.rows[0].Amount -ne ([decimal]12.3456) -or $contacts.rows[0].CreatedAt -ne '2026-01-02T03:04:05' -or -not $contacts.rows[0].Active) { throw 'Native first-row typed values failed independent comparison.' }
            if ($contacts.rows[1].Id -ne 2 -or $contacts.rows[1].DisplayName -ne '' -or $contacts.rows[1].Amount -ne ([decimal](-1.25)) -or $null -ne $contacts.rows[1].CreatedAt -or $contacts.rows[1].Active) { throw 'Native second-row null/empty/typed values failed independent comparison.' }
            if (-not ($contacts.indexes | Where-Object {$_.name -eq 'PK_Contacts' -and $_.primary -and $_.unique -and $_.fields[0] -eq 'Id'})) { throw 'The native primary key was not independently observed.' }
            if (-not ($contacts.indexes | Where-Object {$_.name -eq 'ContactGroups' -and $_.foreign -and $_.fields[0] -eq 'GroupId'})) { throw 'The native foreign index was not independently observed.' }
            $relation = $db.Relations.Item('ContactGroups')
            try { $relationship = [ordered]@{name=$relation.Name;parentTable=$relation.Table;childTable=$relation.ForeignTable;attributes=[int]$relation.Attributes;fields=@($relation.Fields | ForEach-Object { [ordered]@{parentColumn=$_.Name;childColumn=$_.ForeignName} })} }
            finally { [Runtime.InteropServices.Marshal]::FinalReleaseComObject($relation) | Out-Null }
            if ($relationship.parentTable -ne 'Groups' -or $relationship.childTable -ne 'Contacts' -or $relationship.fields[0].parentColumn -ne 'Id' -or $relationship.fields[0].childColumn -ne 'GroupId') { throw 'The native relationship differs from its logical definition.' }
            $seek = $db.OpenRecordset('Contacts', 1)
            try { $seek.Index = 'PK_Contacts'; $seek.Seek('=',2); if ($seek.NoMatch -or $seek.Fields.Item('DisplayName').Value -ne '') { throw 'Independent primary-index seek failed.' } }
            finally { $seek.Close(); [Runtime.InteropServices.Marshal]::FinalReleaseComObject($seek) | Out-Null }
            $export = [ordered]@{observation='DAO read-only reopen and index seek';tables=$tables;relationship=$relationship}
            $exportPath = Join-Path $root ($name + '.expected.json')
            $export | ConvertTo-Json -Depth 12 | Set-Content -LiteralPath $exportPath -Encoding utf8
        } finally { if ($db) { $db.Close(); [Runtime.InteropServices.Marshal]::FinalReleaseComObject($db) | Out-Null } }
        # Mutate a separate owned verification copy only. Generation itself never reads a seed/template file.
        $copy = Join-Path $root ('constraint-check-' + $name)
        Copy-Item -LiteralPath $path -Destination $copy
        $db = $null; $rejections = @()
        try {
            $db = $engine.OpenDatabase($copy, $false, $false)
            foreach ($violation in @(@{Sql='INSERT INTO Contacts (Id, GroupId) VALUES (1,1)';DaoError=3022}, @{Sql='INSERT INTO Contacts (Id, GroupId) VALUES (3,999)';DaoError=3201})) {
                $rejected = $false
                try { $db.Execute($violation.Sql,128) }
                catch {
                    $code=$_.Exception.HResult -band 65535
                    if($code -ne $violation.DaoError){throw}
                    $rejected = $true; $rejections += [ordered]@{statement=$violation.Sql;daoError=$code;hresult=$_.Exception.HResult;error=$_.Exception.Message}
                }
                if (-not $rejected) { throw 'An independently executed constraint violation was unexpectedly accepted.' }
            }
        } finally { if ($db) { $db.Close(); [Runtime.InteropServices.Marshal]::FinalReleaseComObject($db) | Out-Null } }
        $bytes = [IO.File]::ReadAllBytes($path)
        $manifest.files += [ordered]@{path=$name;headerVersion=[int]$bytes[20];length=$bytes.Length;sha256=(Get-FileHash -LiteralPath $path).Hash.ToLowerInvariant();expected=[IO.Path]::GetFileName($exportPath);expectedSha256=(Get-FileHash -LiteralPath $exportPath).Hash.ToLowerInvariant();independentRowCount=2;primaryIndexSeek=$true;constraintRejections=$rejections}
    }
} finally { if ($engine) { [Runtime.InteropServices.Marshal]::FinalReleaseComObject($engine) | Out-Null } }
$manifest | ConvertTo-Json -Depth 14 | Set-Content -LiteralPath (Join-Path $root 'manifest.json') -Encoding utf8
$manifest | ConvertTo-Json -Depth 14

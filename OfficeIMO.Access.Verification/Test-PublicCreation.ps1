param([Parameter(Mandatory)][string]$CorpusDirectory)
$ErrorActionPreference = 'Stop'
if ($env:OS -ne 'Windows_NT') { throw 'Independent creation qualification requires Windows DAO and Microsoft Access.' }
$root = [IO.Path]::GetFullPath($CorpusDirectory)
if (-not (Test-Path -LiteralPath (Join-Path $root 'creation.json'))) { throw 'Generate an owned public model corpus with --create first.' }
$engine = New-Object -ComObject DAO.DBEngine.120
$results = @()
try {
    foreach ($extension in @('mdb','accdb')) {
        foreach ($boundary in @('wide','allocation','lengths','long-columns','keys')) {
            $db=$engine.OpenDatabase((Join-Path $root ($boundary+'.'+$extension)),$false,$true)
            try {
                if ($boundary -eq 'keys') {
                    $rs=$db.OpenRecordset('KeyValues',1)
                    try {
                        $rs.Index='NameKey'; $names=@('', '_', 'alpha beta', 'ALPHA_BETA', ' alpha', 'gamma', 'z9')
                        for($i=0;$i -lt $names.Length;$i++) { $rs.Seek('=',$names[$i]); if($rs.NoMatch -or $rs.Fields.Item('Id').Value -ne $i+1) { throw "Text collation seek differs for '$($names[$i])'." }; if($rs.Fields.Item('Owner').Value -ne 123 -or [string]$rs.Fields.Item('SID').Value -ne '{guid {01234567-89AB-CDEF-0123-456789ABCDEF}}') { throw 'Ordinary Owner/SID field storage differs.' } }
                        $rs.Index='CompositeKey'
                        for($i=0;$i -lt $names.Length;$i++) { $rs.Seek('=',[short]([short]::MinValue+$i),[byte]($i+1)); if($rs.NoMatch -or $rs.Fields.Item('Id').Value -ne $i+1) { throw 'Signed Int16/Byte composite seek differs.' } }
                    } finally { $rs.Close(); [Runtime.InteropServices.Marshal]::FinalReleaseComObject($rs)|Out-Null }
                    $relation=$db.Relations.Item('TreeParent')
                    try { if($relation.Table -ne 'Tree' -or $relation.ForeignTable -ne 'Tree') { throw 'Self relationship differs.' } }
                    finally { [Runtime.InteropServices.Marshal]::FinalReleaseComObject($relation)|Out-Null }
                } elseif ($boundary -eq 'long-columns') {
                    $rs=$db.OpenRecordset('LongColumns',4)
                    try { foreach($i in (0..254)) { $bytes=[byte[]]$rs.Fields.Item('Field'+$i.ToString('D3')).Value; if($bytes.Length -ne 300 -or @($bytes | Where-Object { $_ -ne $i }).Count -ne 0) { throw 'Long-column usage maps or payload differs.' } } }
                    finally { $rs.Close(); [Runtime.InteropServices.Marshal]::FinalReleaseComObject($rs)|Out-Null }
                } elseif ($boundary -eq 'wide') {
                    $td=$db.TableDefs.Item('Wide'); if ($td.Fields.Count -ne 255) { throw 'Multi-page table definition lost fields.' }; [Runtime.InteropServices.Marshal]::FinalReleaseComObject($td)|Out-Null
                    $rs=$db.OpenRecordset('Wide',4)
                    try { foreach ($i in (0..254)) { if ($rs.Fields.Item('Field'+$i.ToString('D3')).Value -ne $i) { throw 'Wide row value differs.' } } }
                    finally { $rs.Close(); [Runtime.InteropServices.Marshal]::FinalReleaseComObject($rs)|Out-Null }
                } elseif ($boundary -eq 'allocation') {
                    $rs=$db.OpenRecordset('Allocated',1)
                    try { $rs.Index='PK_Allocated'; foreach ($id in @(1001,3000,7000)) { $rs.Seek('=',$id); if ($rs.NoMatch -or ([string]$rs.Fields.Item('Padding').Value).Length -ne 255) { throw 'Reference-map allocation or index seek differs.' } }; $rs.MoveLast(); if ($rs.RecordCount -ne 6000) { throw 'Reference-map row count differs.' } }
                    finally { $rs.Close(); [Runtime.InteropServices.Marshal]::FinalReleaseComObject($rs)|Out-Null }
                } else {
                    $rs=$db.OpenRecordset('Lengths',4)
                    try { while (-not $rs.EOF) { $length=[int]$rs.Fields.Item('Id').Value; $bytes=[byte[]]$rs.Fields.Item('Payload').Value; if ($bytes.Length -ne $length) { throw "Long value length $length differs." }; for($i=0;$i -lt $length;$i++) { if($bytes[$i] -ne $i%251) { throw 'Boundary payload differs.' } }; $rs.MoveNext() } }
                    finally { $rs.Close(); [Runtime.InteropServices.Marshal]::FinalReleaseComObject($rs)|Out-Null }
                }
            } finally { $db.Close(); [Runtime.InteropServices.Marshal]::FinalReleaseComObject($db)|Out-Null }
        }
        $empty = $engine.OpenDatabase((Join-Path $root ('empty.'+$extension)), $false, $true)
        try { if (@($empty.TableDefs | Where-Object {-not $_.Name.StartsWith('MSys')}).Count -ne 0) { throw 'Fresh empty database unexpectedly has user tables.' } }
        finally { $empty.Close(); [Runtime.InteropServices.Marshal]::FinalReleaseComObject($empty) | Out-Null }
        $path = Join-Path $root ('created.'+$extension)
        $db = $engine.OpenDatabase($path,$false,$true)
        try {
            if ($db.Properties.Item('AppTitle').Value -ne 'Synthetic native creation') { throw 'Database AppTitle property differs.' }
            $td = $db.TableDefs.Item('Contacts')
            try {
                $schema = @($td.Fields | ForEach-Object { [ordered]@{name=$_.Name;type=[int]$_.Type;attributes=[int]$_.Attributes;size=[int]$_.Size} })
                $indexes = @($td.Indexes | ForEach-Object { [ordered]@{name=$_.Name;unique=[bool]$_.Unique;primary=[bool]$_.Primary;foreign=[bool]$_.Foreign;fields=@($_.Fields | ForEach-Object {$_.Name})} })
                if (($td.Fields.Item('Id').Attributes -band 16) -eq 0) { throw 'Native AutoNumber attribute was lost.' }
            } finally { [Runtime.InteropServices.Marshal]::FinalReleaseComObject($td) | Out-Null }
            $relation = $db.Relations.Item('ContactGroups')
            try { if ($relation.Table -ne 'Groups' -or $relation.ForeignTable -ne 'Contacts') { throw 'Relationship mapping differs.' } }
            finally { [Runtime.InteropServices.Marshal]::FinalReleaseComObject($relation) | Out-Null }
            $rs = $db.OpenRecordset('Contacts',1)
            try {
                $rs.Index = 'PK_Contacts'
                foreach ($id in @(1,2,400,401,4999,5000)) { $rs.Seek('=',$id); if ($rs.NoMatch -or $rs.Fields.Item('DisplayName').Value -ne ('Contact'+($id-1).ToString('D5'))) { throw "Primary index seek failed for $id." } }
                $rs.Seek('=',1)
                if ([decimal]$rs.Fields.Item('Amount').Value -ne [decimal]::Parse('-1.2345',[Globalization.CultureInfo]::InvariantCulture) -or [datetime]$rs.Fields.Item('CreatedAt').Value -ne [datetime]::new(2026,1,2,3,4,5) -or -not [bool]$rs.Fields.Item('Active').Value) { throw 'Independent Currency/DateTime/Boolean differs.' }
                if ([string]$rs.Fields.Item('Identifier').Value -ne '{guid {01234567-89AB-CDEF-0123-456789ABCDEF}}' -or [double]$rs.Fields.Item('Ratio').Value -ne -1.25 -or [single]$rs.Fields.Item('SingleValue').Value -ne 1.5 -or [short]$rs.Fields.Item('Small').Value -ne -32768 -or [byte]$rs.Fields.Item('Octet').Value -ne 255) { throw 'Independent GUID/floating/integer values differ.' }
                if ([decimal]$rs.Fields.Item('Precise').Value -ne [decimal]::Parse('1234567890123456789.123456789',[Globalization.CultureInfo]::InvariantCulture)) { throw 'Independent Decimal precision differs.' }
                $notes = [string]$rs.Fields.Item('Notes').Value
                $expectedNotes = -join (1..1000 | ForEach-Object { "Ł🙂 Synthetic text`r`n" })
                if ($notes -cne $expectedNotes) { throw 'Independent long Unicode text differs.' }
                $bytes = [byte[]]$rs.Fields.Item('Payload').Value
                if ($bytes.Length -ne 20000) { throw 'Independent long binary length differs.' }
                for ($i=0;$i -lt $bytes.Length;$i++) { if ($bytes[$i] -ne ($i*17%251)) { throw 'Independent long binary payload differs.' } }
                $rs.Seek('=',2)
                if ($rs.Fields.Item('Notes').Value -is [DBNull] -or [string]$rs.Fields.Item('Notes').Value -ne '') { throw 'Empty memo differs from null.' }
                $rs.Index = 'NameUnique'
                foreach ($id in @(1,2500,5000)) { $rs.Seek('=',('Contact'+($id-1).ToString('D5'))); if ($rs.NoMatch -or $rs.Fields.Item('Id').Value -ne $id) { throw 'Text index seek failed.' } }
                $rs.MoveLast(); if ($rs.RecordCount -ne 5000) { throw 'Independent row count differs.' }
            } finally { $rs.Close(); [Runtime.InteropServices.Marshal]::FinalReleaseComObject($rs) | Out-Null }
        } finally { $db.Close(); [Runtime.InteropServices.Marshal]::FinalReleaseComObject($db) | Out-Null }
        $copy = Join-Path $root ('access-edited.'+$extension)
        if (Test-Path -LiteralPath $copy) { throw 'The owned Access edit output already exists.' }
        Copy-Item -LiteralPath $path -Destination $copy
        $app = New-Object -ComObject Access.Application
        try {
            $app.Visible=$false; $app.AutomationSecurity=3
            $app.OpenCurrentDatabase($copy)
            $db = $app.CurrentDb()
            $applicationEngine = $app.DBEngine
            try {
                $rejections = @()
                foreach ($case in @(@{Sql="INSERT INTO Contacts (Id,GroupId,DisplayName) VALUES (5000,1,'DuplicateId')";Code=3022},@{Sql="INSERT INTO Contacts (GroupId,DisplayName) VALUES (999,'Orphan')";Code=3201},@{Sql="INSERT INTO Contacts (GroupId,DisplayName) VALUES (1,'contact00000')";Code=3022})) {
                    $rejected=$false
                    try { $db.Execute($case.Sql,128) }
                    catch {
                        $numbers=@($applicationEngine.Errors | ForEach-Object { [int]$_.Number }) + @($_.Exception.HResult -band 65535)
                        if ($case.Code -notin $numbers) { throw }
                        $rejected=$true; $rejections += $case.Code
                    }
                    if (-not $rejected) { throw 'Access accepted an invalid constraint operation.' }
                }
                $db.Execute("UPDATE Contacts SET DisplayName='ChangedByAccess' WHERE Id=5000",128)
                $db.Execute("INSERT INTO Contacts (GroupId,DisplayName) VALUES (1,'AppendedByAccess')",128)
                $db.Execute('DELETE FROM Contacts WHERE Id=4999',128)
                $rs = $db.OpenRecordset("SELECT Id FROM Contacts WHERE DisplayName='AppendedByAccess'",4)
                try { $allocated=[int]$rs.Fields.Item('Id').Value; if ($allocated -le 5000) { throw 'Native AutoNumber state was not advanced.' } }
                finally { $rs.Close(); [Runtime.InteropServices.Marshal]::FinalReleaseComObject($rs) | Out-Null }
            } finally { $db.Close(); [Runtime.InteropServices.Marshal]::FinalReleaseComObject($db) | Out-Null; [Runtime.InteropServices.Marshal]::FinalReleaseComObject($applicationEngine) | Out-Null }
            $app.CloseCurrentDatabase()
        } finally { $app.Quit(); [Runtime.InteropServices.Marshal]::FinalReleaseComObject($app) | Out-Null }
        $db=$engine.OpenDatabase($copy,$false,$true)
        try {
            $rs=$db.OpenRecordset("SELECT Count(*) AS N FROM Contacts",4)
            try { if ($rs.Fields.Item('N').Value -ne 5000) { throw 'Access edit/save changed the expected row count.' } }
            finally { $rs.Close(); [Runtime.InteropServices.Marshal]::FinalReleaseComObject($rs) | Out-Null }
        } finally { $db.Close(); [Runtime.InteropServices.Marshal]::FinalReleaseComObject($db) | Out-Null }
        $results += [ordered]@{file=[IO.Path]::GetFileName($path);schema=$schema;indexes=$indexes;rows=5000;primaryAndTextSeeks=$true;longValuesExact=$true;decimalExact=$true;wide255Columns=$true;long255Columns=$true;textAndCompositeCollation=$true;selfRelationship=$true;referenceMap6000Rows=$true;longValueBoundaries=$true;accessOpenEditSave=$true;constraintRejections=$rejections;appendedAutoNumber=$allocated;editedFile=[IO.Path]::GetFileName($copy);sourceSha256=(Get-FileHash $path).Hash.ToLowerInvariant();editedSha256=(Get-FileHash $copy).Hash.ToLowerInvariant()}
    }
} finally { [Runtime.InteropServices.Marshal]::FinalReleaseComObject($engine) | Out-Null }
$results | ConvertTo-Json -Depth 12 | Set-Content -LiteralPath (Join-Path $root 'independent-creation.json') -Encoding utf8
$results | ForEach-Object { [pscustomobject]@{File=$_.file;Rows=$_.rows;AccessEditSave=$_.accessOpenEditSave;AutoNumber=$_.appendedAutoNumber} }

param([Parameter(Mandatory)][string] $OutputDirectory)
$ErrorActionPreference = 'Stop'
$root = [IO.Path]::GetFullPath($OutputDirectory)
if (Test-Path -LiteralPath $root) { throw 'Use a fresh output directory.' }
New-Item -ItemType Directory -Path $root | Out-Null
$engine = $null
$manifest = [ordered]@{ schemaVersion=1; producer='Microsoft DAO; installed ACE OLE DB for the ANSI-92 Decimal definition'; bitness=[IntPtr]::Size*8; license='MIT (synthetic OfficeIMO fixtures)'; files=@(); failures=@() }
try {
    $engine = New-Object -ComObject DAO.DBEngine.120
    $manifest.engineVersion = $engine.Version
    foreach ($profile in @(@{Name='values-jet4.mdb'; Version=64}, @{Name='values-ace.accdb'; Version=128})) {
        $path = Join-Path $root $profile.Name
        $db = $null
        try {
            $db = $engine.CreateDatabase($path, ';LANGID=0x0415;CP=1250;COUNTRY=0', $profile.Version)
            $db.Execute('CREATE TABLE Scalars (Id COUNTER CONSTRAINT PK_Scalars PRIMARY KEY, Tiny BYTE, Small SHORT, Whole LONG, [Real] SINGLE, Wide DOUBLE, [Money] CURRENCY, Occurred DATETIME, Flag YESNO, Token GUID, Label TEXT(255), Notes MEMO, Payload LONGBINARY, FixedBinary BINARY(12))', 128)
            # ANSI-92 decimal precision is authored through the installed ACE validation oracle. DAO's late-bound setter does not retain it.
            $db.Close(); [Runtime.InteropServices.Marshal]::FinalReleaseComObject($db) | Out-Null; $db=$null
            $connection=New-Object -ComObject ADODB.Connection
            try { $connection.Open('Provider=Microsoft.ACE.OLEDB.16.0;Data Source='+$path); $connection.Execute('ALTER TABLE Scalars ADD Precise DECIMAL(28,9)') | Out-Null }
            finally { $connection.Close(); [Runtime.InteropServices.Marshal]::FinalReleaseComObject($connection) | Out-Null }
            $db=$engine.OpenDatabase($path,$false,$false)
            $table = $db.TableDefs.Item('Scalars')
            try {
                $label = $table.Fields.Item('Label'); $label.AllowZeroLength = $true
                $label.Properties.Append($label.CreateProperty('Caption', 10, 'Unicode label'))
                $label.Properties.Append($label.CreateProperty('DisplayControl', 3, 111))
                $label.Properties.Append($label.CreateProperty('RowSourceType', 10, 'Value List'))
                $label.Properties.Append($label.CreateProperty('RowSource', 12, 'Alpha;Beta'))
                [Runtime.InteropServices.Marshal]::FinalReleaseComObject($label) | Out-Null
                if ($profile.Version -eq 128) {
                    $memo = $table.Fields.Item('Notes'); $memo.Properties.Append($memo.CreateProperty('TextFormat', 3, 1)); [Runtime.InteropServices.Marshal]::FinalReleaseComObject($memo) | Out-Null
                }
            } finally { [Runtime.InteropServices.Marshal]::FinalReleaseComObject($table) | Out-Null }
            $db.Execute("INSERT INTO Scalars (Tiny, Small, Whole, [Real], Wide, [Money], Precise, Occurred, Flag, Label) VALUES (255, -32768, 2147483647, 1.25, -2.5, -922337203685477.5808, 1234567890123456789.123456789, #1899-12-29 06:07:08#, True, 'Zażółć gęślą jaźń 漢字 العربية 🙂')", 128)
            $db.Execute("INSERT INTO Scalars (Tiny, Small, Whole, [Real], Wide, [Money], Precise, Flag, Label) VALUES (0, 32767, -2147483648, -0.125, 1.234567890123, 0.0001, -0.000000001, False, '')", 128)
            $rows = $db.OpenRecordset('Scalars', 2)
            try {
                $rows.MoveFirst(); $rows.Edit()
                $rows.Fields.Item('Token').Value = '{00112233-4455-6677-8899-AABBCCDDEEFF}'
                $rows.Fields.Item('Notes').Value = ('<div>Ł🙂 native long text</div>' * 500)
                [byte[]] $payload = 0..19999 | ForEach-Object { [byte](($_ * 17) % 251) }
                $rows.Fields.Item('Payload').AppendChunk($payload)
                $rows.Fields.Item('FixedBinary').Value = [byte[]](0..11)
                $rows.Update()
            } finally { $rows.Close(); [Runtime.InteropServices.Marshal]::FinalReleaseComObject($rows) | Out-Null }
            $db.Execute('CREATE TABLE Fragmented (Id LONG CONSTRAINT PK_Fragmented PRIMARY KEY, Content TEXT(255), Tail MEMO)', 128)
            for ($i=0; $i -lt 180; $i++) { $db.Execute("INSERT INTO Fragmented (Id, Content) VALUES ($i, 'short')", 128) }
            $db.Execute('DELETE FROM Fragmented WHERE Id MOD 3 = 0', 128)
            $long = 'Growing row creates native overflow storage ' * 5
            $db.Execute("UPDATE Fragmented SET Content = '$long' WHERE Id MOD 2 = 1", 128)
            $db.Execute('CREATE TABLE ParentKeys (A LONG, B LONG, CONSTRAINT PK_ParentKeys PRIMARY KEY (A,B))', 128)
            $db.Execute('CREATE TABLE ChildKeys (Id LONG CONSTRAINT PK_ChildKeys PRIMARY KEY, A LONG, B LONG)', 128)
            $relation = $db.CreateRelation('CompositeKeys', 'ParentKeys', 'ChildKeys', 4352)
            foreach ($name in @('A','B')) { $field=$relation.CreateField($name); $field.ForeignName=$name; $relation.Fields.Append($field); [Runtime.InteropServices.Marshal]::FinalReleaseComObject($field) | Out-Null }
            $db.Relations.Append($relation); [Runtime.InteropServices.Marshal]::FinalReleaseComObject($relation) | Out-Null
            $db.Execute('CREATE INDEX DescendingPair ON ChildKeys (B DESC, A)', 128)
            if ($profile.Version -eq 128) {
                $complex = $db.CreateTableDef('Structured')
                $id=$complex.CreateField('Id',4); $complex.Fields.Append($id)
                $tags=$complex.CreateField('Tags',109); $complex.Fields.Append($tags)
                $files=$complex.CreateField('Files',101); $complex.Fields.Append($files)
                $db.TableDefs.Append($complex)
                foreach ($item in @($id,$tags,$files,$complex)) { [Runtime.InteropServices.Marshal]::FinalReleaseComObject($item) | Out-Null }
                $rs=$db.OpenRecordset('Structured',2)
                try {
                    $rs.AddNew(); $rs.Fields.Item('Id').Value=1
                    $tags=$rs.Fields.Item('Tags').Value
                    try { foreach ($tag in @('Alpha','Ł🙂')) { $tags.AddNew(); $tags.Fields.Item('Value').Value=$tag; $tags.Update() } } finally { $tags.Close(); [Runtime.InteropServices.Marshal]::FinalReleaseComObject($tags) | Out-Null }
                    $attachments=$rs.Fields.Item('Files').Value
                    $attachmentPath=Join-Path $root 'synthetic.bin'; [IO.File]::WriteAllBytes($attachmentPath,[byte[]](0..255))
                    try { $attachments.AddNew(); $attachments.Fields.Item('FileData').LoadFromFile($attachmentPath); $attachments.Update() } finally { $attachments.Close(); [Runtime.InteropServices.Marshal]::FinalReleaseComObject($attachments) | Out-Null }
                    $rs.Update()
                } finally { $rs.Close(); [Runtime.InteropServices.Marshal]::FinalReleaseComObject($rs) | Out-Null }
            }
            $db.Close(); [Runtime.InteropServices.Marshal]::FinalReleaseComObject($db) | Out-Null; $db=$null
            $db=$engine.OpenDatabase($path,$false,$true)
            $observations=@()
            foreach ($name in @('Scalars','Fragmented','ParentKeys','ChildKeys')) {
                $table=$db.TableDefs.Item($name)
                $columns=@($table.Fields | ForEach-Object { [ordered]@{name=$_.Name; daoType=[int]$_.Type; size=[int]$_.Size; attributes=[int]$_.Attributes} })
                $rs=$db.OpenRecordset('SELECT * FROM ['+$name+'] ORDER BY '+ $(if($name -eq 'ParentKeys'){'A,B'}else{'Id'}),4)
                try {
                    $values=@(); while(-not $rs.EOF) {
                        $value=[ordered]@{}
                        foreach($field in $rs.Fields) {
                            $v=$field.Value
                            if($v -is [DBNull]){$v=$null}
                            elseif($v -is [byte[]]){$v=[ordered]@{base64=[Convert]::ToBase64String($v); length=$v.Length}}
                            elseif($v -is [DateTime]){$v=$v.ToString('O',[Globalization.CultureInfo]::InvariantCulture)}
                            elseif($v -is [decimal]){$v=[ordered]@{decimal=$v.ToString([Globalization.CultureInfo]::InvariantCulture)}}
                            $value[$field.Name]=$v
                        }
                        $values+=$value; $rs.MoveNext()
                    }
                } finally { $rs.Close(); [Runtime.InteropServices.Marshal]::FinalReleaseComObject($rs) | Out-Null }
                $observations += [ordered]@{name=$name; columns=$columns; rows=$values}
                [Runtime.InteropServices.Marshal]::FinalReleaseComObject($table) | Out-Null
            }
            $expected=Join-Path $root ($profile.Name+'.expected.json')
            $structured=$null
            if($profile.Version -eq 128) {
                $rs=$db.OpenRecordset('Structured',4)
                try {
                    $tags=$rs.Fields.Item('Tags').Value; $tagValues=@()
                    try { while(-not $tags.EOF){$tagValues+=$tags.Fields.Item('Value').Value; $tags.MoveNext()} } finally{$tags.Close(); [Runtime.InteropServices.Marshal]::FinalReleaseComObject($tags)|Out-Null}
                    $files=$rs.Fields.Item('Files').Value; $attachmentValues=@()
                    try {
                        while(-not $files.EOF) {
                            $decoded=Join-Path $root ('decoded-attachment-'+$attachmentValues.Count+'.bin')
                            $files.Fields.Item('FileData').SaveToFile($decoded)
                            $attachmentValues+=[ordered]@{name=$files.Fields.Item('FileName').Value; type=$files.Fields.Item('FileType').Value; decodedBase64=[Convert]::ToBase64String([IO.File]::ReadAllBytes($decoded)); decodedSha256=(Get-FileHash -LiteralPath $decoded).Hash.ToLowerInvariant()}
                            $files.MoveNext()
                        }
                    }finally{$files.Close(); [Runtime.InteropServices.Marshal]::FinalReleaseComObject($files)|Out-Null}
                    $structured=[ordered]@{id=$rs.Fields.Item('Id').Value; tags=$tagValues; attachments=$attachmentValues}
                }finally{$rs.Close(); [Runtime.InteropServices.Marshal]::FinalReleaseComObject($rs)|Out-Null}
            }
            [ordered]@{observation='DAO read-only reopen; attachment SaveToFile'; tables=$observations; structured=$structured} | ConvertTo-Json -Depth 12 | Set-Content -LiteralPath $expected -Encoding utf8
            $manifest.files += [ordered]@{path=$profile.Name; headerVersion=[int]([IO.File]::ReadAllBytes($path))[20]; length=(Get-Item -LiteralPath $path).Length; sha256=(Get-FileHash -LiteralPath $path).Hash.ToLowerInvariant(); expected=[IO.Path]::GetFileName($expected); expectedSha256=(Get-FileHash -LiteralPath $expected).Hash.ToLowerInvariant()}
        } finally { if($db){$db.Close(); [Runtime.InteropServices.Marshal]::FinalReleaseComObject($db) | Out-Null} }
    }
} finally { if($engine){[Runtime.InteropServices.Marshal]::FinalReleaseComObject($engine) | Out-Null} }
$manifest | ConvertTo-Json -Depth 12 | Set-Content -LiteralPath (Join-Path $root 'manifest.json') -Encoding utf8
$manifest | ConvertTo-Json -Depth 12

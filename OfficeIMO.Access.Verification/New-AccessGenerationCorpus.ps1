param([Parameter(Mandatory)][string] $OutputDirectory)
$ErrorActionPreference='Stop'
$root=[IO.Path]::GetFullPath($OutputDirectory)
if(Test-Path -LiteralPath $root){throw 'Use a fresh output directory.'}
New-Item -ItemType Directory -Path $root | Out-Null
$app=$null; $engine=$null; $manifest=[ordered]@{schemaVersion=1; producer='Microsoft Access and DAO'; license='MIT (synthetic OfficeIMO fixtures)'; files=@()}
function Get-LegacyObservation([string] $Path) {
    $oracle=New-Object -ComObject DAO.DBEngine.120; $source=$null; $rows=$null
    try {
        $source=$oracle.OpenDatabase($Path,$false,$true); $rows=$source.OpenRecordset('Legacy',4)
        $values=[ordered]@{}; foreach($name in @('Id','Label','Notes','Link','Occurred')){$values[$name]=$rows.Fields.Item($name).Value}
        $queries=[ordered]@{}; foreach($name in @('LegacyNames','LegacyRawUnion','LegacyStarExpression','LegacyGrouped')){
            $query=$source.QueryDefs.Item($name); try{$queries[$name]=$query.SQL}finally{[Runtime.InteropServices.Marshal]::FinalReleaseComObject($query)|Out-Null}
        }
        return [ordered]@{consumer='DAO independent read-only reopen'; values=$values; queries=$queries}
    }finally{
        if($rows){$rows.Close(); [Runtime.InteropServices.Marshal]::FinalReleaseComObject($rows)|Out-Null}
        if($source){$source.Close(); [Runtime.InteropServices.Marshal]::FinalReleaseComObject($source)|Out-Null}
        [Runtime.InteropServices.Marshal]::FinalReleaseComObject($oracle)|Out-Null
    }
}
try {
    $app=New-Object -ComObject Access.Application; $app.Visible=$false; $app.AutomationSecurity=3
    if([int]$app.AutomationSecurity -ne 3){throw 'ForceDisable was not accepted.'}
    foreach($profile in @(@{Name='access2000.mdb';Format=9},@{Name='access2002-2003.mdb';Format=10})) {
        $path=Join-Path $root $profile.Name; $app.NewCurrentDatabase($path,$profile.Format)
        try {
            $db=$app.CurrentDb()
            try {
                $db.Execute('CREATE TABLE Legacy (Id LONG CONSTRAINT PK_Legacy PRIMARY KEY, Label TEXT(120), Notes MEMO, Occurred DATETIME)',128)
                $table=$db.TableDefs.Item('Legacy'); $link=$table.CreateField('Link',12); $link.Attributes=32770; $table.Fields.Append($link)
                [Runtime.InteropServices.Marshal]::FinalReleaseComObject($link)|Out-Null
                [Runtime.InteropServices.Marshal]::FinalReleaseComObject($table)|Out-Null
                $db.Execute("INSERT INTO Legacy (Id,Label,Notes,Link,Occurred) VALUES (1,'Zażółć 漢字','Legacy Memo','Caption#https://example.invalid/path#section#Tooltip',#2000-02-29 12:34:56#)",128)
                $query=$db.CreateQueryDef('LegacyNames','SELECT Id, Label FROM Legacy ORDER BY Id;'); [Runtime.InteropServices.Marshal]::FinalReleaseComObject($query)|Out-Null
                $query=$db.CreateQueryDef('LegacyRawUnion','SELECT Id FROM Legacy UNION ALL SELECT Id FROM Legacy;'); [Runtime.InteropServices.Marshal]::FinalReleaseComObject($query)|Out-Null
                $query=$db.CreateQueryDef('LegacyStarExpression','SELECT *, Id+1 AS NextId FROM Legacy;'); [Runtime.InteropServices.Marshal]::FinalReleaseComObject($query)|Out-Null
                $query=$db.CreateQueryDef('LegacyGrouped','SELECT Label, Count(*) AS Total FROM Legacy GROUP BY Label;'); [Runtime.InteropServices.Marshal]::FinalReleaseComObject($query)|Out-Null
                $connect=$db.CreateTableDef('LocalLink'); $connect.Connect=';DATABASE='+$path; $connect.SourceTableName='Legacy'; $db.TableDefs.Append($connect); [Runtime.InteropServices.Marshal]::FinalReleaseComObject($connect)|Out-Null
                $protected=[IO.Path]::GetFullPath((Join-Path $PSScriptRoot '../OfficeIMO.Access.Tests/Fixtures/Profiles/password-jet4.mdb'))
                $connect=$db.CreateTableDef('CredentialLink'); $connect.Connect=';DATABASE='+$protected+';PWD=Fixture123'; $connect.SourceTableName='Probe'; $connect.Attributes=131072; $db.TableDefs.Append($connect); [Runtime.InteropServices.Marshal]::FinalReleaseComObject($connect)|Out-Null
                $fileFormat=[int]$app.CurrentProject.FileFormat
                $version=$db.Properties.Item('AccessVersion').Value
            } finally{$db.Close(); [Runtime.InteropServices.Marshal]::FinalReleaseComObject($db)|Out-Null}
        } finally{$app.CloseCurrentDatabase()}
        $manifest.files+=[ordered]@{path=$profile.Name; accessFileFormat=$fileFormat; accessVersion=$version; sha256=(Get-FileHash -LiteralPath $path).Hash.ToLowerInvariant(); length=(Get-Item -LiteralPath $path).Length; linkedTarget='Synthetic self-link; native readers must never resolve it'; daoObservation=(Get-LegacyObservation $path)}
    }
    $app.Quit(2); [Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)|Out-Null; $app=$null
    $engine=New-Object -ComObject DAO.DBEngine.120
    $path=Join-Path $root 'extended-values.accdb'; $db=$engine.CreateDatabase($path,';LANGID=0x0409;CP=1252',128)
    try {
        $table=$db.CreateTableDef('Extended'); $id=$table.CreateField('Id',4); $table.Fields.Append($id); $date=$table.CreateField('Occurred',26); $table.Fields.Append($date); $number=$table.CreateField('Large',16); $table.Fields.Append($number); $db.TableDefs.Append($table)
        foreach($item in @($id,$date,$number,$table)){[Runtime.InteropServices.Marshal]::FinalReleaseComObject($item)|Out-Null}
        $db.Execute('INSERT INTO Extended (Id, Large) VALUES (1, 9223372036854775807)',128)
        $rows=$db.OpenRecordset('Extended',2)
        try{$rows.MoveFirst(); $rows.Edit(); $rows.Fields.Item('Occurred').Value='0002-01-02T03:04:05.1234567'; $rows.Update()}finally{$rows.Close(); [Runtime.InteropServices.Marshal]::FinalReleaseComObject($rows)|Out-Null}
    }finally{$db.Close(); [Runtime.InteropServices.Marshal]::FinalReleaseComObject($db)|Out-Null}
    $db=$engine.OpenDatabase($path,$false,$true)
    try{$rows=$db.OpenRecordset('Extended',4); try{$observed=[ordered]@{date=$rows.Fields.Item('Occurred').Value; large=$rows.Fields.Item('Large').Value}}finally{$rows.Close(); [Runtime.InteropServices.Marshal]::FinalReleaseComObject($rows)|Out-Null}}finally{$db.Close(); [Runtime.InteropServices.Marshal]::FinalReleaseComObject($db)|Out-Null}
    $manifest.files+=[ordered]@{path='extended-values.accdb'; sha256=(Get-FileHash -LiteralPath $path).Hash.ToLowerInvariant(); length=(Get-Item -LiteralPath $path).Length; headerVersion=[int]([IO.File]::ReadAllBytes($path))[20]; daoObservation=$observed}
    # This DAO host authors scale-zero extended dates. Qualify a controlled scale-seven value with an independent DAO consumer.
    # This fixture edit is a validation probe, not a production writer or independent Access-producer claim.
    $bytes=[IO.File]::ReadAllBytes($path); $matches=[regex]::Matches([Text.Encoding]::ASCII.GetString($bytes),'[0-9]{19}:[0-9]{19}:0\x00')
    if($matches.Count -ne 1){throw 'The synthetic extended-date probe has no unique scale-zero value.'}
    $match=$matches[0]; $units=[long]::Parse($match.Value.Substring(20,19),[Globalization.CultureInfo]::InvariantCulture)
    $encoded=$match.Value.Substring(0,20)+($units*10000000+1234567).ToString('D19',[Globalization.CultureInfo]::InvariantCulture)+":7`0"
    [Array]::Copy([Text.Encoding]::ASCII.GetBytes($encoded),0,$bytes,$match.Index,42)
    $fractionalPath=Join-Path $root 'fractional-values.accdb'; [IO.File]::WriteAllBytes($fractionalPath,$bytes)
    $db=$engine.OpenDatabase($fractionalPath,$false,$true)
    try{$rows=$db.OpenRecordset('Extended',4); try{$fractional=[ordered]@{date=$rows.Fields.Item('Occurred').Value;large=$rows.Fields.Item('Large').Value}}finally{$rows.Close(); [Runtime.InteropServices.Marshal]::FinalReleaseComObject($rows)|Out-Null}}finally{$db.Close(); [Runtime.InteropServices.Marshal]::FinalReleaseComObject($db)|Out-Null}
    if(-not $fractional.date.EndsWith('.1234567')){throw 'Independent DAO did not retain the seven-digit fractional probe.'}
    $manifest.files+=[ordered]@{path='fractional-values.accdb'; producer='Controlled 42-byte value mutation of the Access/DAO source'; consumer='DAO read-only reopen'; sha256=(Get-FileHash -LiteralPath $fractionalPath).Hash.ToLowerInvariant(); length=$bytes.Length; daoObservation=$fractional}
    $path=Join-Path $root 'calculated-ace14.accdb'; $db=$engine.CreateDatabase($path,';LANGID=0x0409;CP=1252',128)
    try {
        $table=$db.CreateTableDef('Calculated'); $id=$table.CreateField('Id',4); $table.Fields.Append($id)
        $value=$table.CreateField('Value',4); $value.Expression='[Id]+1'; $table.Fields.Append($value); $db.TableDefs.Append($table)
        foreach($item in @($id,$value,$table)){[Runtime.InteropServices.Marshal]::FinalReleaseComObject($item)|Out-Null}
        $db.Execute('INSERT INTO Calculated (Id) VALUES (41)',128)
    }finally{$db.Close(); [Runtime.InteropServices.Marshal]::FinalReleaseComObject($db)|Out-Null}
    $db=$engine.OpenDatabase($path,$false,$true)
    try{$rs=$db.OpenRecordset('Calculated',4); try{$calculated=[ordered]@{id=$rs.Fields.Item('Id').Value;value=$rs.Fields.Item('Value').Value}}finally{$rs.Close(); [Runtime.InteropServices.Marshal]::FinalReleaseComObject($rs)|Out-Null}}finally{$db.Close(); [Runtime.InteropServices.Marshal]::FinalReleaseComObject($db)|Out-Null}
    $manifest.files+=[ordered]@{path='calculated-ace14.accdb'; sha256=(Get-FileHash -LiteralPath $path).Hash.ToLowerInvariant(); length=(Get-Item -LiteralPath $path).Length; headerVersion=[int]([IO.File]::ReadAllBytes($path))[20]; daoObservation=$calculated; nativeContract='Exact calculated representation and expression metadata; OfficeIMO does not evaluate it'}
}finally{if($app){$app.Quit(2); [Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)|Out-Null}; if($engine){[Runtime.InteropServices.Marshal]::FinalReleaseComObject($engine)|Out-Null}}
$manifest | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath (Join-Path $root 'manifest.json') -Encoding utf8
$manifest | ConvertTo-Json -Depth 10

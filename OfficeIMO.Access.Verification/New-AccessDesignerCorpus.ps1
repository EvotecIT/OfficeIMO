param([Parameter(Mandatory)][string] $OutputDirectory, [switch] $IncludeDataMacro, [switch] $IncludeEmbeddedMacro)
$ErrorActionPreference = 'Stop'
if ($env:OS -ne 'Windows_NT') { throw 'The designer oracle requires Windows Access.' }
$root = [IO.Path]::GetFullPath($OutputDirectory)
if (Test-Path -LiteralPath $root) { throw 'Use a fresh output directory; this producer creates synthetic databases only.' }
New-Item -ItemType Directory -Path $root | Out-Null
$application = $null
$results = @()
try {
    $application = New-Object -ComObject Access.Application
    $application.Visible = $false
    $application.AutomationSecurity = 3
    $producerVersion = $application.Version
    if ([int]$application.AutomationSecurity -ne 3) { throw 'Access did not accept ForceDisable.' }
    foreach ($profile in @(@{Name='designer-jet4.mdb';Format=9}, @{Name='designer-ace12.accdb';Format=12})) {
        $path = Join-Path $root $profile.Name
        $application.NewCurrentDatabase($path, $profile.Format)
        $database = $null
        try {
            $database = $application.CurrentDb()
            $database.Execute('CREATE TABLE Contacts (Id COUNTER CONSTRAINT PK_Contacts PRIMARY KEY, DisplayName TEXT(120), GroupId LONG)')
            $database.Execute("INSERT INTO Contacts (DisplayName,GroupId) VALUES ('Synthetic',1)")
            $database.Execute('CREATE TABLE Groups (Id LONG, Label TEXT(80))')
            $database.Execute("INSERT INTO Groups VALUES (1,'Group A')")
            $query = $database.CreateQueryDef('ContactQuery', 'PARAMETERS [pId] Long; SELECT Id, DisplayName FROM Contacts WHERE Id=[pId];')
            [Runtime.InteropServices.Marshal]::FinalReleaseComObject($query) | Out-Null
            $objects = @()
            foreach ($number in @(1,2)) {
                $name = "BoundForm$number"
                $formText = [IO.File]::ReadAllText((Join-Path $PSScriptRoot 'Assets/BoundForm.txt'))
                $formText = $formText.Replace('Synthetic form 1', "Synthetic form $number").Replace('Title1',"Title$number").Replace('DisplayName1',"DisplayName$number").Replace('GroupChoice1',"GroupChoice$number")
                if ($IncludeEmbeddedMacro -and $profile.Format -eq 12 -and $number -eq 1) {
                    $formText = $formText.Replace('Version =19','Version =21').Replace('VersionRequired =19','VersionRequired =20')
                    $embedded = @('                    OnClick ="[Embedded Macro]"','                    OnClickEmMacro = Begin','                        Version =196611','                        ColumnsShown =0','                        Begin','                            Action ="StopMacro"','                        End','                    End') -join "`r`n"
                    $formText = $formText.Replace('                    Caption ="Caption 1"', '                    Caption ="Caption 1"' + "`r`n" + $embedded)
                }
                $import = Join-Path $root ($name + '.txt')
                [IO.File]::WriteAllText($import,$formText,[Text.Encoding]::Unicode)
                $application.LoadFromText(2,$name,$import)
                $export = Join-Path $root ($profile.Name + ".$name.txt")
                $application.SaveAsText(2,$name,$export)
                $objects += [ordered]@{kind='form';name=$name;caption="Synthetic form $number";recordSource='Contacts';width=4800;detailHeight=3600;export=[IO.Path]::GetFileName($export);exportSha256=(Get-FileHash -LiteralPath $export).Hash.ToLowerInvariant()}
            }
            $reportFile = if ($profile.Format -eq 9) { 'objects-jet4.mdb.report.txt' } else { 'objects-ace12.accdb.report.txt' }
            $reportSource = Join-Path $PSScriptRoot ("../OfficeIMO.Access.Tests/Fixtures/Application/$reportFile")
            $reportText = [IO.File]::ReadAllText($reportSource).Replace('Synthetic OfficeIMO report','Bound report')
            # Access exports this modern flag in legacy text but its legacy importer rejects it.
            $reportText = [regex]::Replace($reportText, '(?m)^    NoSaveCTIWhenDisabled =1\r?\n', '')
            $reportText = $reportText.Replace('    Caption ="Bound report"', '    RecordSource ="Contacts"' + "`n" + '    Caption ="Bound report"')
            $import=Join-Path $root 'BoundReport.txt'
            [IO.File]::WriteAllText($import,$reportText,[Text.Encoding]::Unicode)
            $application.LoadFromText(3,'BoundReport',$import)
            $export = Join-Path $root ($profile.Name + '.BoundReport.txt'); $application.SaveAsText(3,'BoundReport',$export)
            $objects += [ordered]@{kind='report';name='BoundReport';recordSource='Contacts';export=[IO.Path]::GetFileName($export);exportSha256=(Get-FileHash -LiteralPath $export).Hash.ToLowerInvariant()}
            $moduleText = Join-Path $root 'DesignerModule.bas'
            $moduleLines = @('Attribute VB_Name = "DesignerModule"','Option Compare Database','Option Explicit',"'Zażółć gęślą jaźń",'Public Function ConstantValue() As Long','    ConstantValue = 42','End Function')
            $moduleEncoding = [Text.Encoding]::GetEncoding(1250, [Text.EncoderFallback]::ExceptionFallback, [Text.DecoderFallback]::ExceptionFallback)
            [IO.File]::WriteAllText($moduleText, ($moduleLines -join "`r`n") + "`r`n", $moduleEncoding)
            $application.LoadFromText(5,'DesignerModule',$moduleText)
            $export = Join-Path $root ($profile.Name + '.DesignerModule.txt'); $application.SaveAsText(5,'DesignerModule',$export)
            $objects += [ordered]@{kind='vba-module';name='DesignerModule';export=[IO.Path]::GetFileName($export);exportSha256=(Get-FileHash -LiteralPath $export).Hash.ToLowerInvariant()}
            foreach ($name in @('StopHere','AutoExec')) {
                $macroText = Join-Path $root ($name + '.txt')
                @('Version =196611','ColumnsShown =0','Begin','    Action ="StopMacro"','End') | Set-Content -LiteralPath $macroText -Encoding ascii
                $application.LoadFromText(4,$name,$macroText)
                $export=Join-Path $root ($profile.Name + ".$name.txt"); $application.SaveAsText(4,$name,$export)
                $objects += [ordered]@{kind='action-macro';name=$name;action='StopMacro';export=[IO.Path]::GetFileName($export);exportSha256=(Get-FileHash -LiteralPath $export).Hash.ToLowerInvariant()}
            }
            $property = $database.CreateProperty('AppTitle',10,'Synthetic OfficeIMO application'); $database.Properties.Append($property)
            [Runtime.InteropServices.Marshal]::FinalReleaseComObject($property) | Out-Null
            if ($IncludeDataMacro -and $profile.Format -eq 12) {
                $macroText = Join-Path $root 'ContactDataMacro.xml'
                $xml = '<DataMacros xmlns="http://schemas.microsoft.com/office/accessservices/2009/11/application"><DataMacro Event="AfterInsert"><Statements><Comment>Inert synthetic data macro</Comment></Statements></DataMacro></DataMacros>'
                [IO.File]::WriteAllText($macroText,$xml,[Text.Encoding]::Unicode)
                $application.LoadFromText(12,'Contacts',$macroText)
                $export=Join-Path $root ($profile.Name+'.Contacts.datamacro.xml'); $application.SaveAsText(12,'Contacts',$export)
                $objects += [ordered]@{kind='table-data-macro';name='Contacts';export=[IO.Path]::GetFileName($export);exportSha256=(Get-FileHash -LiteralPath $export).Hash.ToLowerInvariant()}
            }
        } finally {
            if ($database) { $database.Close(); [Runtime.InteropServices.Marshal]::FinalReleaseComObject($database) | Out-Null }
            $application.CloseCurrentDatabase()
        }
        $results += [ordered]@{path=$profile.Name;sha256=(Get-FileHash -LiteralPath $path).Hash.ToLowerInvariant();objects=$objects}
    }
} finally { if ($application) { $application.Quit(2); [Runtime.InteropServices.Marshal]::FinalReleaseComObject($application) | Out-Null } }
[ordered]@{schemaVersion=1;producer='Microsoft Access';producerVersion=$producerVersion;license='MIT (synthetic fixtures)';security='Owned new instance; ForceDisable=3; all content remains inert; no user databases';files=$results} | ConvertTo-Json -Depth 12 | Set-Content -LiteralPath (Join-Path $root 'manifest.json') -Encoding utf8
$results | ConvertTo-Json -Depth 8

param([Parameter(Mandatory)][string] $OutputRoot)
$ErrorActionPreference='Stop'
$application=$null
$opened=$false
$results=@()
try {
    $application=New-Object -ComObject Access.Application
    $application.Visible=$false
    $application.AutomationSecurity=3
    if([int]$application.AutomationSecurity -ne 3){throw 'Macro disabling was not accepted.'}
    $records=Get-Content -LiteralPath (Join-Path $OutputRoot 'vba-editing.json') -Raw | ConvertFrom-Json
    foreach($record in $records) {
        $path=Join-Path $OutputRoot ([IO.Path]::GetFileName($record.output))
        $before=(Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash
        $application.OpenCurrentDatabase($path,$false)
        $opened=$true
        try {
            $application.SaveAsText(5,$record.module,($path+'.module-export.txt'))
            $application.DoCmd.OpenModule($record.module)
            $application.DoCmd.RunCommand(126)
            $module=$application.Modules.Item($record.module)
            if([int]$module.Type -ne [int]$record.moduleKind){throw 'Native edited module kind differs.'}
            $native=[string]$module.Lines(1,$module.CountOfLines)
            $expected=[IO.File]::ReadAllText($path+'.expected-module.txt') -replace '(?m)^Attribute [^\r\n]*\r?\n',''
            [IO.File]::WriteAllText($path+'.native-module.txt',$native)
            if($native.TrimEnd([char[]]"`r`n") -ne $expected.TrimEnd([char[]]"`r`n")){throw 'Native edited source differs.'}
            [Runtime.InteropServices.Marshal]::FinalReleaseComObject($module)|Out-Null
            $application.DoCmd.Close(5,$record.module,1)
            $ordinary=@($application.CurrentProject.AllModules | ForEach-Object {$_.Name})
            if($ordinary -notcontains 'AddedModule' -or $ordinary -notcontains 'AddedClass'){throw 'Native module inventory differs.'}
            $application.DoCmd.OpenModule('AddedModule')
            $module=$application.Modules.Item('AddedModule')
            $native=[string]$module.Lines(1,$module.CountOfLines)
            $expected=[IO.File]::ReadAllText($path+'.expected-added.txt') -replace '(?m)^Attribute [^\r\n]*\r?\n',''
            if($native.TrimEnd([char[]]"`r`n") -ne $expected.TrimEnd([char[]]"`r`n")){throw 'Native added source differs.'}
            [Runtime.InteropServices.Marshal]::FinalReleaseComObject($module)|Out-Null
            $application.DoCmd.Close(5,'AddedModule',1)
            $application.DoCmd.OpenModule('AddedClass')
            $module=$application.Modules.Item('AddedClass')
            if([int]$module.Type -ne 1 -or -not ([string]$module.Lines(1,$module.CountOfLines)).Contains('ClassValue = 45')){throw 'Native class source or kind differs.'}
            [Runtime.InteropServices.Marshal]::FinalReleaseComObject($module)|Out-Null
            $application.DoCmd.RunCommand(126)
            $application.DoCmd.Close(5,'AddedClass',1)
            $results += [ordered]@{file=[IO.Path]::GetFileName($path);profile=$record.profile;writerSha256=$before;nativeAccessReopened=$true;
                sourceExact=$true;addedSourceExact=$true;classKindVerified=$true;vbeCompileCompleted=$true;codeExecuted=$false}
        } finally {if($opened){$application.CloseCurrentDatabase();$opened=$false}}
    }
} finally {if($application){$application.Quit(2);[Runtime.InteropServices.Marshal]::FinalReleaseComObject($application)|Out-Null}}
$results | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath (Join-Path $OutputRoot 'native-vba-editing.json') -Encoding utf8
$results | ConvertTo-Json -Depth 6

param([Parameter(Mandatory)][string]$OutputRoot)
$ErrorActionPreference='Stop';$application=$null;$opened=$false;$results=@()
try {
    $application=New-Object -ComObject Access.Application;$application.Visible=$false;$application.AutomationSecurity=3
    if([int]$application.AutomationSecurity -ne 3){throw 'Macro disabling was not accepted'}
    foreach($file in @('event-only.accdb','cleared.accdb','first-code-behind.accdb')) {
        $path=Join-Path $OutputRoot $file;$sha=(Get-FileHash $path).Hash
        $application.OpenCurrentDatabase($path,$false);$opened=$true
        try {
            $application.SaveAsText(2,'EventForm',($path+'.form.txt'))
            $application.DoCmd.OpenForm('EventForm',1);$form=$application.Forms.Item('EventForm');$source=$null
            $expected=if($file -eq 'event-only.accdb'){'=Len("no-vba")'}elseif($file -eq 'cleared.accdb'){''}else{'[Event Procedure]'}
            if([string]$form.Controls.Item('EventLabel').OnClick -ne $expected){throw 'Native event binding mismatch'}
            if($file -eq 'first-code-behind.accdb') {
                if(-not $form.HasModule -or $form.Module.Type -ne 1){throw 'Native first class binding missing'}
                $source=[string]$form.Module.Lines(1,$form.Module.CountOfLines)
                if($source.Replace("`r`n","`n").TrimEnd("`n") -ne [IO.File]::ReadAllText($path+'.form.expected.txt')){throw 'Native first class source differs from the writer output'}
                $application.DoCmd.RunCommand(126)
            } elseif($form.HasModule){throw 'Designer-only edit created native code-behind'}
            $application.DoCmd.Close(2,'EventForm',2);[Runtime.InteropServices.Marshal]::FinalReleaseComObject($form)|Out-Null
            $results += [ordered]@{file=$file;writerSha256=$sha;nativeBinding=$expected;source=$source;ordinaryModuleCount=$application.CurrentProject.AllModules.Count;codeExecuted=$false}
        }finally{if($opened){$application.CloseCurrentDatabase();$opened=$false}}
    }
}finally{if($application){$application.Quit(2);[Runtime.InteropServices.Marshal]::FinalReleaseComObject($application)|Out-Null}}
$results|ConvertTo-Json -Depth 6|Set-Content (Join-Path $OutputRoot 'native-empty-designer.json') -Encoding utf8
$results|ConvertTo-Json -Depth 6

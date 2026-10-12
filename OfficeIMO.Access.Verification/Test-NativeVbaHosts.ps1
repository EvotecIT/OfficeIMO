param([Parameter(Mandatory)][string]$OutputRoot)
$Events=$true
$ErrorActionPreference='Stop'; $application=$null; $opened=$false; $results=@()
try {
    $application=New-Object -ComObject Access.Application; $application.Visible=$false; $application.AutomationSecurity=3
    if([int]$application.AutomationSecurity -ne 3){throw 'Macro disabling was not accepted'}
    foreach($file in @('designer-jet4.mdb','designer-ace12.accdb')) {
        $path=Join-Path $OutputRoot $file; $before=(Get-FileHash -LiteralPath $path).Hash
        $application.OpenCurrentDatabase($path,$false); $opened=$true
        try {
            $application.SaveAsText(2,'BoundForm1',($path+'.form.txt'))
            $application.SaveAsText(3,'BoundReport',($path+'.report.txt'))
            $application.DoCmd.OpenModule('DesignerModule'); $application.DoCmd.RunCommand(126); $application.DoCmd.Close(5,'DesignerModule',1)
            $application.DoCmd.OpenForm('BoundForm1',1); $form=$application.Forms.Item('BoundForm1')
            if(-not $form.HasModule -or $form.Module.Type -ne 1){throw 'New form code-behind is not a native class'}
            $formSource=[string]$form.Module.Lines(1,$form.Module.CountOfLines)
            if($formSource.Replace("`r`n","`n").TrimEnd("`n") -ne [IO.File]::ReadAllText($path+'.form.expected.txt')){throw 'Native form source differs from the writer output'}
            if($Events -and $file.EndsWith('.accdb')) {
                if($form.OnOpen -ne '[Event Procedure]' -or $form.Controls.Item('GroupChoice1').AfterUpdate -ne '=Len("proof")' -or $form.Controls.Item('Title1').OnClick -ne '=Len("click")'){throw 'Native form event changes were not accepted'}
                if($form.Controls.Item('DisplayName1').OnClick -ne '=Len("text-click")' -or $form.Controls.Item('DisplayName1').AfterUpdate -ne '=Len("text-update")' -or $form.Controls.Item('GroupChoice1').OnClick -ne '=Len("combo-click")'){throw 'Native text-box/combo-box event changes were not accepted'}
            }
            $application.DoCmd.Close(2,'BoundForm1',2); [Runtime.InteropServices.Marshal]::FinalReleaseComObject($form)|Out-Null
            $application.DoCmd.OpenReport('BoundReport',1); $report=$application.Reports.Item('BoundReport')
            if(-not $report.HasModule -or $report.Module.Type -ne 1){throw 'New report code-behind is not a native class'}
            $reportSource=[string]$report.Module.Lines(1,$report.Module.CountOfLines)
            if($reportSource.Replace("`r`n","`n").TrimEnd("`n") -ne [IO.File]::ReadAllText($path+'.report.expected.txt')){throw 'Native report source differs from the writer output'}
            if($Events -and $file.EndsWith('.accdb') -and $report.OnOpen -ne '[Event Procedure]'){throw 'Native report event changes were not accepted'}
            if($Events -and $file.EndsWith('.accdb') -and $report.Controls.Item('Label0').OnClick -ne '=Len("report-click")'){throw 'Native report-control event change was not accepted'}
            $application.DoCmd.Close(3,'BoundReport',2); [Runtime.InteropServices.Marshal]::FinalReleaseComObject($report)|Out-Null
            if($Events -and $file.EndsWith('.accdb')) {
                $application.DoCmd.OpenForm('BoundForm2',1);$form2=$application.Forms.Item('BoundForm2')
                if($form2.Controls.Item('Title2').OnClick -ne '=Len("new")' -or $form2.HasModule){throw 'Designer-only event changed the host module state'}
                $application.DoCmd.Close(2,'BoundForm2',2);[Runtime.InteropServices.Marshal]::FinalReleaseComObject($form2)|Out-Null
            }
            $ordinary=@($application.CurrentProject.AllModules|ForEach-Object {$_.Name})
            if($ordinary.Count -ne 1 -or $ordinary[0] -ne 'DesignerModule'){throw 'A host class entered the ordinary module inventory'}
            $results += [ordered]@{file=$file;writerSha256=$before;sourceExact=$true;nativeClassBindings=$true;ordinaryModuleInventory=$ordinary;vbeCompile=$true;eventBindingsVerified=($Events -and $file.EndsWith('.accdb'));codeExecuted=$false}
        } finally {if($opened){$application.CloseCurrentDatabase();$opened=$false}}
    }
} finally {if($application){$application.Quit(2);[Runtime.InteropServices.Marshal]::FinalReleaseComObject($application)|Out-Null}}
$results | ConvertTo-Json -Depth 5 | Set-Content (Join-Path $OutputRoot 'native-vba-hosts.json') -Encoding utf8
$results | ConvertTo-Json -Depth 5

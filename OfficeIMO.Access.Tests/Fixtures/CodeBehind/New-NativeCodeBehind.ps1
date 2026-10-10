param([Parameter(Mandatory)][string]$FixtureRoot, [Parameter(Mandatory)][string]$OutputRoot)
$ErrorActionPreference='Stop'
New-Item -ItemType Directory -Path $OutputRoot -ErrorAction SilentlyContinue | Out-Null
$application=$null; $results=@(); $opened=$false
try {
    $application=New-Object -ComObject Access.Application
    $application.Visible=$false; $application.AutomationSecurity=3
    if ([int]$application.AutomationSecurity -ne 3) { throw 'Macro disabling was not accepted.' }
    foreach ($file in @('Designer/designer-jet4.mdb','Designer/designer-ace12.accdb')) {
        $target=Join-Path $OutputRoot ([IO.Path]::GetFileName($file))
        if (Test-Path -LiteralPath $target) { throw ('Oracle output already exists: '+$target) }
        Copy-Item -LiteralPath (Join-Path $FixtureRoot $file) -Destination $target
        Write-Output ('Opening owned oracle '+$target)
        $application.OpenCurrentDatabase($target,$false); $opened=$true
        try {
            $application.DoCmd.OpenForm('BoundForm1',1)
            $form=$application.Forms.Item('BoundForm1')
            $form.HasModule=$true
            $form.Module.AddFromString("Private Sub Form_Open(Cancel As Integer)`r`n    ' Synthetic inert form handler, marker 51.`r`nEnd Sub`r`n")
            $form.OnOpen='[Event Procedure]'
            $form.Module.AddFromString("Private Sub GroupChoice1_AfterUpdate()`r`n    ' Synthetic inert control handler, marker 52.`r`nEnd Sub`r`n")
            $form.Controls.Item('GroupChoice1').AfterUpdate='[Event Procedure]'
            $formSource=[string]$form.Module.Lines(1,$form.Module.CountOfLines)
            $formModuleName=[string]$form.Module.Name
            $application.DoCmd.Save(2,'BoundForm1')
            $application.DoCmd.Close(2,'BoundForm1',1)
            [Runtime.InteropServices.Marshal]::FinalReleaseComObject($form)|Out-Null
            $application.SaveAsText(2,'BoundForm1',($target+'.form.txt'))
            $application.DoCmd.OpenReport('BoundReport',1)
            $report=$application.Reports.Item('BoundReport')
            $report.HasModule=$true
            $report.Module.AddFromString("Private Sub Report_Open(Cancel As Integer)`r`n    ' Synthetic inert report handler, marker 53.`r`nEnd Sub`r`n")
            $report.OnOpen='[Event Procedure]'
            $reportSource=[string]$report.Module.Lines(1,$report.Module.CountOfLines)
            $reportModuleName=[string]$report.Module.Name
            $application.DoCmd.Save(3,'BoundReport')
            $application.DoCmd.Close(3,'BoundReport',1)
            [Runtime.InteropServices.Marshal]::FinalReleaseComObject($report)|Out-Null
            $application.SaveAsText(3,'BoundReport',($target+'.report.txt'))
            $results += [ordered]@{file=[IO.Path]::GetFileName($target);form='BoundForm1';report='BoundReport';formModule=$formModuleName;reportModule=$reportModuleName;formSource=$formSource;reportSource=$reportSource;formOpen='[Event Procedure]';control='GroupChoice1';controlAfterUpdate='[Event Procedure]';reportOpen='[Event Procedure]';codeExecuted=$false}
        } finally { if ($opened) { $application.CloseCurrentDatabase(); $opened=$false } }
    }
} finally { if ($application) { $application.Quit(2); [Runtime.InteropServices.Marshal]::FinalReleaseComObject($application)|Out-Null } }
$results | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath (Join-Path $OutputRoot 'native-code-behind.json') -Encoding utf8
$results | ConvertTo-Json -Depth 8

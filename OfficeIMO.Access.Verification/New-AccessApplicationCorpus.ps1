param([Parameter(Mandatory)][string] $OutputDirectory)
$ErrorActionPreference = 'Stop'
if ($env:OS -ne 'Windows_NT') { throw 'The application-object oracle requires installed Windows Access.' }
$root = [IO.Path]::GetFullPath($OutputDirectory)
if (Test-Path -LiteralPath $root) { throw 'Use a fresh directory. Existing databases are never opened or replaced by this producer.' }
New-Item -ItemType Directory -Path $root | Out-Null
$application = $null
$results = @()
try {
    # Always a new owned instance; never attach to the user's Access application.
    $application = New-Object -ComObject Access.Application
    $application.Visible = $false
    $application.AutomationSecurity = 3
    $security = [int]$application.AutomationSecurity
    if ($security -ne 3) { throw 'Access did not accept ForceDisable; no database is opened.' }
    foreach ($profile in @(@{Name='objects-jet4.mdb';Format=9}, @{Name='objects-ace12.accdb';Format=12})) {
        $path = Join-Path $root $profile.Name
        $application.NewCurrentDatabase($path, $profile.Format)
        $objectResults = @()
        try {
            $form = $application.CreateForm()
            $formName = $form.Name
            $form.Caption = 'Synthetic OfficeIMO form'
            $application.DoCmd.Save(2, $formName)
            $application.DoCmd.Close(2, $formName, 1)
            [Runtime.InteropServices.Marshal]::FinalReleaseComObject($form) | Out-Null
            $application.DoCmd.Rename('FixtureForm', 2, $formName)
            $formExport = Join-Path $root ($profile.Name + '.form.txt')
            $application.SaveAsText(2, 'FixtureForm', $formExport)
            $objectResults += [ordered]@{kind='form';name='FixtureForm';producer='Access.CreateForm';export=[IO.Path]::GetFileName($formExport);exportSha256=(Get-FileHash -LiteralPath $formExport).Hash.ToLowerInvariant()}
            $report = $application.CreateReport()
            $reportName = $report.Name
            $report.Caption = 'Synthetic OfficeIMO report'
            $control = $application.CreateReportControl($reportName, 100, 0, '', '', 100, 100, 4000, 400)
            $control.Caption = 'Synthetic report output'
            [Runtime.InteropServices.Marshal]::FinalReleaseComObject($control) | Out-Null
            $application.DoCmd.Save(3, $reportName)
            $application.DoCmd.Close(3, $reportName, 1)
            [Runtime.InteropServices.Marshal]::FinalReleaseComObject($report) | Out-Null
            $application.DoCmd.Rename('FixtureReport', 3, $reportName)
            $reportExport = Join-Path $root ($profile.Name + '.report.txt')
            $application.SaveAsText(3, 'FixtureReport', $reportExport)
            $objectResults += [ordered]@{kind='report';name='FixtureReport';producer='Access.CreateReport';export=[IO.Path]::GetFileName($reportExport);exportSha256=(Get-FileHash -LiteralPath $reportExport).Hash.ToLowerInvariant()}
            # Import a synthetic inert module. No procedure is invoked and there is no startup/event binding.
            $moduleText = Join-Path $root 'FixtureModule.bas'
            @('Attribute VB_Name = "FixtureModule"', 'Option Compare Database', 'Option Explicit', 'Public Function FixtureValue() As Long', '    FixtureValue = 42', 'End Function') | Set-Content -LiteralPath $moduleText -Encoding ascii
            try {
                $application.LoadFromText(5, 'FixtureModule', $moduleText)
                $moduleExport = Join-Path $root ($profile.Name + '.module.txt')
                $application.SaveAsText(5, 'FixtureModule', $moduleExport)
                $objectResults += [ordered]@{kind='vba-module';name='FixtureModule';status='produced-inert';export=[IO.Path]::GetFileName($moduleExport);exportSha256=(Get-FileHash -LiteralPath $moduleExport).Hash.ToLowerInvariant()}
            } catch { $objectResults += [ordered]@{kind='vba-module';status='failed';error=$_.Exception.Message} }
            $macroText = Join-Path $root 'FixtureMacro.txt'
            # SaveAsText's macro envelope is distinct from clipboard XML. Import an inert StopMacro definition; never run it.
            @('Version =196611', 'ColumnsShown =0', 'Begin', '    Action ="StopMacro"', 'End') | Set-Content -LiteralPath $macroText -Encoding ascii
            try {
                $application.LoadFromText(4, 'FixtureMacro', $macroText)
                $macroExport = Join-Path $root ($profile.Name + '.macro.txt')
                $application.SaveAsText(4, 'FixtureMacro', $macroExport)
                $objectResults += [ordered]@{kind='action-macro';name='FixtureMacro';status='produced-inert';export=[IO.Path]::GetFileName($macroExport);exportSha256=(Get-FileHash -LiteralPath $macroExport).Hash.ToLowerInvariant()}
            } catch { $objectResults += [ordered]@{kind='action-macro';status='failed-import';hresult=$_.Exception.HResult} }
            $pdf = Join-Path $root ($profile.Name + '.report.pdf')
            try {
                $application.DoCmd.OutputTo(3, 'FixtureReport', 'PDF Format (*.pdf)', $pdf, $false)
                $objectResults += [ordered]@{kind='report-output';status='exported';export=[IO.Path]::GetFileName($pdf);exportSha256=(Get-FileHash -LiteralPath $pdf).Hash.ToLowerInvariant()}
            } catch { $objectResults += [ordered]@{kind='report-output';status='failed';error=$_.Exception.Message} }
        } finally { $application.CloseCurrentDatabase() }
        $results += [ordered]@{path=$profile.Name;sha256=(Get-FileHash -LiteralPath $path).Hash.ToLowerInvariant();objects=$objectResults;limitations=@('Independent producer only; no OfficeIMO native object codec proof', 'Inert StopMacro only; no general action-macro semantics or execution proof', 'VBA signature and password variants not exercised')}
    }
} finally {
    if ($application) { $application.Quit(2); [Runtime.InteropServices.Marshal]::FinalReleaseComObject($application) | Out-Null }
}
$appPath = (Get-ItemProperty -LiteralPath 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\MSACCESS.EXE' -ErrorAction SilentlyContinue).'(default)'
$producerBuild = if ($appPath -and (Test-Path -LiteralPath $appPath)) { (Get-Item -LiteralPath $appPath).VersionInfo.FileVersion } else { 'not available' }
[ordered]@{schemaVersion=1;producer='Microsoft Access';producerBuild=$producerBuild;security='Owned new instance, ForceDisable=3, new synthetic database only, no startup or event code';license='MIT (synthetic OfficeIMO fixtures)';files=$results} | ConvertTo-Json -Depth 12 | Set-Content -LiteralPath (Join-Path $root 'manifest.json') -Encoding utf8
$results | ConvertTo-Json -Depth 12

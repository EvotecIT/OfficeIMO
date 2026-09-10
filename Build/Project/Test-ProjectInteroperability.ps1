#requires -Version 5.1
<#
.SYNOPSIS
Opens a Project/XML artifact in installed Microsoft Project and records semantic readback.
.DESCRIPTION
Run with Windows PowerShell 5.1. This opt-in interoperability oracle uses a new
application instance, disables macros, and closes without updating the input.
.EXAMPLE
powershell.exe -NoProfile -File Build/Project/Test-ProjectInteroperability.ps1 -InputPath ./authored.xml -OutputPath ./artifacts/project/readback
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)][string] $InputPath,
    [Parameter(Mandatory)][string] $OutputPath,
    [string] $ExpectedPath
)
$ErrorActionPreference = 'Stop'
$inputFile = [IO.Path]::GetFullPath($InputPath)
$outputDirectory = [IO.Path]::GetFullPath($OutputPath)
if (Test-Path -LiteralPath $outputDirectory) { throw 'Choose a new output directory.' }
[void][IO.Directory]::CreateDirectory($outputDirectory)
$missing = [Type]::Missing
$app = New-Object -ComObject MSProject.Application
try {
    $app.Visible = $false
    $app.DisplayAlerts = $false
    $app.AutomationSecurity = 3
    if (!$app.FileOpenEx($inputFile, $true, 0)) { throw 'Microsoft Project did not open the input.' }
    $project = $app.ActiveProject
    $tasks = @($project.Tasks | Where-Object { $null -ne $_ } | ForEach-Object {
        $task = $_
        [ordered]@{
            uid = $task.UniqueID; id = $task.ID; name = $task.Name
            outlineLevel = $task.OutlineLevel; summary = $task.Summary
            durationMinutes = $task.Duration; workMinutes = $task.Work; cost = $task.Cost
            start = [string]$task.Start; finish = [string]$task.Finish
            predecessors = [string]$task.Predecessors; notes = [string]$task.Notes
            text1 = [string]$task.Text1; baselineCost = $task.BaselineCost
        }
    })
    $resources = @($project.Resources | Where-Object { $null -ne $_ } | ForEach-Object {
        [ordered]@{ uid = $_.UniqueID; name = $_.Name; type = [int]$_.Type; standardRate = [string]$_.StandardRate; workMinutes = $_.Work; cost = $_.Cost }
    })
    $calendars = @($project.BaseCalendars | ForEach-Object {
        [ordered]@{ name = $_.Name; exceptions = @($_.Exceptions | ForEach-Object {
            [ordered]@{ name = $_.Name; start = ([datetime]$_.Start).ToString('yyyy-MM-dd'); finish = ([datetime]$_.Finish).ToString('yyyy-MM-dd') }
        }) }
    })
    $assignments = @($project.Tasks | Where-Object { $null -ne $_ } | ForEach-Object { $_.Assignments } | Where-Object { $null -ne $_ } | ForEach-Object {
        [ordered]@{ uid = $_.UniqueID; taskUid = $_.TaskUniqueID; resourceUid = $_.ResourceUniqueID; units = $_.Units; cost = $_.Cost; workMinutes = $_.Work }
    })
    $reexport = Join-Path $outputDirectory 'reexport.xml'
    [void]$app.FileSaveAs($reexport, 0, $missing, $missing, $missing, $missing, $missing, $missing, $missing, 'MSProject.xml')
    $result = [ordered]@{
        producer = 'Microsoft Project'; version = $app.Version
        build = (Get-Item (Join-Path $app.Path 'WINPROJ.EXE')).VersionInfo.FileVersion
        inputSha256 = (Get-FileHash -LiteralPath $inputFile -Algorithm SHA256).Hash.ToLowerInvariant()
        tasks = $tasks; resources = $resources; calendars = $calendars; assignments = $assignments
        reexportSha256 = (Get-FileHash -LiteralPath $reexport -Algorithm SHA256).Hash.ToLowerInvariant()
    }
    $result | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath (Join-Path $outputDirectory 'readback.json') -Encoding UTF8
    if ($ExpectedPath) {
        # An expected file contains selected records and properties; omitted fields remain observable in readback.json.
        $expected = Get-Content -LiteralPath $ExpectedPath -Raw | ConvertFrom-Json
        foreach ($section in @('tasks', 'resources', 'calendars', 'assignments')) {
            foreach ($record in @($expected.$section | Where-Object { $null -ne $_ })) {
                $key = if ($record.PSObject.Properties['uid']) { 'uid' } else { 'name' }
                $actual = @($result[$section] | Where-Object { $_[$key] -eq $record.$key })
                if ($actual.Count -ne 1) { throw "Expected exactly one $section record with $key=$($record.$key)." }
                foreach ($property in $record.PSObject.Properties) {
                    $actualJson = ConvertTo-Json -InputObject $actual[0][$property.Name] -Depth 6 -Compress
                    $expectedJson = ConvertTo-Json -InputObject $property.Value -Depth 6 -Compress
                    if ($actualJson -cne $expectedJson) { throw "Readback differs at $section/$($record.$key)/$($property.Name): expected $expectedJson, got $actualJson." }
                }
            }
        }
    }
    $result | ConvertTo-Json -Depth 6
    [void]$app.FileCloseEx(0)
} finally {
    [void]$app.Quit(0)
    [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)
}

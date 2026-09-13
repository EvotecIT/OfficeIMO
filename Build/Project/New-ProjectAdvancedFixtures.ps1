#requires -Version 5.1
<#
.SYNOPSIS
Creates synthetic advanced scheduling fixtures with an installed Microsoft Project.
.DESCRIPTION
Run in Windows PowerShell 5.1 for the registered Office interop assembly. Each case
uses a new document in a separate application instance. Output contains XML, native
files and object-model observations; the OfficeIMO runtime does not automate Office.
.EXAMPLE
powershell.exe -NoProfile -File Build/Project/New-ProjectAdvancedFixtures.ps1 -OutputPath ./artifacts/project/advanced
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)][string] $OutputPath,
    [ValidateSet('rates', 'calendars', 'effort', 'progress', 'contours', 'leveling', 'fields', 'rollups')]
    [string[]] $ScenarioNames = @('rates', 'calendars', 'effort', 'progress', 'contours', 'leveling', 'fields', 'rollups')
)
$ErrorActionPreference = 'Stop'
$destination = [IO.Path]::GetFullPath($OutputPath)
if (Test-Path -LiteralPath $destination) { throw 'Choose a new output directory.' }
[void][IO.Directory]::CreateDirectory($destination)
$missing = [Type]::Missing
. (Join-Path $PSScriptRoot 'Set-ProjectCustomFieldFixture.ps1')
$app = New-Object -ComObject MSProject.Application

function Save-ProjectCase {
    param([string] $Name)
    [void]$app.CalculateProject()
    $project = $app.ActiveProject
    $tasks = foreach ($task in $project.Tasks) {
        if ($null -eq $task) { continue }
        [pscustomobject]@{
            Uid = $task.UniqueID; Name = $task.Name; Type = [int]$task.Type; EffortDriven = [bool]$task.EffortDriven
            Start = ([datetime]$task.Start).ToString('s'); Finish = ([datetime]$task.Finish).ToString('s')
            Duration = $task.Duration; Work = $task.Work; ActualWork = $task.ActualWork; RemainingWork = $task.RemainingWork
            ActualDuration = $task.ActualDuration; RemainingDuration = $task.RemainingDuration
            PercentComplete = $task.PercentComplete; PercentWorkComplete = $task.PercentWorkComplete
            PhysicalPercentComplete = $task.PhysicalPercentComplete; Cost = $task.Cost; ActualCost = $task.ActualCost
            CustomValues = if ($Name -in @('fields', 'rollups')) {
                $custom = [ordered]@{}
                foreach ($field in @('Number1', 'Number2', 'Number3', 'Number8', 'Number9', 'Number10', 'Number11', 'Number12', 'Number13', 'Number14', 'Number15', 'Text1', 'Text2', 'Flag1', 'Flag2', 'Flag3', 'Date1')) { $custom[$field] = [string]$task.$field }
                $custom
            } else { $null }
        }
    }
    $assignments = foreach ($task in $project.Tasks) { foreach ($assignment in $task.Assignments) {
        [pscustomobject]@{
            Uid = $assignment.UniqueID; TaskUid = $assignment.Task.UniqueID; ResourceUid = $assignment.Resource.UniqueID
            Start = ([datetime]$assignment.Start).ToString('s'); Finish = ([datetime]$assignment.Finish).ToString('s')
            Units = $assignment.Units; Work = $assignment.Work; ActualWork = $assignment.ActualWork
            RemainingWork = $assignment.RemainingWork; OvertimeWork = $assignment.OvertimeWork
            Cost = $assignment.Cost; ActualCost = $assignment.ActualCost; Delay = $assignment.Delay
            WorkContour = [int]$assignment.WorkContour
        }
    } }
    $snapshot = [pscustomobject]@{ Tasks = @($tasks); Assignments = @($assignments) }
    [IO.File]::WriteAllText((Join-Path $destination ($Name + '.json')), ($snapshot | ConvertTo-Json -Depth 8), [Text.UTF8Encoding]::new($false))
    [void]$app.FileSaveAs((Join-Path $destination ($Name + '.mpp')), 0, $missing, $missing, $missing, $missing, $missing, $missing, $missing, 'MSProject.mpp')
    [void]$app.FileSaveAs((Join-Path $destination ($Name + '.xml')), 0, $missing, $missing, $missing, $missing, $missing, $missing, $missing, 'MSProject.xml')
}

try {
    $app.Visible = $false; $app.DisplayAlerts = $false; $app.AutomationSecurity = 3
    foreach ($scenario in $ScenarioNames) {
        [void]$app.FileNew()
        $project = $app.ActiveProject
        $project.ProjectStart = [datetime]'2026-10-05T08:00:00'
        $project.CurrentDate = [datetime]'2026-10-05T08:00:00'
        $first = $project.Resources.Add('Engineer'); $first.StandardRate = 100; $first.OvertimeRate = 150; $first.CostPerUse = 25
        $second = $project.Resources.Add('Reviewer'); $second.StandardRate = 120
        switch ($scenario) {
            'rates' {
                [void]$first.PayRates.Add([datetime]'2026-10-06', '200/h', '300/h', 25)
                $availability = $first.Availabilities.Item(1)
                $availability.AvailableFrom = [datetime]'2026-10-05'; $availability.AvailableTo = [datetime]'2026-10-06'; $availability.AvailableUnit = 100
                [void]$first.Availabilities.Add([datetime]'2026-10-07', [datetime]'2026-12-31', 50)
                $task = $project.Tasks.Add('Dated rates'); $task.Manual = $false; $task.Duration = 960; $task.EffortDriven = $false
                [void]$task.Assignments.Add($task.ID, $first.ID, 1)
                $material = $project.Resources.Add('Material'); $material.Type = 1; $material.MaterialLabel = 'kg'; $material.StandardRate = 2; $material.CostPerUse = 3
                [void]$task.Assignments.Add($task.ID, $material.ID, 5)
                $variable = $project.Tasks.Add('Variable material'); $variable.Manual = $false; $variable.Duration = 960
                $materialAssignment = $variable.Assignments.Add($variable.ID, $material.ID, 2)
                $materialAssignment.Units = '2/h'
            }
            'calendars' {
                [void]$app.BaseCalendarCreate('Morning', 'Standard')
                $calendar = $project.BaseCalendars.Item('Morning')
                foreach ($day in 2..6) { $weekday = $calendar.WeekDays.Item($day); $weekday.Shift2.Clear() }
                $second.BaseCalendar = 'Morning'
                $task = $project.Tasks.Add('Independent calendars'); $task.Manual = $false; $task.Duration = 960; $task.EffortDriven = $false
                [void]$task.Assignments.Add($task.ID, $first.ID, 1)
                [void]$task.Assignments.Add($task.ID, $second.ID, 1)
                $delayed = $project.Tasks.Add('Delayed assignment'); $delayed.Manual = $false; $delayed.Duration = 480
                $assignment = $delayed.Assignments.Add($delayed.ID, $first.ID, 0.5); $assignment.Delay = 120
            }
            'effort' {
                foreach ($type in 0..2) {
                    $task = $project.Tasks.Add("Effort type $type"); $task.Manual = $false; $task.Duration = 960; $task.Type = $type
                    if ($type -lt 2) { $task.EffortDriven = $true }
                    [void]$task.Assignments.Add($task.ID, $first.ID, 1)
                }
                Save-ProjectCase 'effort-before'
                foreach ($task in $project.Tasks) { [void]$task.Assignments.Add($task.ID, $second.ID, 1) }
            }
            'progress' {
                $project.StatusDate = [datetime]'2026-10-07T08:00:00'
                $task = $project.Tasks.Add('Progress and overtime'); $task.Manual = $false; $task.Duration = 960
                $assignment = $task.Assignments.Add($task.ID, $first.ID, 1)
                $assignment.OvertimeWork = 120
                [void]$app.CalculateProject(); [void]$app.BaselineSave($true)
                $assignment.ActualWork = 480
                $task.PhysicalPercentComplete = 25
                $task.Split([datetime]'2026-10-06T08:00:00', [datetime]'2026-10-08T08:00:00')
            }
            'contours' {
                foreach ($contour in 0..7) {
                    $task = $project.Tasks.Add("Contour $contour"); $task.Manual = $false; $task.Duration = 960
                    $assignment = $task.Assignments.Add($task.ID, $first.ID, 1); $assignment.WorkContour = $contour
                }
            }
            'leveling' {
                foreach ($name in @('First', 'Second')) {
                    $task = $project.Tasks.Add($name); $task.Manual = $false; $task.Duration = 480
                    [void]$task.Assignments.Add($task.ID, $first.ID, 1)
                }
                [void]$app.CalculateProject()
                [void]$app.LevelingOptions($false, $false, $true, 0, $true, $missing, $missing, 0, $false, $false, $true)
                [void]$app.LevelNow($true)
            }
            'fields' { Set-ProjectCustomFieldFixture -Application $app -Scenario fields }
            'rollups' { Set-ProjectCustomFieldFixture -Application $app -Scenario rollups }
        }
        Save-ProjectCase $scenario
        [void]$app.FileCloseEx(0)
    }
    $context = [pscustomobject]@{ Producer = 'Microsoft Project'; Version = $app.Version; Build = (Get-Item (Join-Path $app.Path 'WINPROJ.EXE')).VersionInfo.FileVersion; CreatedUtc = [datetime]::UtcNow.ToString('o') }
    [IO.File]::WriteAllText((Join-Path $destination 'context.json'), ($context | ConvertTo-Json), [Text.UTF8Encoding]::new($false))
} catch {
    Write-Error ($_.Exception.Message + [Environment]::NewLine + $_.ScriptStackTrace) -ErrorAction Continue
    throw
} finally {
    try { [void]$app.Quit(0) } finally { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app) }
}

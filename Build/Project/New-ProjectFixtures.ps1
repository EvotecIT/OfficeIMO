#requires -Version 5.1
<#
.SYNOPSIS
Creates synthetic Project/XML fixture pairs using an installed Microsoft Project.
.DESCRIPTION
Run with Windows PowerShell 5.1 so Office's registered interop assembly is available.
Uses a separate application instance and never opens user documents. No application
automation is part of the OfficeIMO.Project runtime package.
.EXAMPLE
powershell.exe -NoProfile -File Build/Project/New-ProjectFixtures.ps1 -OutputPath ./artifacts/project/oracle
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)][string] $OutputPath,
    [ValidateSet('empty', 'delivery', 'calendars', 'actuals', 'relationships', 'resources', 'custom-fields')]
    [string[]] $ScenarioNames = @('empty', 'delivery', 'calendars', 'actuals', 'relationships', 'resources', 'custom-fields')
)

$ErrorActionPreference = 'Stop'
$destination = [IO.Path]::GetFullPath($OutputPath)
if (Test-Path -LiteralPath $destination) {
    throw 'Choose a new output directory so existing evidence is not overwritten.'
}
[void][IO.Directory]::CreateDirectory($destination)
$missing = [Type]::Missing
$app = New-Object -ComObject MSProject.Application
try {
    $app.Visible = $false
    $app.DisplayAlerts = $false
    $app.AutomationSecurity = 3
    $producerVersion = $app.Version
    $producerBuild = (Get-Item (Join-Path $app.Path 'WINPROJ.EXE')).VersionInfo.FileVersion
    foreach ($scenario in $ScenarioNames) {
        [void]$app.FileNew()
        $project = $app.ActiveProject
        $project.ProjectStart = [datetime]'2026-10-05T08:00:00'
        [void]$app.ProjectSummaryInfo($missing, "OfficeIMO $scenario", 'Synthetic interoperability fixture', 'OfficeIMO', 'OfficeIMO', 'OfficeIMO', 'Created entirely for OfficeIMO format verification.')
        if ($scenario -ne 'empty') {
            $summary = $project.Tasks.Add('Delivery')
            $design = $project.Tasks.Add('Design')
            $design.OutlineLevel = 2
            $design.Manual = $false
            $design.Duration = 3 * 480
            $build = $project.Tasks.Add('Build')
            $build.OutlineLevel = 2
            $build.Manual = $false
            $build.Duration = 5 * 480
            $build.Predecessors = [string]$design.ID
            $build.Text1 = 'Platform'
            $build.Notes = 'Synthetic notes: caf' + [char]0x00e9 + ' / ' + [char]0x0141 + [char]0x00f3 + 'd' + [char]0x017a + ' / ' + [char]0x65e5 + [char]0x672c + [char]0x8a9e
            $resource = $project.Resources.Add('Engineer')
            $resource.StandardRate = 125
            [void]$build.Assignments.Add($build.ID, $resource.ID, 1)
            $milestone = $project.Tasks.Add('Release')
            $milestone.OutlineLevel = 2
            $milestone.Manual = $false
            $milestone.Duration = 0
            $milestone.Predecessors = [string]$build.ID
            [void]$app.CalculateProject()
            [void]$app.BaselineSave($true)
            if ($scenario -eq 'calendars') {
                [void]$app.BaseCalendarCreate('Workshop', 'Standard')
                $calendar = $project.BaseCalendars.Item('Workshop')
                $calendar.WeekDays.Item(6).Working = $false
                [void]$calendar.Exceptions.Add(1, [datetime]'2026-10-12', [datetime]'2026-10-12', 1, 'Maintenance')
                $build.Calendar = 'Workshop'
                $resource.BaseCalendar = 'Workshop'
                [void]$app.CalculateProject()
            }
            if ($scenario -eq 'actuals') {
                $design.PercentComplete = 100
                $build.PercentComplete = 40
                $build.FixedCost = 120.50
                $resource.CostPerUse = 25.75
                $build.Deadline = [datetime]'2026-10-30T17:00:00'
                $project.StatusDate = [datetime]'2026-10-09T17:00:00'
                [void]$app.CalculateProject()
            }
            if ($scenario -eq 'relationships') {
                foreach ($kind in @('FS', 'SS', 'FF', 'SF')) {
                    $linked = $project.Tasks.Add("Dependency $kind")
                    $linked.OutlineLevel = 2
                    $linked.Manual = $false
                    $linked.Duration = 480
                    $linked.Predecessors = [string]$design.ID + $kind + '+2h'
                }
                $lead = $project.Tasks.Add('Negative lag')
                $lead.OutlineLevel = 2
                $lead.Manual = $false
                $lead.Duration = 480
                $lead.Predecessors = [string]$design.ID + 'FS-1d'
                $percent = $project.Tasks.Add('Percent lag')
                $percent.OutlineLevel = 2
                $percent.Manual = $false
                $percent.Duration = 480
                $percent.Predecessors = [string]$design.ID + 'FS+50%'
                [void]$app.CalculateProject()
            }
            if ($scenario -eq 'resources') {
                $material = $project.Resources.Add('Steel')
                $material.Type = 1
                $material.MaterialLabel = 'kg'
                $material.StandardRate = 2
                [void]$build.Assignments.Add($build.ID, $material.ID, 3)
                $costResource = $project.Resources.Add('Travel')
                $costResource.Type = 2
                $costAssignment = $build.Assignments.Add($build.ID, $costResource.ID)
                $costAssignment.Cost = 300
                [void]$app.CalculateProject()
            }
            if ($scenario -eq 'custom-fields') {
                [void]$app.CustomFieldRename(188743731, 'Area')
                [void]$app.CustomFieldValueListAdd(188743731, 'Platform', 'Shared platform work')
                [void]$app.CustomFieldValueListAdd(188743731, 'UX', 'User experience work')
                $build.Text1 = 'Platform'
                $design.Text1 = 'UX'
                [void]$app.CalculateProject()
            }
        }
        $native = Join-Path $destination ($scenario + '.mpp')
        $xml = Join-Path $destination ($scenario + '.xml')
        [void]$app.FileSaveAs($native, 0, $missing, $missing, $missing, $missing, $missing, $missing, $missing, 'MSProject.mpp')
        [void]$app.FileSaveAs($xml, 0, $missing, $missing, $missing, $missing, $missing, $missing, $missing, 'MSProject.xml')
        [void]$app.FileCloseEx(0)
    }
    $files = @(Get-ChildItem -LiteralPath $destination -File | ForEach-Object {
        [ordered]@{ file = $_.Name; sha256 = (Get-FileHash -LiteralPath $_.FullName -Algorithm SHA256).Hash.ToLowerInvariant(); bytes = $_.Length }
    })
    [ordered]@{
        producer = 'Microsoft Project'; version = $producerVersion; build = $producerBuild
        locale = [Globalization.CultureInfo]::CurrentCulture.Name
        provenance = 'Synthetic schedules created by Build/Project/New-ProjectFixtures.ps1; no customer data or third-party fixture content.'
        redistribution = 'OfficeIMO-authored synthetic content; MIT repository license. Microsoft Project is the producer and is not redistributed.'
        formatFamily = 'MPP14 and Project XML SaveVersion 14'
        files = $files
    } | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path $destination 'manifest.json') -Encoding UTF8
    $files
} finally {
    [void]$app.Quit(0)
    [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)
}

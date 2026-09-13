#requires -Version 5.1
<#
.SYNOPSIS
Creates a two-level outline-code fixture with an installed Microsoft Project.
.EXAMPLE
powershell.exe -NoProfile -File Build/Project/New-ProjectOutlineFixture.ps1 -OutputPath ./artifacts/project/outline
#>
[CmdletBinding()]
param([Parameter(Mandatory)][string] $OutputPath)
$ErrorActionPreference = 'Stop'
$destination = [IO.Path]::GetFullPath($OutputPath)
if (Test-Path -LiteralPath $destination) { throw 'Output already exists.' }
[void][IO.Directory]::CreateDirectory($destination)
$app = New-Object -ComObject MSProject.Application
try {
    $app.Visible = $false; $app.DisplayAlerts = $false; $app.AutomationSecurity = 3
    [void]$app.FileNew(); $project = $app.ActiveProject
    $project.ProjectStart = [datetime]'2026-10-05T08:00:00'
    $code = $project.OutlineCodes.Add([Microsoft.Office.Interop.MSProject.PjCustomField]::pjCustomTaskOutlineCode1, 'Discipline')
    [void]$code.CodeMask.Add(3, 0, '.')
    [void]$code.CodeMask.Add(0, 2, '-')
    $root = $code.LookupTable.AddChild('Design', [Type]::Missing)
    $child = $code.LookupTable.AddChild('01', $root.UniqueID)
    $child.Description = 'Initial design'; $code.OnlyCompleteCodes = $true; $code.OnlyLookUpTableCodes = $true
    $task = $project.Tasks.Add('Concept'); $task.Manual = $false; $task.Duration = 480; $task.OutlineCode1 = 'Design.01'
    [void]$app.CalculateProject()
    [void]$app.FileSaveAs((Join-Path $destination 'outline.mpp'), 0, $false, [Type]::Missing, [Type]::Missing, [Type]::Missing, [Type]::Missing, [Type]::Missing, [Type]::Missing, 'MSProject.MPP')
    [void]$app.FileSaveAs((Join-Path $destination 'outline.xml'), 0, $false, [Type]::Missing, [Type]::Missing, [Type]::Missing, [Type]::Missing, [Type]::Missing, [Type]::Missing, 'MSProject.XML')
    [pscustomobject]@{ FieldId = [int]$code.FieldID; Root = $root.UniqueID; Child = $child.UniqueID; Value = $task.OutlineCode1 } | ConvertTo-Json | Set-Content (Join-Path $destination 'observations.json') -Encoding UTF8
} finally { $app.Quit(0) }

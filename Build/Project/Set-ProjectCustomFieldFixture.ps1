function Set-ProjectCustomFieldFixture {
    [CmdletBinding()]
    param([Parameter(Mandatory)] $Application, [Parameter(Mandatory)][ValidateSet('fields', 'rollups')][string] $Scenario)
    $project = $Application.ActiveProject
    $missing = [Type]::Missing
    $fieldType = [Microsoft.Office.Interop.MSProject.PjCustomField]
    if ($Scenario -eq 'fields') {
        $project.ProjectSummaryTask.Number5 = 480; $project.ProjectSummaryTask.Number6 = 1
        $task = $project.Tasks.Add('Design'); $task.Manual = $false; $task.Duration = 960
        $task.Number2 = 3; $task.Number4 = 2; $task.Number5 = 480; $task.Number6 = 1; $task.Number7 = 0
        [void]$task.Assignments.Add($task.ID, $project.Resources.Item(1).ID, 1)
        $formulas = [ordered]@{
            pjCustomTaskNumber1 = '[Number2] * 2 + [Cost] / 100'
            # Field references avoid a numeric-literal/separator corruption in this producer's XML exporter.
            pjCustomTaskNumber3 = 'IIf([Number2] > [Number4], Round([Duration] / [Number5], [Number6]), [Number7])'
            pjCustomTaskText1 = 'UCase([Name]) & ":" & CStr([Number1])'
            pjCustomTaskFlag1 = '[Number2] > 2 And [% Complete] < 100'
            pjCustomTaskDate1 = 'DateAdd("d", 2, [Start])'
            pjCustomTaskNumber16 = '2 ^ 3 ^ 2'
            pjCustomTaskNumber17 = 'IIf(True Xor True Or True, [Number6], [Number7])'
            pjCustomResourceNumber1 = '[Cost] / 100'
            pjCustomResourceText1 = 'UCase([Name])'
        }
        foreach ($entry in $formulas.GetEnumerator()) {
            $field = [int][Enum]::Parse($fieldType, $entry.Key)
            # The installed producer accepts semicolon-separated local VBA function arguments.
            [void]$Application.CustomFieldSetFormula($field, $entry.Value.Replace(',', ';'))
            [void]$Application.CustomFieldPropertiesEx($field, 1, 11, $missing, $missing, $missing)
        }
        $lookupField = [int][Enum]::Parse($fieldType, 'pjCustomTaskText2')
        [void]$Application.CustomFieldValueListAdd($lookupField, 'Ready', 'Work can begin')
        [void]$Application.CustomFieldValueListAdd($lookupField, 'Waiting', 'Needs input')
        [void]$Application.CustomFieldPropertiesEx($lookupField, 2, 10, $missing, $missing, $missing)
        $task.Text2 = 'Ready'
        [void]$Application.CustomFieldIndicatorAdd([int][Enum]::Parse($fieldType, 'pjCustomTaskNumber1'),
            [int][Microsoft.Office.Interop.MSProject.PjComparison]::pjCompareGreaterThan, '10', [int][Microsoft.Office.Interop.MSProject.PjIndicator]::pjIndicatorSphereGreen)
    } else {
        $summary = $project.Tasks.Add('Phase'); $summary.Manual = $false
        $first = $project.Tasks.Add('First'); $first.Manual = $false; $first.Duration = 480; $first.OutlineLevel = 2
        $nested = $project.Tasks.Add('Nested phase'); $nested.Manual = $false; $nested.OutlineLevel = 2
        $second = $project.Tasks.Add('Second'); $second.Manual = $false; $second.Duration = 480; $second.OutlineLevel = 3
        $third = $project.Tasks.Add('Third'); $third.Manual = $false; $third.Duration = 480; $third.OutlineLevel = 3
        $fourth = $project.Tasks.Add('Fourth'); $fourth.Manual = $false; $fourth.Duration = 480; $fourth.OutlineLevel = 1
        $rows = @($first, $second, $third, $fourth)
        $amounts = @(2, 4, 6, 10)
        for ($index = 0; $index -lt $rows.Count; $index++) {
            foreach ($number in 8..15) { $rows[$index].('Number' + $number) = $amounts[$index] }
            $rows[$index].Flag2 = $index -eq 1; $rows[$index].Flag3 = $index -ne 1
        }
        $rollups = [ordered]@{ Number8 = 3; Number9 = 4; Number10 = 5; Number11 = 2; Number12 = 6; Number13 = 7; Number14 = 0; Number15 = 1; Flag2 = 0; Flag3 = 1 }
        foreach ($entry in $rollups.GetEnumerator()) {
            $field = [int][Enum]::Parse($fieldType, 'pjCustomTask' + $entry.Key)
            [void]$Application.CustomFieldPropertiesEx($field, 0, $entry.Value, $missing, $missing, $missing)
        }
    }
}

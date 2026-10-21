using System.Management.Automation;
using System.Management.Automation.Runspaces;
using OfficeIMO.Project;
using PSWriteOffice.Cmdlets.Project;

namespace PSWriteOffice.Tests;

[Collection(PowerShellRunspaceCollection.Name)]
public sealed class ProjectCmdletTests {
    [Theory]
    [InlineData(".xml")]
    [InlineData(".mpp")]
    public void ProjectPipelineSchedulesSavesAndReadsTasksAndAssignments(string extension) {
        string directory = Path.Combine(Path.GetTempPath(), "officeimo-project-cmdlets-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        try {
            var state = InitialSessionState.CreateDefault();
            state.ImportPSModule(new[] { typeof(NewOfficeProjectCommand).Assembly.Location });
            using var runspace = RunspaceFactory.CreateRunspace(state);
            runspace.Open();
            runspace.SessionStateProxy.SetVariable("outputPath", Path.Combine(directory, "delivery" + extension));
            using var shell = PowerShell.Create();
            shell.Runspace = runspace;
            shell.AddScript("""
                $project = New-OfficeProject -Name Delivery -StartDate '2026-10-20 08:00'
                try {
                    $design = $project | Add-OfficeProjectTask -Name Design -DurationDays 3 -PassThru
                    $build = $project | Add-OfficeProjectTask -Name Build -DurationDays 5 -PassThru
                    $engineer = $project | Add-OfficeProjectResource -Name Engineer -StandardRate 125 -PassThru
                    $project | Add-OfficeProjectDependency -PredecessorUid $design.Uid -SuccessorUid $build.Uid
                    $project | Add-OfficeProjectAssignment -TaskUid $build.Uid -ResourceUid $engineer.Uid -UnitsPercent 50
                    $schedule = $project | Invoke-OfficeProjectSchedule -CalculateAssignments -Apply
                    $project | Save-OfficeProject -Path $outputPath -WhatIf
                    if (Test-Path -LiteralPath $outputPath) { throw 'WhatIf wrote the project' }
                    $acceptNativeScheduleCacheLoss = [IO.Path]::GetExtension($outputPath) -eq '.mpp'
                    if ($acceptNativeScheduleCacheLoss) {
                        try {
                            $project | Save-OfficeProject -Path $outputPath -ReportVariable blockedReport -ErrorAction Stop
                            throw 'Strict native save unexpectedly accepted schedule cache loss'
                        } catch {
                            if (-not $blockedReport.HasLoss) { throw }
                        }
                        if (Test-Path -LiteralPath $outputPath) { throw 'Rejected save wrote a destination' }
                    }
                    $project | Save-OfficeProject -Path $outputPath -AllowLoss:$acceptNativeScheduleCacheLoss -ReportVariable saveReport
                    Get-Item -LiteralPath $outputPath | Get-OfficeProject -ReadOnly
                } finally {
                    $project | Close-OfficeProject
                }
                """);
            var output = shell.Invoke();
            Assert.Empty(shell.Streams.Error);
            using var saved = Assert.IsType<ProjectDocument>(Assert.Single(output).BaseObject);
            Assert.Equal(new[] { "Design", "Build" }, saved.AllTasks.Select(task => task.Name));
            Assert.Equal("Engineer", Assert.Single(saved.Resources, resource => resource.Uid != 0).Name);
            Assert.Single(saved.Assignments);
            Assert.Single(saved.Dependencies);
            Assert.All(saved.AllTasks, task => Assert.NotNull(task.Start));
        } finally {
            Directory.Delete(directory, true);
        }
    }
}

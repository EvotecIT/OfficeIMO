#region QuickStart
using OfficeIMO.Project;

using var project = ProjectDocument.Create();
project.Name = "Delivery";
project.Settings.StartDate = new DateTime(2026, 10, 5, 8, 0, 0);
project.Calendar = project.Calendars.AddStandardWorkingWeek();

var engineer = project.Resources.AddWork("Engineer");
var phase = project.Tasks.AddSummary("Delivery");
var design = phase.Children.Add("Design");
design.Duration = ProjectDuration.WorkingDays(3);
var build = phase.Children.Add("Build");
build.Duration = ProjectDuration.WorkingDays(5);
project.Dependencies.Add(design, build);
project.Assignments.Add(build, engineer, ProjectUnits.Percent(50));
project.Recalculate(new ProjectScheduleOptions { CalculateAssignments = true });
project.Validate().ThrowIfErrors();
project.Save("Delivery.xml");
#endregion

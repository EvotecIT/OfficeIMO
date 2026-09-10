using OfficeIMO.Project;
using OfficeIMO;

internal static class EditCases {
    internal static int Run(string fixtures, string output) {
        if (Directory.Exists(output)) throw new IOException("Choose a new output directory.");
        Directory.CreateDirectory(output);
        using (var project = ProjectDocument.Load(Path.Combine(fixtures, "delivery.xml"))) {
            var task = project.Tasks.Add("New root"); task.IsManual = false; task.Duration = ProjectDuration.WorkingDays(2);
            project.Save(Path.Combine(output, "root-task-added.xml"), new ProjectSaveOptions { LossPolicy = OfficeConversionLossPolicy.Allow });
            File.WriteAllText(Path.Combine(output, "root-task-added.expected.json"), "{\"tasks\":[{\"uid\":" + task.Uid + ",\"name\":\"New root\",\"outlineLevel\":1,\"durationMinutes\":960}]}");
        }
        using (var project = ProjectDocument.Load(Path.Combine(fixtures, "calendars.xml"))) {
            var calendar = project.Calendars.Single(c => c.Name == "Workshop");
            var exception = calendar.Exceptions.Single();
            exception.FromDate = new DateTime(2026, 10, 13);
            exception.ToDate = new DateTime(2026, 10, 13);
            project.Save(Path.Combine(output, "calendar-date.xml"));
            calendar.Exceptions.Remove(exception);
            project.Save(Path.Combine(output, "calendar-remove.xml"), new ProjectSaveOptions { LossPolicy = OfficeConversionLossPolicy.Allow });
        }
        using (var project = ProjectDocument.Create()) {
            project.Settings.StartDate = new DateTime(2026, 10, 5, 8, 0, 0);
            project.Calendar = project.Calendars.AddStandardWorkingWeek();
            var task = project.Tasks.Add("Build"); task.IsManual = false; task.Duration = ProjectDuration.WorkingDays(1);
            task.Start = new DateTime(2026, 10, 5, 8, 0, 0); task.Finish = new DateTime(2026, 10, 5, 17, 0, 0);
            task.RemainingDuration = task.Duration;
            var steel = project.Resources.AddMaterial("Steel"); steel.MaterialLabel = "kg"; steel.StandardRate = 2;
            var travel = project.Resources.AddCost("Travel");
            var material = project.Assignments.Add(task, steel, ProjectUnits.Fraction(3));
            material.Work = ProjectWork.Hours(3); material.RemainingWork = material.Work;
            var cost = project.Assignments.Add(task, travel); cost.Cost = 300; cost.RemainingCost = 300;
            cost.Start = task.Start; cost.Finish = task.Finish;
            project.Save(Path.Combine(output, "resources-authored.xml"));
        }
        File.WriteAllText(Path.Combine(output, "calendar-date.expected.json"), "{\"calendars\":[{\"name\":\"Workshop\",\"exceptions\":[{\"name\":\"Maintenance\",\"start\":\"2026-10-13\",\"finish\":\"2026-10-13\"}]}]}");
        File.WriteAllText(Path.Combine(output, "calendar-remove.expected.json"), "{\"calendars\":[{\"name\":\"Workshop\",\"exceptions\":[]}]}");
        File.WriteAllText(Path.Combine(output, "resources-authored.expected.json"), "{\"resources\":[{\"name\":\"Steel\",\"type\":1},{\"name\":\"Travel\",\"type\":2}]}");
        return 0;
    }
}

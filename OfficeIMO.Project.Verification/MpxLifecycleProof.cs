using OfficeIMO;
using OfficeIMO.Project;
using System.Text.Json;

internal static class MpxLifecycleProof {
    internal static int Run(string input, string output) {
        if (Directory.Exists(output)) throw new IOException("Choose a new output directory.");
        Directory.CreateDirectory(output);
        var edits = new (string Name, Action<ProjectDocument> Edit)[] {
            ("unchanged", _ => { }),
            ("field-growth", d => { var task = d.AllTasks.First(t => t.Uid > 0); task.Name = "Edited café, with a quoted \"label\""; task.Notes = "First line\nSecond line"; }),
            ("add-entities", d => { var task = d.Tasks.Add("Handover"); task.Duration = ProjectDuration.WorkingDays(1); task.Start = new DateTime(2026, 10, 13, 8, 0, 0); task.Finish = task.Start.Value.AddHours(9); var resource = d.Resources.AddWork("Reviewer"); resource.StandardRate = 50; resource.Calendar = d.Calendar; var assignment = d.Assignments.Add(task, resource); assignment.Work = ProjectWork.Hours(8); assignment.Cost = 400; }),
            ("delete", d => { var task = d.AllTasks.Last(t => t.Uid > 0); (task.Parent?.Children ?? d.Tasks).Remove(task, ProjectRemovalMode.Cascade); }),
            ("reparent", d => { var task = d.AllTasks.FirstOrDefault(t => t.Parent != null); task?.MoveTo(null); }),
            ("calendar", d => { var exception = d.Calendar!.Exceptions.Add(); exception.FromDate = new DateTime(2026, 10, 20); exception.ToDate = exception.FromDate; exception.IsWorking = false; }),
            ("clear", d => { foreach (var exception in d.Calendar!.Exceptions.ToArray()) d.Calendar.Exceptions.Remove(exception); foreach (var task in d.AllTasks) { task.Notes = null; task.Priority = null; } })
        };
        foreach (var edit in edits) {
            using var document = input == "new" ? Create() : ProjectDocument.Load(input);
            edit.Edit(document);
            string path = Path.Combine(output, edit.Name + ".mpx");
            var options = new ProjectSaveOptions { LossPolicy = OfficeConversionLossPolicy.Allow };
            var report = document.AssessSave(path, options); report.ThrowIfErrors(); document.Save(path, options);
            using var reopened = ProjectDocument.Load(path);
            reopened.Validate().ThrowIfErrors();
            File.WriteAllText(Path.Combine(output, edit.Name + "-report.json"), JsonSerializer.Serialize(report.Diagnostics, new JsonSerializerOptions { WriteIndented = true }));
            // Explicit fixture currency for the XML evidence projection, never inferred from an arbitrary input's symbol.
            document.Settings.CurrencyCode = "USD";
            File.WriteAllText(Path.Combine(output, edit.Name + "-model.xml"), document.ToXml(options));
            Console.WriteLine(edit.Name + ": tasks=" + reopened.AllTasks.Count() + " resources=" + reopened.Resources.Count + " assignments=" + reopened.Assignments.Count);
        }
        return 0;
    }
    internal static int Encodings(string output) {
        if (Directory.Exists(output)) throw new IOException("Choose a new output directory.");
        Directory.CreateDirectory(output);
        foreach (var encoding in Enum.GetValues<ProjectMpxEncoding>()) {
            using var document = Create();
            document.Tasks[0].Name = encoding switch { ProjectMpxEncoding.Windows1252 => "Euro €", ProjectMpxEncoding.Dos437 => "Omega Ω", ProjectMpxEncoding.Dos850 => "Latin ø", _ => "Apple " };
            document.Save(Path.Combine(output, encoding + ".mpx"), new ProjectSaveOptions { MpxEncoding = encoding, LossPolicy = OfficeConversionLossPolicy.Allow });
        }
        return 0;
    }
    internal static ProjectDocument Create() {
        var document = ProjectDocument.Create(); document.Name = "MPX authored project";
        document.Settings.StartDate = new DateTime(2026, 10, 5, 8, 0, 0); document.Settings.FinishDate = new DateTime(2026, 10, 9, 17, 0, 0); document.Settings.ScheduleFromStart = true;
        document.Calendar = document.Calendars.AddStandardWorkingWeek();
        var exception = document.Calendar.Exceptions.Add(); exception.FromDate = new DateTime(2026, 10, 12); exception.ToDate = exception.FromDate; exception.IsWorking = false;
        var summary = document.Tasks.AddSummary("Delivery"); var first = summary.Children.Add("Design café"); var second = summary.Children.Add("Build");
        first.Duration = ProjectDuration.WorkingDays(1); first.Start = document.Settings.StartDate; first.Finish = first.Start.Value.AddHours(9); first.Priority = 600; first.Notes = "First line\nSecond line, with quotes: \"design\"";
        second.Duration = ProjectDuration.WorkingDays(2); second.Start = new DateTime(2026, 10, 6, 8, 0, 0); second.Finish = new DateTime(2026, 10, 7, 17, 0, 0);
        document.Dependencies.Add(first, second).Lag = ProjectDuration.WorkingHours(1);
        var resource = document.Resources.AddWork("Engineer"); resource.Calendar = document.Calendar; resource.StandardRate = 100; resource.MaxUnits = ProjectUnits.Percent(100);
        var assignment = document.Assignments.Add(first, resource); assignment.Work = ProjectWork.Hours(8); assignment.ActualWork = ProjectWork.Hours(2); assignment.Cost = 800; assignment.ActualCost = 200; assignment.Start = first.Start; assignment.Finish = first.Finish;
        var baseline = first.Baselines.Add(); baseline.Number = 0; baseline.Start = first.Start; baseline.Finish = first.Finish; baseline.Duration = first.Duration; baseline.Work = ProjectWork.Hours(8); baseline.Cost = 800;
        foreach (var pair in new[] { ("188743731", "Text1", "Work area"), ("188743767", "Number1", "12.5"), ("188743752", "Flag1", "1"), ("188743783", "Duration1", "PT4H0M0S") }) {
            var definition = document.CustomFields.Add(); definition.FieldId = pair.Item1; definition.FieldName = pair.Item2;
            var value = first.CustomFields.Add(); value.FieldId = pair.Item1; value.Value = pair.Item3;
        }
        return document;
    }
}

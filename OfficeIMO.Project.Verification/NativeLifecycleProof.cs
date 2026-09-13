using OfficeIMO;
using OfficeIMO.Project;
using System.Text.Json;

/// <summary>Generation-independent edit and template cases for native application readback.</summary>
internal static class NativeLifecycleProof {
    internal static int Run(string input, string output) {
        if (Directory.Exists(output)) throw new IOException("Choose a new output directory.");
        Directory.CreateDirectory(output);
        void Save(string name, Action<ProjectDocument> edit, string extension = ".mpp") {
            using var document = ProjectDocument.Load(input); edit(document);
            var options = new ProjectSaveOptions { LossPolicy = OfficeConversionLossPolicy.Allow };
            var path = Path.Combine(output, name + extension);
            var report = document.AssessSave(path, options); report.ThrowIfErrors();
            document.Save(path, options);
            using var reopened = ProjectDocument.Load(path);
            if (reopened.NativeInfo!.Generation != document.NativeInfo!.Generation) throw new InvalidDataException("An implicit save changed generation.");
            if (reopened.Calendar!.Exceptions.Count != document.Calendar!.Exceptions.Count)
                throw new InvalidDataException("The calendar exception count changed after the native edit: " + name);
            var expected = new {
                tasks = document.AllTasks.Where(t => t.Uid != 0).Select(t => new { uid = t.Uid, name = t.Name ?? "" }),
                resources = document.Resources.Where(r => r.Uid != 0).Select(r => new { uid = r.Uid, name = r.Name ?? "" }),
                calendars = document.Calendars.Where(c => c.IsBaseCalendar == true).Select(c => new {
                    name = c.Name ?? "", exceptions = c.Exceptions.Select(e => new {
                        name = e.Name ?? "", start = e.FromDate!.Value.ToString("yyyy-MM-dd"), finish = e.ToDate!.Value.ToString("yyyy-MM-dd") })
                })
            };
            File.WriteAllText(Path.Combine(output, name + "-expected.json"), JsonSerializer.Serialize(expected));
            File.WriteAllText(Path.Combine(output, name + "-report.json"), JsonSerializer.Serialize(report.Diagnostics));
            if (extension == ".mpt") {
                using var instance = ProjectDocument.CreateFromTemplate(path);
                instance.Save(Path.Combine(output, "from-template.mpp"), options);
            }
        }
        Save("field-growth", d => {
            var task = d.AllTasks.FirstOrDefault(t => t.Uid > 0);
            if (task != null) task.Name = "Native edit café / Łódź / 日本語 with a longer label";
            d.Title = "Native metadata edit";
        });
        Save("add-entities", d => {
            var predecessor = d.AllTasks.LastOrDefault(t => t.Uid > 0 && !t.IsSummary);
            var task = d.Tasks.Add("Handover"); task.Duration = ProjectDuration.WorkingDays(1);
            task.Start = d.Settings.StartDate; task.Finish = task.Start?.AddHours(9);
            var resource = d.Resources.AddWork("Reviewer"); resource.StandardRate = 50;
            resource.MaxUnits = ProjectUnits.Percent(100); resource.Calendar = d.Calendars.Add("Reviewer", d.Calendar);
            var assignment = d.Assignments.Add(task, resource); assignment.Work = ProjectWork.Hours(8);
            assignment.Units = ProjectUnits.Percent(100); assignment.Start = task.Start; assignment.Finish = task.Finish;
            if (predecessor != null) d.Dependencies.Add(predecessor, task);
        });
        Save("delete-task", d => {
            var task = d.AllTasks.LastOrDefault(t => t.Uid > 0);
            if (task != null) (task.Parent?.Children ?? d.Tasks).Remove(task, ProjectRemovalMode.Cascade);
        });
        Save("reparent", d => {
            var task = d.AllTasks.FirstOrDefault(t => t.Parent != null);
            if (task != null) task.MoveTo(null);
        });
        Save("calendar", d => {
            var day = d.Calendar!.Exceptions.Add(); day.FromDate = d.Settings.StartDate!.Value.Date.AddDays(20);
            day.ToDate = day.FromDate; day.IsWorking = false;
        });
        Save("calendar-clear", d => {
            foreach (var day in d.Calendar!.Exceptions.ToArray()) d.Calendar.Exceptions.Remove(day);
            foreach (var week in d.Calendar.WorkWeeks.ToArray()) d.Calendar.WorkWeeks.Remove(week);
        });
        Save("template", d => d.Title = "Native template proof", ".mpt");
        return 0;
    }
}

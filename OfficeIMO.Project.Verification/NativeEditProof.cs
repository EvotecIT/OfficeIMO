using OfficeIMO;
using OfficeIMO.Project;
using System.Text.Json;

internal static class NativeEditProof {
    internal static int Run(string input, string output) {
        if (Directory.Exists(output)) throw new IOException("Choose a new output directory.");
        Directory.CreateDirectory(output);
        var allow = new ProjectSaveOptions { LossPolicy = OfficeConversionLossPolicy.Allow };
        void Save(string name, Action<ProjectDocument> edit, bool loss = true, string extension = ".mpp") {
            using var document = ProjectDocument.Load(input); edit(document);
            string path = Path.Combine(output, name + extension);
            var report = document.AssessSave(path, loss ? allow : null);
            File.WriteAllText(Path.Combine(output, name + "-report.json"), JsonSerializer.Serialize(report.Diagnostics, new JsonSerializerOptions { WriteIndented = true }));
            document.Save(path, loss ? allow : null);
            using var reload = ProjectDocument.Load(path);
            File.WriteAllText(Path.Combine(output, name + "-model.json"), JsonSerializer.Serialize(new {
                tasks = reload.AllTasks.Select(t => new { t.Uid, t.Name, parent = t.Parent?.Uid, t.DisplayId, duration = t.Duration?.Value, t.Start, t.Finish }),
                resources = reload.Resources.Select(r => new { r.Uid, r.Name, calendar = r.Calendar?.Uid }),
                assignments = reload.Assignments.Select(a => new { a.Uid, task = a.Task?.Uid, resource = a.Resource?.Uid, work = a.Work?.Minutes }),
                dependencies = reload.Dependencies.Select(d => new { predecessor = d.Predecessor?.Uid, successor = d.Successor.Uid })
            }, new JsonSerializerOptions { WriteIndented = true }));
            Console.WriteLine(name + ": " + report.Diagnostics.Count + " diagnostics");
        }
        Save("field-growth", d => { d.Tasks.GetByUid(3).Name = "Build revised café / Łódź / 日本語 with a longer label"; d.Title = "Native metadata updated"; }, false);
        Save("reparent", d => d.Tasks.GetByUid(3).MoveTo(null));
        Save("reparent-xml", d => d.Tasks.GetByUid(3).MoveTo(null), true, ".xml");
        Save("delete-task", d => { var task = d.Tasks.GetByUid(3); (task.Parent?.Children ?? d.Tasks).Remove(task, ProjectRemovalMode.Cascade); });
        Save("add-entities", d => {
            var task = d.Tasks.Add("Handover"); task.Duration = ProjectDuration.WorkingDays(1); task.Start = d.Settings.StartDate; task.Finish = task.Start?.AddHours(9);
            var resource = d.Resources.AddWork("Reviewer"); resource.Calendar = d.Calendars.Add("Reviewer", d.Calendar); resource.StandardRate = 50; resource.MaxUnits = ProjectUnits.Percent(100);
            var assignment = d.Assignments.Add(task, resource); assignment.Units = ProjectUnits.Percent(100); assignment.Work = ProjectWork.Hours(8); assignment.Start = task.Start; assignment.Finish = task.Finish;
            d.Dependencies.Add(d.Tasks.GetByUid(4), task);
        });
        Save("delete-resource", d => {
            var resource = d.Resources.GetByUid(1); var calendar = resource.Calendar; d.Resources.Remove(resource, ProjectRemovalMode.Cascade);
            if (calendar != null) d.Calendars.Remove(calendar);
        });
        Save("calendar", d => {
            var item = d.Calendar!.Exceptions.Add(); item.Name = "Review holiday"; item.FromDate = new DateTime(2026, 10, 7); item.ToDate = item.FromDate; item.IsWorking = false;
        });
        Save("task-calendar", d => { var calendar = d.Calendars.AddStandardWorkingWeek("Task working week"); d.Tasks.GetByUid(3).Calendar = calendar; });
        Save("project-calendar", d => { var calendar = d.Calendars.AddStandardWorkingWeek("Project working week"); d.Calendar = calendar; });
        Save("resource-calendar", d => { var resource = d.Resources.GetByUid(1); var old = resource.Calendar;
            resource.Calendar = d.Calendars.Add("Engineer revised", d.Calendar); d.Calendars.Remove(old!); });
        Save("custom-baseline", d => {
            var task = d.Tasks.GetByUid(3);
            var value = task.CustomFields.FirstOrDefault(f => f.FieldId == "188743731") ?? task.CustomFields.Add(); value.FieldId = "188743731"; value.Value = "Native custom edit";
            task.Baselines.First(b => b.Number == 0).Cost = 4321;
        });
        Save("all-baselines", d => {
            var task = d.Tasks.GetByUid(3); var resource = d.Resources.GetByUid(1); var assignment = d.Assignments.GetByUid(4);
            foreach (var collection in new[] { task.Baselines, resource.Baselines, assignment.Baselines })
                foreach (var baseline in collection.ToArray()) collection.Remove(baseline);
            for (int number = 0; number <= 10; number++) {
                foreach (var collection in new[] { task.Baselines, resource.Baselines, assignment.Baselines }) {
                    var baseline = collection.Add(); baseline.Number = number; baseline.Cost = 1200 + number; baseline.Work = ProjectWork.Hours(10 + number);
                    if (collection != resource.Baselines) { baseline.Start = task.Start; baseline.Finish = task.Finish; }
                    if (collection == task.Baselines) baseline.Duration = task.Duration;
                }
            }
        });
        Save("custom-scalars", d => {
            var task = d.Tasks.GetByUid(3); var resource = d.Resources.GetByUid(1);
            foreach (bool tasks in new[] { true, false }) {
                var fields = tasks ? task.CustomFields : resource.CustomFields;
                foreach (var field in (tasks ? ProjectNativeCustomField.TaskFields : ProjectNativeCustomField.ResourceFields).Where(f => f.Number == 1)) {
                    string id = field.Id.ToString(System.Globalization.CultureInfo.InvariantCulture);
                    var value = fields.FirstOrDefault(f => f.FieldId == id) ?? fields.Add(); value.FieldId = id;
                    value.Value = field.Kind switch { "Text" => "Custom scalar edit", "Number" => "12.5", "Cost" => "12345", "Flag" => "1", "Date" => "2026-10-06T08:00:00", "Duration" => "PT16H0M0S", _ => throw new InvalidOperationException() };
                    if (field.Kind == "Duration" && tasks) value.DurationFormat = 7;
                }
            }
        });
        Save("identities", d => { d.Tasks.GetByUid(1).Guid = Guid.NewGuid(); d.Tasks.GetByUid(3).Guid = Guid.NewGuid(); d.Resources.GetByUid(1).Guid = Guid.NewGuid(); d.Calendar!.Guid = Guid.NewGuid(); });
        Save("template", d => { d.Title = "Template proof"; }, false, ".mpt");
        Save("native-to-xml", d => { d.Tasks.GetByUid(3).Name = "Converted native edit"; }, true, ".xml");
        return 0;
    }
}

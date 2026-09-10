using System.Text.Json;
using OfficeIMO.Project;
using OfficeIMO;

internal static class ScheduleReadback {
    internal static int Create(string input, string output) {
        if (Directory.Exists(output)) throw new IOException("Choose a new output directory.");
        using var document = input == "authored" ? Authored() : ProjectDocument.Load(input);
        var week = document.Calendars.SelectMany(c => c.WorkWeeks).FirstOrDefault();
        if (input != "authored") {
            if (week != null) week.SetWorkingDay(DayOfWeek.Friday, ProjectWorkingTime.Hours(8, 12));
            else document.Tasks.GetByUid(2).Duration = ProjectDuration.WorkingDays(4);
        }
        var schedule = document.Recalculate();
        Directory.CreateDirectory(output);
        // This controlled producer fixture replaces one work-week interval. The default structural
        // opaque-reference warning is retained; application readback validates this explicit allowance.
        document.Save(Path.Combine(output, "calculated.xml"), new ProjectSaveOptions { LossPolicy = OfficeConversionLossPolicy.Allow });
        File.WriteAllText(Path.Combine(output, "expected.json"), JsonSerializer.Serialize(schedule.Tasks, new JsonSerializerOptions { WriteIndented = true }));
        return 0;
    }
    private static ProjectDocument Authored() {
        var document = ProjectDocument.Create(); document.Settings.StartDate = new DateTime(2026, 10, 5, 8, 0, 0); document.Settings.ScheduleFromStart = true;
        document.Calendar = document.Calendars.AddStandardWorkingWeek();
        var calendar = document.Calendars.AddStandardWorkingWeek("Workshop");
        var closure = calendar.Exceptions.Add(); closure.Name = "Shutdown";
        closure.FromDate = new DateTime(2026, 10, 8); closure.ToDate = new DateTime(2026, 10, 9); closure.IsWorking = false;
        var week = calendar.WorkWeeks.Add(); week.Name = "Short Friday";
        week.FromDate = new DateTime(2026, 10, 12); week.ToDate = new DateTime(2026, 10, 16, 23, 59, 0);
        week.SetWorkingDay(DayOfWeek.Friday, ProjectWorkingTime.Hours(8, 12));
        var resource = document.Resources.AddWork("Engineer"); resource.StandardRate = 125; resource.Calendar = calendar;
        var summary = document.Tasks.AddSummary("Delivery");
        var design = summary.Children.Add("Design"); design.Duration = ProjectDuration.WorkingDays(3); design.IsManual = false;
        var build = summary.Children.Add("Build"); build.Duration = ProjectDuration.WorkingDays(5); build.IsManual = false; build.Calendar = calendar;
        document.Assignments.Add(build, resource, ProjectUnits.Fraction(1)).Work = ProjectWork.Hours(40); document.Dependencies.Add(design, build);
        return document;
    }
    internal static int Verify(string expectations, string reexport) {
        using var document = ProjectDocument.Load(reexport);
        using var expected = JsonDocument.Parse(File.ReadAllText(expectations));
        var observations = new List<object>();
        foreach (var record in expected.RootElement.EnumerateArray()) {
            int uid = record.GetProperty("TaskUid").GetInt32();
            var task = document.Tasks.GetByUid(uid);
            var start = record.GetProperty("Start").GetDateTime(); var finish = record.GetProperty("Finish").GetDateTime();
            if (task.Start != start || task.Finish != finish) throw new InvalidDataException("Producer readback date mismatch: task " + uid);
            observations.Add(new { task.Uid, task.Start, task.Finish, task.Duration });
        }
        Console.WriteLine(JsonSerializer.Serialize(new { tasks = observations.Count, observations }));
        return 0;
    }
}

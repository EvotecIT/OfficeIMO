using System.Text.Json;
using OfficeIMO;
using OfficeIMO.Project;

internal static class AdvancedScheduleProof {
    internal static int Authored(string output) {
        if (Directory.Exists(output)) throw new IOException("Choose a new output directory."); Directory.CreateDirectory(output);
        using var document = ProjectDocument.Create(); document.Calendar = document.Calendars.AddStandardWorkingWeek();
        var monday = new DateTime(2026, 10, 5, 8, 0, 0); document.Settings.StartDate = monday; document.Settings.StatusDate = monday.AddHours(4);
        var resource = document.Resources.AddWork("Engineer"); resource.StandardRate = 100; resource.CostPerUse = 25;
        for (int index = 0; index < 2; index++) {
            var task = document.Tasks.Add("Task " + index); task.Duration = ProjectDuration.WorkingDays(1); task.Type = ProjectTaskType.FixedUnits;
            task.IsManual = false; task.EffortDriven = false;
            document.Assignments.Add(task, resource, ProjectUnits.Fraction(1));
        }
        var level = document.CalculateLeveling(); document.ApplyLeveling(level);
        document.CaptureBaseline(document.CalculateSchedule(new ProjectScheduleOptions { CalculateAssignments = true }), 2);
        document.Save(Path.Combine(output, "leveled-baseline.xml"));
        Console.WriteLine(JsonSerializer.Serialize(document.AnalyzeEarnedValue(2)));
        return 0;
    }
    internal static int Create(string source, string output) {
        if (Directory.Exists(output)) throw new IOException("Choose a new output directory.");
        Directory.CreateDirectory(output);
        foreach (var file in Directory.GetFiles(source, "*.xml").OrderBy(f => f, StringComparer.Ordinal)) {
            using var document = ProjectDocument.Load(file);
            var result = document.Recalculate(new ProjectScheduleOptions { CalculateAssignments = true });
            string target = Path.Combine(output, Path.GetFileName(file));
            document.Save(target, new ProjectSaveOptions { LossPolicy = OfficeConversionLossPolicy.Allow });
            Console.WriteLine(JsonSerializer.Serialize(new { name = Path.GetFileName(file), tasks = result.Tasks.Count, assignments = result.Assignments.Count, result.Report.Diagnostics }));
        }
        return 0;
    }
    internal static int Verify(string expectedFile, string actualFile) {
        using var expected = ProjectDocument.Load(expectedFile);
        using var actual = ProjectDocument.Load(actualFile);
        var differences = new List<object>();
        void Compare(string owner, int uid, string field, object? first, object? second) {
            if (first is decimal a && second is decimal b && Math.Abs(a - b) <= .02m) return;
            if (!Equals(first, second)) differences.Add(new { owner, uid, field, expected = first, actual = second });
        }
        foreach (var task in expected.AllTasks.Where(t => t.Uid != 0)) {
            var observed = actual.Tasks.GetByUid(task.Uid);
            Compare("task", task.Uid, "Start", task.Start, observed.Start); Compare("task", task.Uid, "Finish", task.Finish, observed.Finish);
            Compare("task", task.Uid, "Work", task.Work?.Minutes, observed.Work?.Minutes); Compare("task", task.Uid, "ActualWork", task.ActualWork?.Minutes, observed.ActualWork?.Minutes);
            Compare("task", task.Uid, "Cost", task.Cost, observed.Cost); Compare("task", task.Uid, "ActualCost", task.ActualCost, observed.ActualCost);
            Compare("task", task.Uid, "PercentComplete", task.PercentComplete, observed.PercentComplete);
            Compare("task", task.Uid, "PercentWorkComplete", task.PercentWorkComplete, observed.PercentWorkComplete);
            Compare("task", task.Uid, "PhysicalPercentComplete", task.PhysicalPercentComplete ?? 0, observed.PhysicalPercentComplete ?? 0);
            foreach (var baseline in task.Baselines) {
                var actualBaseline = observed.Baselines.SingleOrDefault(b => b.Number == baseline.Number);
                Compare("task", task.Uid, "Baseline" + baseline.Number + "Cost", baseline.Cost, actualBaseline?.Cost);
                Compare("task", task.Uid, "Baseline" + baseline.Number + "Work", baseline.Work?.Minutes, actualBaseline?.Work?.Minutes);
            }
        }
        foreach (var assignment in expected.Assignments) {
            var observed = actual.Assignments.Single(a => a.Task?.Uid == assignment.Task?.Uid && a.Resource?.Uid == assignment.Resource?.Uid);
            Compare("assignment", assignment.Uid, "Start", assignment.Start, observed.Start); Compare("assignment", assignment.Uid, "Finish", assignment.Finish, observed.Finish);
            Compare("assignment", assignment.Uid, "Work", assignment.Work?.Minutes, observed.Work?.Minutes);
            Compare("assignment", assignment.Uid, "ActualWork", assignment.ActualWork?.Minutes, observed.ActualWork?.Minutes);
            Compare("assignment", assignment.Uid, "OvertimeWork", assignment.OvertimeWork?.Minutes, observed.OvertimeWork?.Minutes);
            Compare("assignment", assignment.Uid, "Cost", assignment.Cost, observed.Cost); Compare("assignment", assignment.Uid, "ActualCost", assignment.ActualCost, observed.ActualCost);
        }
        Console.WriteLine(JsonSerializer.Serialize(new { name = Path.GetFileName(expectedFile), differences }, new JsonSerializerOptions { WriteIndented = true }));
        return differences.Count == 0 ? 0 : 1;
    }
}

using OfficeIMO.Project;

internal static class SplitScheduleProof {
    internal static int Create(string output) {
        if (Directory.Exists(output)) throw new IOException("Choose a new output directory.");
        Directory.CreateDirectory(output);
        using var document = ProjectDocument.Create();
        document.Calendar = document.Calendars.AddStandardWorkingWeek();
        var monday = new DateTime(2026, 10, 5, 8, 0, 0); document.Settings.StartDate = monday;
        var resource = document.Resources.AddWork("Engineer"); resource.StandardRate = 100;
        var locked = document.Tasks.Add("Reserved Tuesday"); locked.IsManual = false;
        locked.Duration = ProjectDuration.WorkingDays(1); locked.Type = ProjectTaskType.FixedUnits;
        locked.Priority = 1000; locked.ConstraintType = ProjectConstraintType.MustStartOn; locked.ConstraintDate = monday.AddDays(1);
        document.Assignments.Add(locked, resource, ProjectUnits.Fraction(1));
        var task = document.Tasks.Add("Interrupted work"); task.IsManual = false;
        task.Duration = ProjectDuration.WorkingDays(3); task.Type = ProjectTaskType.FixedUnits;
        document.Assignments.Add(task, resource, ProjectUnits.Fraction(1));
        document.ApplyLeveling(document.CalculateLeveling(new ProjectLevelingOptions { AllowSplitting = true }));
        document.Save(Path.Combine(output, "split.xml"));
        using var recurrence = ProjectDocument.Create(); recurrence.Calendar = recurrence.Calendars.AddStandardWorkingWeek();
        recurrence.Settings.StartDate = monday;
        recurrence.AddRecurringTask("Weekly review", new[] { monday, monday.AddDays(7), monday.AddDays(14) }, ProjectDuration.WorkingDays(1));
        recurrence.Recalculate(new ProjectScheduleOptions { CalculateAssignments = true });
        recurrence.Save(Path.Combine(output, "recurrence.xml"));
        CreateMulti(output, false); CreateMulti(output, true);
        return 0;
    }
    private static void CreateMulti(string output, bool completed) {
        using var document = ProjectDocument.Create(); document.Calendar = document.Calendars.AddStandardWorkingWeek();
        var monday = new DateTime(2026, 10, 5, 8, 0, 0); document.Settings.StartDate = monday;
        var first = document.Resources.AddWork("First"); first.StandardRate = 100;
        var second = document.Resources.AddWork("Second"); second.StandardRate = 100;
        var task = document.Tasks.Add("Delivery"); task.Duration = ProjectDuration.WorkingDays(3); task.IsManual = false;
        task.Type = ProjectTaskType.FixedUnits; task.EffortDriven = false;
        var a = document.Assignments.Add(task, first, ProjectUnits.Fraction(1)); a.Work = ProjectWork.Hours(24);
        var b = document.Assignments.Add(task, second, ProjectUnits.Fraction(1)); b.Work = ProjectWork.Hours(16); b.DelayMinutes = 480;
        if (completed) {
            b.DelayMinutes = 0; b.Work = ProjectWork.Hours(4); b.ActualWork = ProjectWork.Hours(4); b.RemainingWork = ProjectWork.Hours(0);
            b.ActualStart = monday; b.ActualFinish = monday.AddHours(4); b.Stop = b.ActualFinish;
            task.ActualStart = monday; task.ActualDuration = ProjectDuration.WorkingHours(2); task.RemainingDuration = ProjectDuration.WorkingHours(22);
        }
        var locked = document.Tasks.Add("Reserved"); locked.Duration = ProjectDuration.WorkingDays(1); locked.Priority = 1000; locked.IsManual = false;
        locked.ConstraintType = ProjectConstraintType.MustStartOn; locked.ConstraintDate = monday.AddDays(1);
        document.Assignments.Add(locked, completed ? first : second, ProjectUnits.Fraction(1));
        document.ApplyLeveling(document.CalculateLeveling(new ProjectLevelingOptions { AllowSplitting = true }));
        document.Save(Path.Combine(output, completed ? "multi-completed.xml" : "multi-delayed.xml"));
    }
}

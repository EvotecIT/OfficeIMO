namespace OfficeIMO.Project.Tests;

public sealed class ProjectSchedulingTests {
    private static readonly DateTime Monday = new DateTime(2026, 10, 5, 8, 0, 0);
    private static ProjectDocument Standard() {
        var document = ProjectDocument.Create(); document.Settings.StartDate = Monday; document.Settings.ScheduleFromStart = true;
        var calendar = document.Calendars.Add("Standard"); document.Calendar = calendar;
        foreach (DayOfWeek day in Enum.GetValues(typeof(DayOfWeek)))
            calendar.SetWorkingDay(day, day == DayOfWeek.Saturday || day == DayOfWeek.Sunday ? Array.Empty<ProjectWorkingTime>() :
                new[] { ProjectWorkingTime.Hours(8, 12), ProjectWorkingTime.Hours(13, 17) });
        return document;
    }
    [Fact]
    public void CalendarArithmeticTraversesSplitShiftsExceptionsAndInheritedWorkWeeks() {
        using var document = Standard();
        var parent = document.Calendar!;
        var closure = parent.Exceptions.Add(); closure.FromDate = Monday.AddDays(1).Date; closure.ToDate = closure.FromDate; closure.IsWorking = false;
        var child = document.Calendars.Add("Derived", parent);
        var summer = child.WorkWeeks.Add(); summer.FromDate = Monday.Date; summer.ToDate = Monday.AddDays(4).Date;
        summer.SetWorkingDay(DayOfWeek.Wednesday, ProjectWorkingTime.Hours(9, 12));
        Assert.Equal(Monday.AddDays(2).AddHours(4), child.AddWorkingMinutes(Monday, 660));
        Assert.Equal(660, child.WorkingMinutesBetween(Monday, Monday.AddDays(2).AddHours(4)));
        Assert.Equal(Monday, child.AddWorkingMinutes(Monday.AddDays(2).AddHours(4), -660));
        string xml = document.ToXml();
        using var reload = ProjectDocument.Parse(xml);
        Assert.Single(reload.Calendars.GetByUid(child.Uid).WorkWeeks);
        Assert.Equal(660, reload.Calendars.GetByUid(child.Uid).WorkingMinutesBetween(Monday, Monday.AddDays(2).AddHours(4)));
    }
    [Fact]
    public void ReloadedWorkWeeksPermitIntervalAndStructuralEditsWithoutLossPermission() {
        using var document = Standard();
        var week = document.Calendar!.WorkWeeks.Add(); week.FromDate = Monday.Date; week.ToDate = Monday.AddDays(4).Date;
        week.SetWorkingDay(DayOfWeek.Wednesday, ProjectWorkingTime.Hours(9, 12));
        using var reload = ProjectDocument.Parse(document.ToXml());
        reload.Calendar!.WorkWeeks[0].SetWorkingDay(DayOfWeek.Wednesday, ProjectWorkingTime.Hours(10, 12));
        reload.Tasks.Add("Added after reload");
        using var edited = ProjectDocument.Parse(reload.ToXml());
        Assert.Equal("Added after reload", edited.Tasks.Single().Name);
        Assert.Single(edited.Calendar!.WorkWeeks);
        Assert.Equal(120, edited.Calendar.WorkingMinutesBetween(Monday.AddDays(2), Monday.AddDays(2).AddHours(4)));
    }
    [Fact]
    public void OvernightShiftsAndLocalDstBoundaryDoNotDependOnMachineTimezone() {
        using var document = Standard(); var calendar = document.Calendar!;
        foreach (DayOfWeek day in Enum.GetValues(typeof(DayOfWeek))) calendar.SetWorkingDay(day);
        calendar.SetWorkingDay(DayOfWeek.Saturday, ProjectWorkingTime.Hours(22, 6));
        var start = new DateTime(2026, 10, 24, 22, 0, 0);
        Assert.Equal(start.AddHours(8), calendar.AddWorkingMinutes(start, 480));
        Assert.Equal(480, calendar.WorkingMinutesBetween(start, start.AddHours(8)));
        Assert.Equal(360, calendar.WorkingMinutesBetween(start.AddHours(2), start.AddHours(8)));
        Assert.Throws<ArgumentException>(() => calendar.AddWorkingMinutes(DateTime.SpecifyKind(start, DateTimeKind.Utc), 1));
    }
    [Fact]
    public void ClosedCalendarSearchIsBoundedAndCancelable() {
        using var document = Standard();
        foreach (DayOfWeek day in Enum.GetValues(typeof(DayOfWeek))) document.Calendar!.SetWorkingDay(day);
        Assert.Throws<InvalidOperationException>(() => document.Calendar!.AddWorkingMinutes(Monday, 1, 10));
        Assert.Throws<OperationCanceledException>(() => document.Calendar!.AddWorkingMinutes(Monday, 1, cancellationToken: new CancellationToken(true)));
    }
    [Fact]
    public void CriticalPathRollupsAndExplicitApplyKeepStoredDatesSeparate() {
        using var document = Standard();
        var summary = document.Tasks.AddSummary("Delivery");
        var a = summary.Children.Add("A"); a.Duration = ProjectDuration.WorkingDays(2);
        var b = summary.Children.Add("B"); b.Duration = ProjectDuration.WorkingDays(1);
        var c = summary.Children.Add("C"); c.Duration = ProjectDuration.WorkingDays(1);
        document.Dependencies.Add(a, c); document.Dependencies.Add(b, c);
        var before = document.Revision; var schedule = document.CalculateSchedule(); schedule.Report.ThrowIfErrors();
        Assert.Equal(before, document.Revision); Assert.Null(a.Start);
        var shortTask = schedule.Tasks.Single(t => t.TaskUid == b.Uid);
        Assert.Equal(480, shortTask.TotalSlackMinutes); Assert.Equal(480, shortTask.FreeSlackMinutes); Assert.False(shortTask.IsCritical);
        var finish = schedule.Tasks.Single(t => t.TaskUid == summary.Uid).Finish;
        Assert.Equal(Monday.AddDays(2).AddHours(9), finish);
        document.ApplySchedule(schedule); Assert.Equal(Monday, a.Start); Assert.False(document.IsScheduleStale);
        Assert.True(document.AreWorkCostTotalsStale);
        Assert.Contains(document.Validate().Diagnostics, d => d.Code == "PROJECT_WORK_COST_STALE");
        Assert.Throws<InvalidOperationException>(() => document.ApplySchedule(schedule));
        var fresh = document.CalculateSchedule(); a.Name = "Renamed";
        Assert.Throws<InvalidOperationException>(() => document.ApplySchedule(fresh));
    }
    [Theory]
    [InlineData(ProjectDependencyType.FinishToStart, 2, 10)]
    [InlineData(ProjectDependencyType.StartToStart, 0, 10)]
    [InlineData(ProjectDependencyType.FinishToFinish, 1, 10)]
    [InlineData(ProjectDependencyType.StartToFinish, -3, 10)]
    public void DependencyKindsUseBothEndpoints(ProjectDependencyType type, int days, int hour) {
        using var document = Standard();
        var a = document.Tasks.Add("Predecessor"); a.Duration = ProjectDuration.WorkingDays(2);
        var b = document.Tasks.Add("Successor"); b.Duration = ProjectDuration.WorkingDays(1);
        document.Dependencies.Add(a, b, type).Lag = ProjectDuration.WorkingHours(2);
        var result = document.CalculateSchedule(); result.Report.ThrowIfErrors();
        Assert.Equal(Monday.Date.AddDays(days).AddHours(hour), result.Tasks.Single(t => t.TaskUid == b.Uid).Start);
    }
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void PositivePercentageAndNegativeDurationLagAreDeterministic(bool percentage) {
        using var document = Standard();
        var a = document.Tasks.Add("Predecessor"); a.Duration = ProjectDuration.WorkingDays(2);
        var b = document.Tasks.Add("Successor"); b.Duration = ProjectDuration.WorkingDays(1);
        var link = document.Dependencies.Add(a, b);
        if (percentage) link.LagPercent = 50; else link.Lag = ProjectDuration.WorkingDays(-1);
        var result = document.CalculateSchedule(); result.Report.ThrowIfErrors();
        Assert.Equal(Monday.AddDays(percentage ? 3 : 1), result.Tasks.Single(t => t.TaskUid == b.Uid).Start);
    }
    [Fact]
    public void BackwardSchedulingAndDeadlinesExposeFloatWithoutMutatingTheDocument() {
        using var document = Standard(); document.Settings.ScheduleFromStart = false; document.Settings.FinishDate = Monday.AddDays(4).AddHours(9);
        var a = document.Tasks.Add("A"); a.Duration = ProjectDuration.WorkingDays(2);
        var b = document.Tasks.Add("B"); b.Duration = ProjectDuration.WorkingDays(1); document.Dependencies.Add(a, b);
        var result = document.CalculateSchedule(); result.Report.ThrowIfErrors();
        Assert.Equal(Monday.AddDays(2), result.Tasks.Single(t => t.TaskUid == a.Uid).Start);
        Assert.Equal(document.Settings.FinishDate, result.Tasks.Single(t => t.TaskUid == b.Uid).Finish);
        document.Settings.ScheduleFromStart = true; b.Deadline = Monday.AddDays(1).AddHours(9);
        result = document.CalculateSchedule(); result.Report.ThrowIfErrors();
        Assert.Equal(-480, result.Tasks.Single(t => t.TaskUid == b.Uid).TotalSlackMinutes);
        Assert.Contains(result.Report.Diagnostics, d => d.Code == "PROJECT_DEADLINE_MISSED");
    }
    [Fact]
    public void CyclesAndConflictingManualDatesBlockApplication() {
        using var document = Standard();
        var a = document.Tasks.Add("A"); a.Duration = ProjectDuration.WorkingDays(2);
        var b = document.Tasks.Add("B"); b.Duration = ProjectDuration.WorkingDays(1); b.IsManual = true; b.Start = Monday; b.Finish = Monday.AddHours(9);
        document.Dependencies.Add(a, b);
        var result = document.CalculateSchedule(); Assert.True(result.Report.HasErrors);
        var revision = document.Revision; Assert.Throws<InvalidDataException>(() => document.ApplySchedule(result)); Assert.Equal(revision, document.Revision);
        document.Dependencies.Add(b, a); result = document.CalculateSchedule();
        Assert.Contains(result.Report.Diagnostics, d => d.Code == "PROJECT_DEPENDENCY_CYCLE");
    }
    [Theory]
    [InlineData("<IgnoreResourceCalendar>1</IgnoreResourceCalendar>")]
    [InlineData("<ExternalTask>1</ExternalTask>")]
    public void UnsupportedRetainedTaskSchedulingInputsBlockApply(string extra) {
        using var original = Standard(); original.Tasks.Add("A").Duration = ProjectDuration.WorkingDays(1);
        var xml = System.Xml.Linq.XDocument.Parse(original.ToXml()); var ns = xml.Root!.Name.Namespace;
        var task = xml.Root.Element(ns + "Tasks")!.Element(ns + "Task")!;
        var field = System.Xml.Linq.XElement.Parse(extra); field.Name = ns + field.Name.LocalName; task.Add(field);
        using var document = ProjectDocument.Parse(xml.ToString()); var revision = document.Revision;
        var result = document.CalculateSchedule();
        Assert.Contains(result.Report.Diagnostics, d => d.Code == "PROJECT_TASK_SCHEDULING_PROFILE");
        Assert.Throws<InvalidDataException>(() => document.ApplySchedule(result)); Assert.Equal(revision, document.Revision);
    }
    [Fact]
    public void ImportedAssignmentDatesCannotSilentlyOverrideAnAppliedTaskSchedule() {
        using var document = ProjectDocument.Load(ProjectNativeTests.Fixture("delivery.xml"));
        document.Tasks.GetByUid(2).Duration = ProjectDuration.WorkingDays(4);
        long revision = document.Revision; var result = document.CalculateSchedule();
        Assert.NotEmpty(result.Tasks);
        Assert.Contains(result.Report.Diagnostics, d => d.Code == "PROJECT_ASSIGNMENT_DATE_RECALCULATION_REQUIRED");
        Assert.Throws<InvalidDataException>(() => document.ApplySchedule(result)); Assert.Equal(revision, document.Revision);
    }
    [Fact]
    public void ApplyingAuthoredDatesInitializesAssignmentEndpointsAndRemainingDurationWithoutChangingAmounts() {
        using var document = Standard(); document.Settings.ScheduleFromStart = null;
        var task = document.Tasks.Add("Assigned"); task.Duration = ProjectDuration.WorkingDays(2);
        var assignment = document.Assignments.Add(task, document.Resources.AddWork("Engineer"), ProjectUnits.Percent(50));
        assignment.Work = ProjectWork.Hours(8); assignment.Cost = 1000;
        document.Recalculate();
        Assert.True(document.Settings.ScheduleFromStart); Assert.Equal(task.Duration, task.RemainingDuration);
        Assert.Equal(task.Start, assignment.Start); Assert.Equal(task.Finish, assignment.Finish);
        Assert.Equal(480, assignment.Work?.Minutes); Assert.Equal(1000, assignment.Cost);
        Assert.False(document.IsScheduleStale); Assert.True(document.AreWorkCostTotalsStale);
        using var reloaded = ProjectDocument.Parse(document.ToXml());
        Assert.Equal(assignment.Start, reloaded.Assignments[0].Start); Assert.Equal(assignment.Finish, reloaded.Assignments[0].Finish);
    }
    [Fact]
    public void CalendarKindsCannotCreateAnAmbiguousTaskCalendarOnExport() {
        using var document = Standard();
        var invalidBase = document.Calendars.Add("Invalid base"); invalidBase.BaseCalendar = document.Calendar;
        Assert.Contains(document.Validate().Diagnostics, d => d.Code == "PROJECT_CALENDAR_KIND");
        invalidBase.IsBaseCalendar = false;
        var task = document.Tasks.Add("Task"); task.Duration = ProjectDuration.WorkingDays(1); task.Calendar = invalidBase;
        Assert.Contains(document.CalculateSchedule().Report.Diagnostics, d => d.Code == "PROJECT_TASK_CALENDAR_KIND");
        Assert.Throws<InvalidDataException>(() => document.ToXml());
        task.Calendar = null; document.Resources.AddWork("Engineer").Calendar = invalidBase;
        document.Validate().ThrowIfErrors();
    }
}

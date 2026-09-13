namespace OfficeIMO.Project.Tests;

public sealed class ProjectSaveBoundaryTests {
    private static readonly DateTime Monday = new(2026, 10, 5, 8, 0, 0);

    [Theory]
    [InlineData("task")]
    [InlineData("resource")]
    [InlineData("assignment")]
    public void UntypedTimephasedRecordsRemainSerializableButBlockScheduling(string owner) {
        using var document = ProjectDocument.Create(); document.Calendar = document.Calendars.AddStandardWorkingWeek();
        document.Settings.StartDate = Monday;
        var task = document.Tasks.Add("Task"); task.Duration = ProjectDuration.WorkingDays(1);
        var resource = document.Resources.AddWork("Engineer");
        var assignment = document.Assignments.Add(task, resource);
        var intervals = owner == "task" ? task.TimephasedData : owner == "resource" ? resource.TimephasedData : assignment.TimephasedData;
        intervals.Add().Uid = 1;
        using var copy = ProjectDocument.Parse(document.ToXml());
        var result = copy.CalculateSchedule(new ProjectScheduleOptions { CalculateAssignments = true });
        Assert.Contains(result.Report.Diagnostics, d => d.Code == "PROJECT_TIMEPHASED_TYPE_REQUIRED");
        Assert.Throws<InvalidDataException>(() => copy.ApplySchedule(result));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void EarnedValueDoesNotIgnoreAnUntypedTaskOrBaselineCurve(bool inBaseline) {
        using var document = ProjectDocument.Create(); document.Calendar = document.Calendars.AddStandardWorkingWeek();
        var task = document.Tasks.Add("Task"); task.ActualCost = 50; task.PercentComplete = 50;
        var baseline = task.Baselines.Add(); baseline.Number = 0; baseline.Cost = 100;
        (inBaseline ? baseline.TimephasedData : task.TimephasedData).Add().Uid = task.Uid;
        var result = document.AnalyzeEarnedValue(statusDate: Monday);
        Assert.Contains(result.Report.Diagnostics, d => d.Code == "PROJECT_EARNED_VALUE_INCOMPLETE");
        var item = Assert.Single(result.Tasks);
        Assert.Null(item.ActualCost); Assert.Null(item.EarnedValue); Assert.Null(item.PlannedValue);
    }

    [Theory]
    [InlineData("task")]
    [InlineData("resource")]
    [InlineData("assignment")]
    [InlineData("task baseline")]
    [InlineData("assignment baseline")]
    public void TimephasedUidIsRequiredButOtherFieldsRemainOptional(string owner) {
        using var document = ProjectDocument.Create();
        var task = document.Tasks.Add("Task"); var resource = document.Resources.AddWork("Engineer");
        var assignment = document.Assignments.Add(task, resource);
        var intervals = owner == "resource" ? resource.TimephasedData : owner == "assignment" ? assignment.TimephasedData : task.TimephasedData;
        if (owner.EndsWith("baseline", StringComparison.Ordinal)) {
            var baseline = (owner == "task baseline" ? task.Baselines : assignment.Baselines).Add();
            baseline.Number = 0; intervals = baseline.TimephasedData;
        }
        var interval = intervals.Add();
        Assert.Contains(document.Validate().Diagnostics, d => d.Code == "PROJECT_TIMEPHASED_UID");
        using var output = new MemoryStream();
        Assert.Throws<InvalidDataException>(() => document.Save(output)); Assert.Equal(0, output.Length);
        int ownerUid = owner == "resource" ? resource.Uid : owner.StartsWith("assignment", StringComparison.Ordinal) ? assignment.Uid : task.Uid;
        interval.Uid = ownerUid;
        Assert.False(document.Validate().HasErrors);
        using var copy = ProjectDocument.Parse(document.ToXml()); Assert.False(copy.Validate().HasErrors);
        interval.Value = "opaque retained value";
        Assert.Contains("opaque retained value", document.ToXml());
        interval.Uid = ownerUid + 100;
        Assert.Contains(document.Validate().Diagnostics, d => d.Code == "PROJECT_TIMEPHASED_UID"
            && d.Message.IndexOf("must match", StringComparison.Ordinal) >= 0);
        using var rejected = new MemoryStream();
        Assert.Throws<InvalidDataException>(() => document.Save(rejected)); Assert.Equal(0, rejected.Length);
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    public void IncompleteWorkWeekBoundsRoundTripButCannotDriveCalendarArithmetic(bool from, bool to) {
        using var document = ProjectDocument.Create(); var calendar = document.Calendars.AddStandardWorkingWeek();
        var week = calendar.WorkWeeks.Add(); if (from) week.FromDate = Monday; if (to) week.ToDate = Monday.AddDays(4);
        Assert.Contains(document.Validate().Diagnostics, d => d.Code == "PROJECT_CALENDAR_PERIOD");
        Assert.False(document.AssessSave().HasErrors);
        using (var retained = ProjectDocument.Parse(document.ToXml())) {
            var saved = retained.Calendars.GetByUid(calendar.Uid).WorkWeeks[0];
            Assert.Equal(week.FromDate, saved.FromDate); Assert.Equal(week.ToDate, saved.ToDate);
        }
        Assert.Throws<NotSupportedException>(() => calendar.GetWorkingIntervals(Monday));
        foreach (var format in new[] { ProjectFileFormat.Mpp8, ProjectFileFormat.Mpp9, ProjectFileFormat.Mpp12, ProjectFileFormat.Mpp14, ProjectFileFormat.Mpx4 })
            Assert.True(document.AssessSave(new ProjectSaveOptions { Format = format }).HasErrors);
        week.FromDate = Monday; week.ToDate = Monday.AddDays(4);
        using var copy = ProjectDocument.Parse(document.ToXml()); Assert.False(copy.Validate().HasErrors);
    }

    [Fact]
    public void DerivedCalendarRequiresItsParentAfterEditingOrLoading() {
        using var document = ProjectDocument.Create(); var parent = document.Calendars.AddStandardWorkingWeek();
        var calendar = document.Calendars.Add("Resource calendar", parent); calendar.BaseCalendar = null;
        Assert.Contains(document.Validate().Diagnostics, d => d.Code == "PROJECT_CALENDAR_REFERENCE");
        Assert.True(document.AssessSave().HasErrors);
        Assert.Throws<InvalidDataException>(() => calendar.GetWorkingIntervals(Monday));
        calendar.BaseCalendar = parent;
        using var copy = ProjectDocument.Parse(document.ToXml()); Assert.False(copy.Validate().HasErrors);
        copy.Calendars.GetByUid(calendar.Uid).BaseCalendar = null;
        Assert.True(copy.AssessSave().HasErrors);
    }

    [Theory]
    [InlineData(ProjectDataKind.Tasks)]
    [InlineData(ProjectDataKind.Resources)]
    [InlineData(ProjectDataKind.Assignments)]
    [InlineData(ProjectDataKind.Calendars)]
    public void TableImportRejectsNegativeIdentifiers(ProjectDataKind kind) {
        var fields = kind == ProjectDataKind.Assignments
            ? new[] { ProjectDataField.Uid, ProjectDataField.TaskUid, ProjectDataField.ResourceUid }
            : new[] { ProjectDataField.Uid, ProjectDataField.Name };
        var values = kind == ProjectDataKind.Assignments ? new[] { "-1", "1", "1" } : new[] { "-1", "Record" };
        var table = new ProjectDataTable(fields.Select(f => f.ToString()), new[] { values });
        var mapped = new ProjectMappedTable(kind, table, fields.Select(f => new ProjectDataColumn(f, f.ToString())));
        var error = Assert.Throws<InvalidDataException>(() => ProjectDocument.ImportTables(new[] { mapped }));
        Assert.Contains("Uid: expected a nonnegative integer", error.Message);
    }

    [Theory]
    [InlineData("none", false)]
    [InlineData("none", true)]
    [InlineData("task", false)]
    [InlineData("task", true)]
    [InlineData("assignment", false)]
    [InlineData("summary", false)]
    [InlineData("summary", true)]
    public void SummaryStartCostAccruesAtTheFirstActualStart(string source, bool milestone) {
        using var document = ProjectDocument.Create(); document.Calendar = document.Calendars.AddStandardWorkingWeek();
        document.Settings.StartDate = Monday;
        var outer = document.Tasks.AddSummary("Delivery"); outer.FixedCost = 200; outer.FixedCostAccrual = ProjectCostAccrual.Start;
        var inner = outer.Children.AddSummary("Phase"); inner.FixedCost = 100; inner.FixedCostAccrual = ProjectCostAccrual.Start;
        var task = inner.Children.Add("Task"); task.Duration = ProjectDuration.WorkingDays(milestone ? 0 : 1); task.IsMilestone = milestone;
        bool started = source != "none";
        if (source == "task") task.ActualStart = Monday;
        if (source == "summary") inner.ActualStart = Monday;
        if (source == "assignment") {
            var resource = document.Resources.AddWork("Engineer"); resource.StandardRate = 0;
            document.Assignments.Add(task, resource).ActualStart = Monday;
        }
        var options = new ProjectScheduleOptions { CalculateAssignments = true };
        var schedule = document.CalculateSchedule(options); schedule.Report.ThrowIfErrors();
        document.ApplySchedule(schedule);
        Assert.Equal(started ? 300m : 0m, outer.ActualCost); Assert.Equal(started ? 100m : 0m, inner.ActualCost);
        Assert.Equal(0, outer.PercentComplete);
        using var copy = ProjectDocument.Parse(document.ToXml());
        var repeated = copy.CalculateSchedule(options); repeated.Report.ThrowIfErrors(); copy.ApplySchedule(repeated);
        Assert.Equal(outer.ActualCost, copy.Tasks.GetByUid(outer.Uid).ActualCost);
    }
}

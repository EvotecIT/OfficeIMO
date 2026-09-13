namespace OfficeIMO.Project.Tests;

public sealed class ProjectValidationBoundaryTests {
    private static readonly DateTime Monday = new(2026, 10, 5, 8, 0, 0);

    [Theory]
    [InlineData("task")]
    [InlineData("resource")]
    [InlineData("assignment")]
    public void ScalarValuesRequireUniqueNormalizedFieldIds(string owner) {
        using var document = ProjectDocument.Create();
        var task = document.Tasks.Add("Task"); var resource = document.Resources.AddWork("Engineer");
        var fields = owner == "task" ? task.CustomFields : owner == "resource" ? resource.CustomFields : document.Assignments.Add(task, resource).CustomFields;
        fields.Add().FieldId = "188743731"; fields.Add().FieldId = "0188743731";
        Assert.Contains(document.Validate().Diagnostics, d => d.Code == "PROJECT_CUSTOM_FIELD_ID");
        Assert.True(document.AssessSave().HasErrors);
        Assert.True(document.CalculateCustomFields().Report.HasErrors);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ProjectSummaryRollupIsIndependentOfImportedRowOrder(bool xmlRoundTrip) {
        var fields = new[] { ProjectDataField.Uid, ProjectDataField.Name, ProjectDataField.Summary, ProjectDataField.ParentUid, ProjectDataField.DurationMinutes };
        var table = new ProjectDataTable(fields.Select(f => f.ToString()), new[] {
            new[] { "1", "Phase", "true", "", "480" }, new[] { "2", "Work", "false", "1", "480" },
            new[] { "0", "Project", "true", "", "480" }
        });
        var mapped = new ProjectMappedTable(ProjectDataKind.Tasks, table, fields.Select(f => new ProjectDataColumn(f, f.ToString())));
        using var imported = ProjectDocument.ImportTables(new[] { mapped }).Document;
        imported.Calendar = imported.Calendars.AddStandardWorkingWeek(); imported.Settings.StartDate = Monday;
        imported.Tasks.GetByUid(2).FixedCost = 200m;
        using var document = xmlRoundTrip ? imported.Clone() : imported;
        var result = document.CalculateSchedule(new ProjectScheduleOptions { CalculateAssignments = true }); result.Report.ThrowIfErrors();
        var summary = Assert.Single(result.Tasks, t => t.TaskUid == 0);
        Assert.Equal(Monday, summary.Start); Assert.Equal(Monday.AddHours(9), summary.Finish);
        document.ApplySchedule(result); Assert.Equal(200m, document.Tasks.GetByUid(0).Cost);
    }

    [Theory]
    [InlineData("0188743731")]
    [InlineData("+188743731")]
    [InlineData(" 188743731 ")]
    public void EquivalentCustomFieldIdentitiesRejectSaveAndCalculation(string alternate) {
        using var document = ProjectDocument.Create(); document.Tasks.Add("Task");
        document.CustomFields.Add().FieldId = "188743731";
        document.CustomFields.Add().FieldId = alternate;
        Assert.Contains(document.Validate().Diagnostics, d => d.Code == "PROJECT_CUSTOM_FIELD_ID");
        Assert.True(document.AssessSave().HasErrors);
        Assert.True(document.CalculateCustomFields().Report.HasErrors);
        using var output = new MemoryStream();
        Assert.Throws<InvalidDataException>(() => document.Save(output)); Assert.Equal(0, output.Length);
    }

    [Theory]
    [InlineData(-1, false)]
    [InlineData(1001, false)]
    [InlineData(0, true)]
    [InlineData(1000, true)]
    public void TaskPriorityMustStayWithinTheLevelingContract(int priority, bool valid) {
        using var document = ProjectDocument.Create(); var task = document.Tasks.Add("Task"); task.Priority = priority;
        Assert.Equal(valid, !document.Validate().HasErrors);
        Assert.Equal(valid, !document.AssessSave().HasErrors);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void OversizedDelayValuesBecomeValidationDiagnostics(bool taskDelay) {
        using var document = ProjectDocument.Create(); document.Calendar = document.Calendars.AddStandardWorkingWeek();
        document.Settings.StartDate = Monday;
        var task = document.Tasks.Add("Task"); task.Duration = ProjectDuration.WorkingDays(1);
        var resource = document.Resources.AddWork("Engineer"); var assignment = document.Assignments.Add(task, resource);
        if (taskDelay) task.LevelingDelay = ProjectDuration.WorkingMinutes(decimal.MaxValue);
        else assignment.DelayMinutes = decimal.MaxValue;
        string code = taskDelay ? "PROJECT_LEVELING_DELAY" : "PROJECT_ASSIGNMENT_DELAY";
        Assert.Contains(document.Validate().Diagnostics, diagnostic => diagnostic.Code == code);
        Assert.Contains(document.AssessSave().Diagnostics, diagnostic => diagnostic.Code == code);
        Assert.Contains(document.CalculateSchedule(new ProjectScheduleOptions { CalculateAssignments = true }).Report.Diagnostics,
            diagnostic => diagnostic.Code == code);
    }

    [Theory]
    [InlineData("leveling", false)]
    [InlineData("leveling", true)]
    [InlineData("lag", false)]
    [InlineData("lag", true)]
    [InlineData("negative lag", false)]
    [InlineData("negative lag", true)]
    public void XmlDurationsThatCrossTheTimeSpanBoundaryAreRejectedBeforeSave(string field, bool outside) {
        using var document = ProjectDocument.Create(); document.Calendar = document.Calendars.AddStandardWorkingWeek();
        var predecessor = document.Tasks.Add("Predecessor"); predecessor.Duration = ProjectDuration.WorkingMinutes(1);
        var successor = document.Tasks.Add("Successor"); successor.Duration = ProjectDuration.WorkingMinutes(1);
        decimal value = decimal.Floor(long.MaxValue / (decimal)TimeSpan.TicksPerMinute);
        if (outside) value += 1m;
        if (field == "negative lag") value = -value;
        if (field == "leveling") successor.LevelingDelay = ProjectDuration.WorkingMinutes(value);
        else document.Dependencies.Add(predecessor, successor).Lag = ProjectDuration.WorkingMinutes(value);
        string code = field == "leveling" ? "PROJECT_LEVELING_DELAY" : "PROJECT_LAG_RANGE";
        var options = new ProjectSaveOptions { Format = ProjectFileFormat.Xml };
        Assert.Equal(outside, document.AssessSave(options).Diagnostics.Any(diagnostic => diagnostic.Code == code));
        using var output = new MemoryStream();
        if (outside) { Assert.Throws<InvalidDataException>(() => document.Save(output, options)); Assert.Equal(0, output.Length); }
        else { document.Save(output, options); using var reopened = ProjectDocument.Load(new MemoryStream(output.ToArray())); Assert.False(reopened.Validate().HasErrors); }
    }

    [Theory]
    [InlineData(ProjectFileFormat.Mpp8)]
    [InlineData(ProjectFileFormat.Mpp9)]
    [InlineData(ProjectFileFormat.Mpp12)]
    [InlineData(ProjectFileFormat.Mpp14)]
    public void OversizedMonthSettingsAreRejectedForEveryNativeGeneration(ProjectFileFormat format) {
        using var document = ProjectNativeAuthoringTests.Create();
        document.Settings.MinutesPerDay = int.MaxValue; document.Settings.DaysPerMonth = int.MaxValue;
        Assert.Contains(document.Validate().Diagnostics, diagnostic => diagnostic.Code == "PROJECT_WORKING_TIME");
        var options = new ProjectSaveOptions { Format = format, LossPolicy = OfficeConversionLossPolicy.Allow };
        Assert.Contains(document.AssessSave(options).Diagnostics, diagnostic => diagnostic.Code == "PROJECT_WORKING_TIME");
        using var output = new MemoryStream(); Assert.Throws<InvalidDataException>(() => document.Save(output, options)); Assert.Equal(0, output.Length);
    }

    [Fact]
    public void OversizedMonthSettingsBlockVariableMaterialScheduling() {
        using var document = ProjectDocument.Create(); document.Calendar = document.Calendars.AddStandardWorkingWeek();
        document.Settings.StartDate = Monday; document.Settings.MinutesPerDay = int.MaxValue; document.Settings.DaysPerMonth = int.MaxValue;
        var task = document.Tasks.Add("Task"); task.Duration = ProjectDuration.WorkingDays(1);
        var material = document.Resources.AddMaterial("Material"); material.MaterialLabel = "unit";
        var assignment = document.Assignments.Add(task, material); assignment.HasFixedRateUnits = false; assignment.MaterialRateScale = 5;
        var result = document.CalculateSchedule(new ProjectScheduleOptions { CalculateAssignments = true });
        Assert.Contains(result.Report.Diagnostics, diagnostic => diagnostic.Code == "PROJECT_WORKING_TIME");
    }

    [Theory]
    [InlineData("weekday", false, true)]
    [InlineData("exception", false, true)]
    [InlineData("workweek", false, true)]
    [InlineData("weekday", true, false)]
    [InlineData("exception", true, false)]
    [InlineData("workweek", true, false)]
    [InlineData("weekday", null, true)]
    [InlineData("exception", null, true)]
    [InlineData("workweek", null, true)]
    public void WorkingStatusMustAgreeWithItsIntervals(string kind, bool? working, bool addInterval) {
        using var document = ProjectDocument.Create(); var calendar = document.Calendars.Add("Calendar");
        ProjectCollection<ProjectWorkingInterval> times;
        if (kind == "exception") {
            var item = calendar.Exceptions.Add(); item.FromDate = Monday; item.ToDate = Monday;
            item.IsWorking = working; times = item.WorkingTimes;
        } else {
            var days = calendar.WeekDays;
            if (kind == "workweek") {
                var week = calendar.WorkWeeks.Add(); week.FromDate = Monday; week.ToDate = Monday.AddDays(4);
                days = week.WeekDays;
            }
            var day = days.Add(); day.Day = DayOfWeek.Monday; day.IsWorking = working; times = day.WorkingTimes;
        }
        if (addInterval) { var interval = times.Add(); interval.From = TimeSpan.FromHours(8); interval.To = TimeSpan.FromHours(12); }
        Assert.Contains(document.Validate().Diagnostics, d => d.Code == "PROJECT_CALENDAR_WORKING_STATUS");
        Assert.True(document.AssessSave().HasErrors);
    }

    [Theory]
    [InlineData("weekday", "duplicate")]
    [InlineData("weekday", "partial")]
    [InlineData("weekday", "overnight")]
    [InlineData("weekday", "full day")]
    [InlineData("legacy", "duplicate")]
    [InlineData("legacy", "partial")]
    [InlineData("legacy", "overnight")]
    [InlineData("legacy", "full day")]
    [InlineData("exception", "duplicate")]
    [InlineData("exception", "partial")]
    [InlineData("exception", "overnight")]
    [InlineData("exception", "full day")]
    [InlineData("workweek", "duplicate")]
    [InlineData("workweek", "partial")]
    [InlineData("workweek", "overnight")]
    [InlineData("workweek", "full day")]
    public void WorkingIntervalsAtTheSamePrecedenceCannotOverlap(string kind, string pattern) {
        using var document = ProjectDocument.Create(); var calendar = document.Calendars.Add("Calendar");
        ProjectCollection<ProjectWorkingInterval> times;
        if (kind == "exception") {
            var item = calendar.Exceptions.Add(); item.FromDate = Monday; item.ToDate = Monday; item.IsWorking = true; times = item.WorkingTimes;
        } else {
            var days = calendar.WeekDays;
            if (kind == "workweek") {
                var week = calendar.WorkWeeks.Add(); week.FromDate = Monday; week.ToDate = Monday.AddDays(4); days = week.WeekDays;
            }
            var day = days.Add(); day.IsWorking = true;
            if (kind == "legacy") { day.FromDate = Monday; day.ToDate = Monday; }
            else day.Day = DayOfWeek.Monday;
            times = day.WorkingTimes;
        }
        (int From, int To, int OtherFrom, int OtherTo) = pattern switch {
            "duplicate" => (8, 12, 8, 12), "overnight" => (22, 2, 1, 3), "full day" => (8, 8, 12, 13), _ => (8, 12, 10, 14)
        };
        var first = times.Add(); first.From = TimeSpan.FromHours(From); first.To = TimeSpan.FromHours(To);
        var second = times.Add(); second.From = TimeSpan.FromHours(OtherFrom); second.To = TimeSpan.FromHours(OtherTo);
        Assert.Contains(document.Validate().Diagnostics, d => d.Code == "PROJECT_WORKING_INTERVAL_OVERLAP");
        foreach (var format in new[] { ProjectFileFormat.Xml, ProjectFileFormat.Mpx4, ProjectFileFormat.Mpp8, ProjectFileFormat.Mpp9,
            ProjectFileFormat.Mpp12, ProjectFileFormat.Mpp14, ProjectFileFormat.Mpt8, ProjectFileFormat.Mpt9, ProjectFileFormat.Mpt12, ProjectFileFormat.Mpt14 })
            Assert.True(document.AssessSave(new ProjectSaveOptions { Format = format }).HasErrors);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void AdjacentWorkingIntervalsDoNotOverlap(bool overnight) {
        using var document = ProjectDocument.Create(); var calendar = document.Calendars.Add("Calendar");
        var day = calendar.WeekDays.Add(); day.Day = DayOfWeek.Monday; day.IsWorking = true;
        var first = day.WorkingTimes.Add(); first.From = TimeSpan.FromHours(overnight ? 22 : 8); first.To = TimeSpan.FromHours(overnight ? 2 : 12);
        var second = day.WorkingTimes.Add(); second.From = first.To; second.To = TimeSpan.FromHours(overnight ? 6 : 14);
        Assert.DoesNotContain(document.Validate().Diagnostics, d => d.Code == "PROJECT_WORKING_INTERVAL_OVERLAP");
        second.From = TimeSpan.FromHours(overnight ? 1 : 11);
        Assert.Contains(document.Validate().Diagnostics, d => d.Code == "PROJECT_WORKING_INTERVAL_OVERLAP");
    }

    [Fact]
    public void LargeNonOverlappingWorkingIntervalSetsRemainValid() {
        using var document = ProjectDocument.Create(); var calendar = document.Calendars.Add("Calendar");
        var day = calendar.WeekDays.Add(); day.Day = DayOfWeek.Monday; day.IsWorking = true;
        for (int index = 0; index < 10_000; index++) {
            var interval = day.WorkingTimes.Add(); interval.From = TimeSpan.FromTicks(index * 4L); interval.To = TimeSpan.FromTicks(index * 4L + 1);
        }
        Assert.DoesNotContain(document.Validate().Diagnostics, d => d.Code == "PROJECT_WORKING_INTERVAL_OVERLAP");
    }

    [Theory]
    [InlineData(false, true)]
    [InlineData(true, false)]
    public void CalendarExceptionCalculationRequiresBothDateBounds(bool from, bool to) {
        using var document = ProjectDocument.Create(); var calendar = document.Calendars.AddStandardWorkingWeek();
        var exception = calendar.Exceptions.Add();
        if (from) exception.FromDate = Monday; if (to) exception.ToDate = Monday;
        exception.IsWorking = false;
        Assert.Contains(document.Validate().Diagnostics, diagnostic => diagnostic.Code == "PROJECT_CALENDAR_PERIOD"
            && diagnostic.Severity == ProjectDiagnosticSeverity.Warning);
        Assert.Throws<NotSupportedException>(() => calendar.GetWorkingIntervals(Monday));
    }

    [Fact]
    public void XmlUidZeroCannotDeclareAnOrdinaryTask() {
        const string xml = "<Project xmlns=\"http://schemas.microsoft.com/project\"><Tasks><Task><UID>0</UID><Name>Work</Name><Summary>0</Summary></Task></Tasks></Project>";
        using var input = new MemoryStream(System.Text.Encoding.UTF8.GetBytes(xml));
        using var document = ProjectDocument.Load(input);
        Assert.Contains(document.Validate().Diagnostics, d => d.Code == "PROJECT_TASK_RESERVED_UID");
        Assert.True(document.AssessSave().HasErrors);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void TaskTableUidZeroRequiresAnExplicitProjectSummary(bool summary) {
        var fields = new[] { ProjectDataField.Uid, ProjectDataField.Name, ProjectDataField.Summary, ProjectDataField.DurationMinutes };
        var table = new ProjectDataTable(fields.Select(f => f.ToString()), new[] {
            new[] { "0", "Project", summary ? "true" : "false", "60" }, new[] { "1", "Work", "false", "60" }
        });
        var mapped = new ProjectMappedTable(ProjectDataKind.Tasks, table, fields.Select(f => new ProjectDataColumn(f, f.ToString())));
        if (!summary) { Assert.Throws<InvalidDataException>(() => ProjectDocument.ImportTables(new[] { mapped })); return; }
        using var document = ProjectDocument.ImportTables(new[] { mapped }).Document;
        document.Calendar = document.Calendars.AddStandardWorkingWeek(); document.Settings.StartDate = Monday;
        var result = document.CalculateSchedule(); result.Report.ThrowIfErrors();
        Assert.Equal(2, result.Tasks.Count); Assert.True(result.Tasks.Single(t => t.TaskUid == 0).IsSummary);
        using var copy = document.Clone(); Assert.True(copy.Tasks.GetByUid(0).IsSummary);
        using var mpx = new MemoryStream(); document.Save(mpx, new ProjectSaveOptions { Format = ProjectFileFormat.Mpx4, LossPolicy = OfficeConversionLossPolicy.Allow });
        using var reopened = ProjectDocument.Load(new MemoryStream(mpx.ToArray()));
        Assert.Equal(0, reopened.Tasks.GetByUid(0).DisplayId); Assert.Equal(0, reopened.Tasks.GetByUid(0).SourceOutlineLevel);
        Assert.Null(reopened.Tasks.GetByUid(1).Parent);
    }
}

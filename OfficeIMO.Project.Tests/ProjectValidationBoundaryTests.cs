namespace OfficeIMO.Project.Tests;

public sealed class ProjectValidationBoundaryTests {
    private static readonly DateTime Monday = new(2026, 10, 5, 8, 0, 0);

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
    }
}

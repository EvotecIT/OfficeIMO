namespace OfficeIMO.Project.Tests;

public sealed class ProjectSaveBoundaryTests {
    private static readonly DateTime Monday = new(2026, 10, 5, 8, 0, 0);

    [Theory]
    [InlineData("task cost")]
    [InlineData("task actual")]
    [InlineData("task remaining")]
    [InlineData("task fixed")]
    [InlineData("task baseline cost")]
    [InlineData("task baseline fixed")]
    [InlineData("task baseline bcws")]
    [InlineData("task baseline bcwp")]
    [InlineData("resource per use")]
    [InlineData("resource cost")]
    [InlineData("resource actual")]
    [InlineData("resource remaining")]
    [InlineData("resource rate per use")]
    [InlineData("resource baseline cost")]
    [InlineData("resource baseline bcws")]
    [InlineData("resource baseline bcwp")]
    [InlineData("assignment cost")]
    [InlineData("assignment actual")]
    [InlineData("assignment remaining")]
    [InlineData("assignment baseline cost")]
    [InlineData("assignment baseline bcws")]
    [InlineData("assignment baseline bcwp")]
    public void ProjectHundredthsStorageRejectsOversizedCostsDuringAssessment(string owner) {
        using var document = ProjectDocument.Create();
        var task = document.Tasks.Add("Task"); var resource = document.Resources.AddWork("Engineer");
        var assignment = document.Assignments.Add(task, resource);
        var taskBaseline = task.Baselines.Add(); taskBaseline.Number = 0;
        var resourceBaseline = resource.Baselines.Add(); resourceBaseline.Number = 0;
        var assignmentBaseline = assignment.Baselines.Add(); assignmentBaseline.Number = 0;
        var rate = resource.Rates.Add(); rate.From = Monday; rate.To = Monday.AddDays(1);
        string location = owner switch {
            "task cost" => "/Task[UID=1]/Cost", "task actual" => "/Task[UID=1]/ActualCost", "task remaining" => "/Task[UID=1]/RemainingCost",
            "task fixed" => "/Task[UID=1]/FixedCost", "task baseline cost" => "/Task[UID=1]/Baseline[0]/Cost",
            "task baseline fixed" => "/Task[UID=1]/Baseline[0]/FixedCost", "task baseline bcws" => "/Task[UID=1]/Baseline[0]/BCWS",
            "task baseline bcwp" => "/Task[UID=1]/Baseline[0]/BCWP", "resource per use" => "/Resource[UID=1]/CostPerUse",
            "resource cost" => "/Resource[UID=1]/Cost", "resource actual" => "/Resource[UID=1]/ActualCost",
            "resource remaining" => "/Resource[UID=1]/RemainingCost", "resource rate per use" => "/Resource[UID=1]/Rate[0]/CostPerUse",
            "resource baseline cost" => "/Resource[UID=1]/Baseline[0]/Cost", "resource baseline bcws" => "/Resource[UID=1]/Baseline[0]/BCWS",
            "resource baseline bcwp" => "/Resource[UID=1]/Baseline[0]/BCWP", "assignment cost" => "/Assignment[UID=1]/Cost",
            "assignment actual" => "/Assignment[UID=1]/ActualCost", "assignment remaining" => "/Assignment[UID=1]/RemainingCost",
            "assignment baseline cost" => "/Assignment[UID=1]/Baseline[0]/Cost", "assignment baseline bcws" => "/Assignment[UID=1]/Baseline[0]/BCWS",
            _ => "/Assignment[UID=1]/Baseline[0]/BCWP"
        };
        void Set(decimal value) {
            switch (owner) {
                case "task cost": task.Cost = value; break;
                case "task actual": task.ActualCost = value; break;
                case "task remaining": task.RemainingCost = value; break;
                case "task fixed": task.FixedCost = value; break;
                case "task baseline cost": taskBaseline.Cost = value; break;
                case "task baseline fixed": taskBaseline.FixedCost = value; break;
                case "task baseline bcws": taskBaseline.Bcws = value; break;
                case "task baseline bcwp": taskBaseline.Bcwp = value; break;
                case "resource per use": resource.CostPerUse = value; break;
                case "resource cost": resource.Cost = value; break;
                case "resource actual": resource.ActualCost = value; break;
                case "resource remaining": resource.RemainingCost = value; break;
                case "resource rate per use": rate.CostPerUse = value; break;
                case "resource baseline cost": resourceBaseline.Cost = value; break;
                case "resource baseline bcws": resourceBaseline.Bcws = value; break;
                case "resource baseline bcwp": resourceBaseline.Bcwp = value; break;
                case "assignment cost": assignment.Cost = value; break;
                case "assignment actual": assignment.ActualCost = value; break;
                case "assignment remaining": assignment.RemainingCost = value; break;
                case "assignment baseline cost": assignmentBaseline.Cost = value; break;
                case "assignment baseline bcws": assignmentBaseline.Bcws = value; break;
                case "assignment baseline bcwp": assignmentBaseline.Bcwp = value; break;
            }
        }
        var options = new ProjectSaveOptions { Format = ProjectFileFormat.Xml };
        decimal maximum = decimal.MaxValue / 100m, minimum = decimal.MinValue / 100m;
        Set(maximum); Assert.DoesNotContain(document.AssessSave(options).Diagnostics, d => d.Code == "PROJECT_COST_RANGE");
        Set(maximum + .01m); Assert.Contains(document.AssessSave(options).Diagnostics, d => d.Code == "PROJECT_COST_RANGE" && d.Location == location);
        Set(minimum); Assert.DoesNotContain(document.AssessSave(options).Diagnostics, d => d.Code == "PROJECT_COST_RANGE");
        Set(minimum - .01m); Assert.Contains(document.AssessSave(options).Diagnostics, d => d.Code == "PROJECT_COST_RANGE" && d.Location == location);
        using var output = new MemoryStream();
        Assert.Throws<InvalidDataException>(() => document.Save(output)); Assert.Equal(0, output.Length);
    }

    [Theory]
    [InlineData("duration")]
    [InlineData("work")]
    [InlineData("baseline duration")]
    [InlineData("baseline work")]
    [InlineData("total slack")]
    [InlineData("free slack")]
    public void XmlAssessmentReportsOversizedScaledValuesBeforeSerialization(string field) {
        using var document = ProjectDocument.Create(); var task = document.Tasks.Add("Task");
        var baseline = task.Baselines.Add(); baseline.Number = 0;
        string code;
        switch (field) {
            case "duration": task.Duration = ProjectDuration.WorkingMinutes(decimal.MaxValue); code = "PROJECT_XML_DURATION_RANGE"; break;
            case "work": task.Work = new ProjectWork(decimal.MaxValue); code = "PROJECT_XML_WORK_RANGE"; break;
            case "baseline duration": baseline.Duration = ProjectDuration.WorkingMinutes(decimal.MaxValue); code = "PROJECT_XML_DURATION_RANGE"; break;
            case "baseline work": baseline.Work = new ProjectWork(decimal.MaxValue); code = "PROJECT_XML_WORK_RANGE"; break;
            case "total slack": task.TotalSlackMinutes = decimal.MaxValue; code = "PROJECT_XML_TENTHS_RANGE"; break;
            default: task.FreeSlackMinutes = decimal.MaxValue; code = "PROJECT_XML_TENTHS_RANGE"; break;
        }
        var options = new ProjectSaveOptions { Format = ProjectFileFormat.Xml };
        Assert.Contains(document.AssessSave(options).Diagnostics, diagnostic => diagnostic.Code == code);
        using var output = new MemoryStream();
        Assert.Throws<InvalidDataException>(() => document.Save(output, options)); Assert.Equal(0, output.Length);
    }

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

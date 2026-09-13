namespace OfficeIMO.Project.Tests;

public sealed class ProjectRoundTripAnalysisBoundaryTests {
    private static readonly DateTime Monday = new(2026, 10, 5, 8, 0, 0);

    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    [InlineData(2)]
    [InlineData(3)]
    [InlineData(4)]
    public void XmlRejectsConflictingDurationFormatsBeforeWriting(int conflict) {
        using var document = ProjectDocument.Create(); var task = document.Tasks.Add("Work");
        task.Duration = ProjectDuration.WorkingDays(2);
        task.ActualDuration = conflict switch {
            0 => ProjectDuration.WorkingHours(8),
            1 => ProjectDuration.ElapsedDays(1),
            2 => new ProjectDuration(1, ProjectDurationUnit.Day, false, true),
            _ => ProjectDuration.WorkingDays(1)
        };
        task.RemainingDuration = conflict >= 3 ? ProjectDuration.WorkingHours(8) : ProjectDuration.WorkingDays(1);
        if (conflict == 4) task.Duration = null;
        var options = new ProjectSaveOptions { Format = ProjectFileFormat.Xml };
        Assert.Contains(document.AssessSave(options).Diagnostics, d => d.Code == "PROJECT_XML_DURATION_FORMAT" && d.Severity == ProjectDiagnosticSeverity.Error);
        Assert.Throws<InvalidDataException>(() => document.ToXml());
        using var destination = new MemoryStream(); destination.WriteByte(123);
        Assert.Throws<InvalidDataException>(() => document.Save(destination, options)); Assert.Equal(new byte[] { 123 }, destination.ToArray());
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(true, false)]
    [InlineData(false, true)]
    [InlineData(true, true)]
    public void XmlRetainsProgressFormatWithoutMainDuration(bool elapsed, bool estimated) {
        using var document = ProjectDocument.Create(); var task = document.Tasks.Add("Work");
        task.ActualDuration = new ProjectDuration(2, ProjectDurationUnit.Hour, elapsed, estimated);
        task.RemainingDuration = new ProjectDuration(3, ProjectDurationUnit.Hour, elapsed, estimated);
        using var copy = ProjectDocument.Load(new MemoryStream(Encoding.UTF8.GetBytes(document.ToXml())));
        Assert.Null(copy.Tasks[0].Duration); Assert.Equal(task.ActualDuration, copy.Tasks[0].ActualDuration); Assert.Equal(task.RemainingDuration, copy.Tasks[0].RemainingDuration);
        copy.Tasks[0].Name = "Edited";
        using var again = copy.Clone(); Assert.Equal(task.ActualDuration, again.Tasks[0].ActualDuration); Assert.Equal(task.RemainingDuration, again.Tasks[0].RemainingDuration);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ResourceRemainingCostCacheIsAnIndependentObservation(bool missingContributor) {
        using var document = ProjectDocument.Create(); var task = document.Tasks.Add("Work"); var resource = document.Resources.AddWork("Engineer");
        var assignment = document.Assignments.Add(task, resource); assignment.Cost = resource.Cost = 100m; assignment.ActualCost = resource.ActualCost = 25m;
        assignment.RemainingCost = missingContributor ? null : 75m; resource.RemainingCost = 10m;
        long revision = document.Revision;
        var result = document.AnalyzeAssignments(); var totals = Assert.Single(result.Resources);
        Assert.Equal(10m, totals.StoredResourceRemainingCost); Assert.Equal(assignment.RemainingCost, totals.RemainingCost);
        Assert.Equal(!missingContributor, result.Report.Diagnostics.Any(d => d.Code == "PROJECT_RESOURCE_CACHED_TOTAL"));
        Assert.Equal(revision, document.Revision); Assert.Equal(10m, resource.RemainingCost);
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData(" \t ")]
    public void TableProjectionRetainsNamelessEntitiesAndReservedResource(string? name) {
        const string xml = "<Project xmlns=\"http://schemas.microsoft.com/project\"><Calendars><Calendar><UID>1</UID><IsBaseCalendar>1</IsBaseCalendar></Calendar></Calendars><Tasks><Task><UID>1</UID><CalendarUID>1</CalendarUID></Task></Tasks><Resources><Resource><UID>0</UID></Resource><Resource><UID>1</UID><CalendarUID>1</CalendarUID></Resource></Resources><Assignments><Assignment><UID>1</UID><TaskUID>1</TaskUID><ResourceUID>1</ResourceUID></Assignment></Assignments></Project>";
        using var document = ProjectDocument.Load(new MemoryStream(Encoding.UTF8.GetBytes(xml)));
        document.Tasks[0].Name = name; document.Resources.GetByUid(1).Name = name; document.Calendars[0].Name = name;
        var export = document.ExportTables(allowLossyProjection: true);
        using var copy = ProjectDocument.ImportTables(export.Tables).Document;
        Assert.Null(copy.Resources.GetByUid(0).Name); Assert.Equal(name, copy.Tasks[0].Name);
        Assert.Equal(name, copy.Resources.GetByUid(1).Name); Assert.Equal(name, copy.Calendars[0].Name);
        Assert.Equal(1, copy.Assignments[0].Task!.Uid); Assert.Equal(1, copy.Assignments[0].Resource!.Uid);
        Assert.Same(copy.Calendars[0], copy.Tasks[0].Calendar);
    }

    [Theory]
    [InlineData("task", -1)]
    [InlineData("task", 0)]
    [InlineData("task", 1)]
    [InlineData("assignment", -1)]
    [InlineData("assignment", 0)]
    [InlineData("assignment", 1)]
    [InlineData("fixed", -1)]
    [InlineData("fixed", 0)]
    [InlineData("fixed", 1)]
    public void CurrentStatusEstablishesEnteredActualCostsWithoutInventingHistoricalCosts(string owner, int offset) {
        using var document = ProjectDocument.Create(); document.Calendar = document.Calendars.AddStandardWorkingWeek(); document.Settings.StatusDate = Monday.AddHours(4);
        var task = document.Tasks.Add("Work"); task.ActualStart = Monday; task.ActualCost = 100m;
        task.EarnedValueMethod = ProjectEarnedValueMethod.PhysicalPercentComplete; task.PhysicalPercentComplete = 50;
        var baseline = task.Baselines.Add(); baseline.Number = 0; baseline.Cost = 200m; baseline.Start = Monday; baseline.Finish = Monday.AddHours(9);
        if (owner != "task") {
            var assignment = document.Assignments.Add(task, document.Resources.AddWork("Engineer")); assignment.ActualStart = Monday;
            assignment.ActualCost = owner == "fixed" ? 60m : 100m;
        }
        long revision = document.Revision;
        var result = Assert.Single(document.AnalyzeEarnedValue(statusDate: document.Settings.StatusDate.Value.AddHours(offset)).Tasks);
        Assert.Equal(offset < 0 ? null : 100m, result.ActualCost); Assert.Equal(revision, document.Revision);
        Assert.Equal(offset < 0 ? null : 1m, result.CostPerformanceIndex);
        Assert.Equal(offset < 0 ? null : 200m, result.EstimateAtCompletion);
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(true, false)]
    [InlineData(false, true)]
    [InlineData(true, true)]
    public void AppliedProgressSharesTheTaskDurationFormat(bool elapsed, bool estimated) {
        using var document = ProjectDocument.Create(); document.Calendar = document.Calendars.AddStandardWorkingWeek(); document.Settings.StartDate = Monday;
        var task = document.Tasks.Add("Progress"); task.Duration = new ProjectDuration(1, ProjectDurationUnit.Day, elapsed, estimated);
        task.ActualStart = Monday; task.ActualDuration = new ProjectDuration(.25m, ProjectDurationUnit.Day, elapsed, estimated);
        document.Recalculate(new ProjectScheduleOptions { CalculateAssignments = true });
        Assert.Equal(new ProjectDuration(.25m, ProjectDurationUnit.Day, elapsed, estimated), task.ActualDuration);
        Assert.Equal(new ProjectDuration(.75m, ProjectDurationUnit.Day, elapsed, estimated), task.RemainingDuration);
        using var copy = document.Clone();
        Assert.Equal(task.Duration, copy.Tasks[0].Duration); Assert.Equal(task.ActualDuration, copy.Tasks[0].ActualDuration); Assert.Equal(task.RemainingDuration, copy.Tasks[0].RemainingDuration);
    }

    [Theory]
    [InlineData("start")]
    [InlineData("stop")]
    [InlineData("finish")]
    public void CurrentStatusDoesNotOverrideFutureActualEvidence(string boundary) {
        using var document = ProjectDocument.Create(); document.Calendar = document.Calendars.AddStandardWorkingWeek(); document.Settings.StatusDate = Monday.AddHours(4);
        var task = document.Tasks.Add("Work"); task.ActualStart = boundary == "start" ? Monday.AddDays(1) : Monday; task.ActualCost = 100m;
        if (boundary == "stop") task.Stop = Monday.AddDays(1);
        if (boundary == "finish") task.ActualFinish = Monday.AddDays(1);
        var baseline = task.Baselines.Add(); baseline.Number = 0; baseline.Cost = 200m;
        Assert.Null(Assert.Single(document.AnalyzeEarnedValue().Tasks).ActualCost);
    }
}

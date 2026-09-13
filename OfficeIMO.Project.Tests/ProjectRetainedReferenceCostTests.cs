using System.Xml.Linq;

namespace OfficeIMO.Project.Tests;

public sealed class ProjectRetainedReferenceCostTests {
    private static readonly DateTime Monday = new(2026, 10, 5, 8, 0, 0);

    [Theory]
    [InlineData(ProjectFileFormat.Mpp8)]
    [InlineData(ProjectFileFormat.Mpp9)]
    [InlineData(ProjectFileFormat.Mpp12)]
    [InlineData(ProjectFileFormat.Mpp14)]
    [InlineData(ProjectFileFormat.Mpx4)]
    public void ResourceRemainingCostRoundTripsAcrossFormats(ProjectFileFormat format) {
        using var document = ProjectDocument.Create(); document.Calendar = document.Calendars.AddStandardWorkingWeek();
        document.Settings.StartDate = Monday;
        var resource = document.Resources.AddWork("Engineer"); resource.IsNull = false; resource.Cost = 125; resource.ActualCost = 25; resource.RemainingCost = 100;
        var options = new ProjectSaveOptions { Format = format, LossPolicy = format == ProjectFileFormat.Mpx4 ? OfficeConversionLossPolicy.Allow : OfficeConversionLossPolicy.Block };
        using var stream = new MemoryStream(); document.Save(stream, options); stream.Position = 0;
        using var copy = ProjectDocument.Load(stream); Assert.Equal(100m, copy.Resources.Single(r => r.Name == "Engineer").RemainingCost);
        copy.Resources.Single(r => r.Name == "Engineer").RemainingCost = 50;
        options.LossPolicy = OfficeConversionLossPolicy.Allow;
        using var edited = new MemoryStream(); copy.Save(edited, options); edited.Position = 0;
        using var reopened = ProjectDocument.Load(edited); Assert.Equal(50m, reopened.Resources.Single(r => r.Name == "Engineer").RemainingCost);
    }

    [Fact]
    public void MissingRateCannotEraseOnlyRecordedResourceRemainingCost() {
        using var document = ProjectDocument.Create(); document.Calendar = document.Calendars.AddStandardWorkingWeek(); document.Settings.StartDate = Monday;
        var task = document.Tasks.Add("Work"); task.Duration = ProjectDuration.WorkingDays(1);
        var resource = document.Resources.AddWork("Engineer"); resource.RemainingCost = 100;
        document.Assignments.Add(task, resource); string before = document.ToXml();
        var result = document.CalculateSchedule(new ProjectScheduleOptions { CalculateAssignments = true });
        Assert.Contains(result.Report.Diagnostics, d => d.Code == "PROJECT_COST_RECALCULATION_INCOMPLETE" && d.Location.EndsWith("/RemainingCost"));
        Assert.Throws<InvalidDataException>(() => document.ApplySchedule(result)); Assert.Equal(before, document.ToXml());
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ResourceRemainingCostCacheFollowsAssignments(bool removeAssignment) {
        using var source = ProjectDocument.Create(); source.Calendar = source.Calendars.AddStandardWorkingWeek(); source.Settings.StartDate = Monday;
        var task = source.Tasks.Add("Work"); task.Duration = ProjectDuration.WorkingDays(1);
        var resource = source.Resources.AddWork("Engineer"); resource.StandardRate = 100;
        source.Assignments.Add(task, resource);
        var xml = XDocument.Parse(source.ToXml()); var ns = xml.Root!.Name.Namespace;
        xml.Descendants(ns + "Resource").Single().Add(new XElement(ns + "RemainingCost", "99900"));
        using var document = ProjectDocument.Parse(xml.ToString());
        if (removeAssignment) document.Assignments.Remove(document.Assignments.Single());
        document.ApplySchedule(document.CalculateSchedule(new ProjectScheduleOptions { CalculateAssignments = true }));
        var saved = XDocument.Parse(document.ToXml());
        Assert.Equal(removeAssignment ? 0m : 80000m, (decimal)saved.Descendants(ns + "Resource").Single().Element(ns + "RemainingCost")!);
        Assert.False(document.AreWorkCostTotalsStale);
    }

    [Theory]
    [InlineData("project", true)]
    [InlineData("task", true)]
    [InlineData("resource", true)]
    [InlineData("base", true)]
    [InlineData("project", false)]
    [InlineData("task", false)]
    [InlineData("resource", false)]
    [InlineData("base", false)]
    public void ClearingDanglingCalendarReferenceSurvivesSave(string owner, bool preserveBytes) {
        string field = owner == "base" ? "BaseCalendarUID" : "CalendarUID";
        string inner = owner switch {
            "project" => "<CalendarUID>999</CalendarUID>",
            "task" => "<Tasks><Task><UID>1</UID><CalendarUID>999</CalendarUID></Task></Tasks>",
            "resource" => "<Resources><Resource><UID>1</UID><CalendarUID>999</CalendarUID></Resource></Resources>",
            _ => "<Calendars><Calendar><UID>1</UID><BaseCalendarUID>999</BaseCalendarUID></Calendar></Calendars>"
        };
        using var document = ProjectDocument.Parse("<Project xmlns=\"http://schemas.microsoft.com/project\"><SaveVersion>14</SaveVersion>" + inner + "</Project>");
        long revision = document.Revision;
        switch (owner) {
            case "project": document.Calendar = null; break;
            case "task": document.Tasks.Single().Calendar = null; break;
            case "resource": document.Resources.Single().Calendar = null; break;
            default: document.Calendars.Single().BaseCalendar = null; break;
        }
        Assert.True(document.Revision > revision);
        document.Validate().ThrowIfErrors();
        string saved = document.ToXml(new ProjectSaveOptions { PreserveUnchangedBytes = preserveBytes });
        Assert.DoesNotContain(XDocument.Parse(saved).Descendants(), e => e.Name.LocalName == field && e.Value == "999");
        using var copy = ProjectDocument.Parse(saved); copy.Validate().ThrowIfErrors();
    }

    [Theory]
    [InlineData(false, "span")]
    [InlineData(true, "span")]
    [InlineData(false, "stop")]
    [InlineData(true, "stop")]
    [InlineData(false, "missing-rate")]
    [InlineData(true, "missing-rate")]
    [InlineData(false, "finish")]
    [InlineData(true, "finish")]
    public void RetainedActualCostCurvesMustFitAssignmentDates(bool material, string boundary) {
        using var document = ProjectDocument.Create(); document.Calendar = document.Calendars.AddStandardWorkingWeek(); document.Settings.StartDate = Monday;
        var task = document.Tasks.Add("Work"); task.Duration = ProjectDuration.WorkingDays(1);
        var resource = material ? document.Resources.AddMaterial("Parts") : document.Resources.AddWork("Engineer"); resource.StandardRate = 100;
        var assignment = document.Assignments.Add(task, resource); assignment.ActualCost = 25;
        if (boundary == "missing-rate") resource.StandardRate = null;
        if (boundary == "finish") { assignment.ActualStart = Monday; assignment.ActualFinish = Monday.AddHours(1); assignment.Work = assignment.ActualWork = ProjectWork.Hours(1); }
        if (boundary == "stop") { assignment.ActualStart = Monday; assignment.Stop = Monday.AddHours(1); }
        var value = assignment.TimephasedData.Add(); value.Uid = assignment.Uid; value.Type = 6; value.Value = "2500";
        value.Start = boundary is "span" or "missing-rate" ? Monday.AddDays(1) : Monday.AddHours(2); value.Finish = value.Start;
        string before = document.ToXml();
        var result = document.CalculateSchedule(new ProjectScheduleOptions { CalculateAssignments = true });
        Assert.True(result.Report.HasErrors); Assert.Throws<InvalidDataException>(() => document.ApplySchedule(result));
        Assert.Equal(before, document.ToXml());
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(true, false)]
    [InlineData(false, true)]
    [InlineData(true, true)]
    public void FailedRateRecalculationCannotRetainInvalidCostCurves(bool material, bool inconsistentAmount) {
        using var document = ProjectDocument.Create(); document.Calendar = document.Calendars.AddStandardWorkingWeek(); document.Settings.StartDate = Monday;
        var task = document.Tasks.Add("Work"); task.Duration = ProjectDuration.WorkingDays(1);
        var resource = material ? document.Resources.AddMaterial("Parts") : document.Resources.AddWork("Engineer");
        var assignment = document.Assignments.Add(task, resource); assignment.ActualCost = 25;
        var value = assignment.TimephasedData.Add(); value.Uid = assignment.Uid; value.Type = 6;
        value.Start = inconsistentAmount ? Monday : Monday.AddDays(1); value.Finish = value.Start; value.Value = inconsistentAmount ? "5000" : "2500";
        string before = document.ToXml();
        var result = document.CalculateSchedule(new ProjectScheduleOptions { CalculateAssignments = true, RecalculateActualCosts = true });
        Assert.True(result.Report.HasErrors); Assert.Throws<InvalidDataException>(() => document.ApplySchedule(result)); Assert.Equal(before, document.ToXml());
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ValidRetainedActualCostsPreserveTheirTiming(bool material) {
        using var document = ProjectDocument.Create(); document.Calendar = document.Calendars.AddStandardWorkingWeek(); document.Settings.StartDate = Monday;
        var task = document.Tasks.Add("Work"); task.Duration = ProjectDuration.WorkingDays(1);
        var resource = material ? document.Resources.AddMaterial("Parts") : document.Resources.AddWork("Engineer"); resource.StandardRate = 100;
        var assignment = document.Assignments.Add(task, resource); assignment.ActualStart = Monday; assignment.ActualCost = 25;
        var value = assignment.TimephasedData.Add(); value.Uid = assignment.Uid; value.Type = 6; value.Start = Monday; value.Finish = Monday.AddHours(1); value.Value = "2500";
        var schedule = document.CalculateSchedule(new ProjectScheduleOptions { CalculateAssignments = true }); schedule.Report.ThrowIfErrors(); document.ApplySchedule(schedule);
        Assert.Same(value, assignment.TimephasedData.Single(v => v.Type == 6)); Assert.Equal(Monday, value.Start); Assert.Equal(Monday.AddHours(1), value.Finish);
        using var copy = document.Clone(); Assert.Equal(25m, copy.Assignments.Single().ActualCost);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void BackwardAlapDeadlineMatchesApplicationScheduling(bool assignments) {
        using var document = ProjectDocument.Create(); document.Calendar = document.Calendars.AddStandardWorkingWeek();
        document.Settings.ScheduleFromStart = false; document.Settings.FinishDate = Monday.AddDays(4).AddHours(9);
        var task = document.Tasks.Add("Deadline"); task.Duration = ProjectDuration.WorkingDays(1); task.ConstraintType = ProjectConstraintType.AsLateAsPossible;
        task.Deadline = Monday.AddDays(2).AddHours(9);
        var result = document.CalculateSchedule(new ProjectScheduleOptions { CalculateAssignments = assignments }); result.Report.ThrowIfErrors();
        Assert.Equal(task.Deadline, Assert.Single(result.Tasks).Finish);
        Assert.Equal(Monday.AddDays(2), Assert.Single(result.Tasks).Start);
    }
}

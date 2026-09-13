namespace OfficeIMO.Project.Tests;

public sealed class ProjectResourceCapacityTests {
    [Fact]
    public void CapacityAnalysisSplitsAtAvailabilityChangesAndExcludesMaterials() {
        using var document = ProjectDocument.Create(); document.Calendar = document.Calendars.AddStandardWorkingWeek();
        var monday = new DateTime(2026, 10, 5, 8, 0, 0); document.Settings.StartDate = monday;
        var resource = document.Resources.AddWork("Engineer"); resource.StandardRate = 100;
        var available = resource.AvailabilityPeriods.Add(); available.From = monday; available.Through = monday.AddHours(2).AddMinutes(-1); available.Units = ProjectUnits.Fraction(1);
        var reduced = resource.AvailabilityPeriods.Add(); reduced.From = monday.AddHours(2); reduced.Through = monday.AddDays(1); reduced.Units = ProjectUnits.Percent(50);
        for (int index = 0; index < 2; index++) {
            var task = document.Tasks.Add("Task " + index); task.Duration = ProjectDuration.WorkingMinutes(240);
            document.Assignments.Add(task, resource, ProjectUnits.Percent(50));
        }
        var schedule = document.CalculateSchedule(new ProjectScheduleOptions { CalculateAssignments = true }); schedule.Report.ThrowIfErrors();
        long revision = document.Revision;
        var result = document.AnalyzeResourceAllocation(schedule);
        Assert.Equal(revision, document.Revision); Assert.Equal(2, result.Intervals.Count);
        var overload = Assert.Single(result.Overallocations);
        Assert.Equal(monday.AddHours(2), overload.Start); Assert.Equal(monday.AddHours(4), overload.Finish);
        Assert.Equal(1m, overload.Units); Assert.Equal(.5m, overload.Capacity); Assert.Equal(.5m, overload.ExcessUnits);
        Assert.Equal(2, overload.AssignmentUids.Count);
        Assert.Throws<InvalidOperationException>(() => document.AnalyzeResourceAllocation(schedule, 1));
        resource.MaxUnits = ProjectUnits.Fraction(2);
        Assert.Throws<InvalidOperationException>(() => document.AnalyzeResourceAllocation(schedule));
    }
    internal static string Fixture(string name) => Path.Combine(AppContext.BaseDirectory, "Fixtures", "Project2024Advanced", name + ".xml");
    [Fact]
    public void ProducerRatesAndAvailabilityKeepUnitsBoundariesAndCurrencyAmounts() {
        using var document = ProjectDocument.Load(Fixture("rates"));
        document.Validate().ThrowIfErrors();
        var resource = document.Resources.GetByUid(1);
        Assert.Equal(ProjectCostAccrual.Prorated, resource.AccrueAt);
        Assert.Equal(2, resource.AvailabilityPeriods.Count); Assert.Equal(1m, resource.AvailabilityPeriods[0].Units!.Value.Value);
        Assert.Equal(new DateTime(2026, 10, 6, 23, 59, 0), resource.AvailabilityPeriods[0].Through);
        Assert.Equal(0.5m, resource.AvailabilityPeriods[1].Units!.Value.Value);
        Assert.Equal(2, resource.Rates.Count); Assert.Equal(100m, resource.Rates[0].StandardRate); Assert.Equal(200m, resource.Rates[1].StandardRate);
        Assert.Equal(resource.Rates[0].To, resource.Rates[1].From); Assert.Equal(25m, resource.Rates[1].CostPerUse);
        var material = document.Assignments.Single(a => a.Task?.Name == "Variable material");
        Assert.False(material.HasFixedRateUnits); Assert.Equal(2, material.MaterialRateScale); Assert.Equal(2m, material.Units!.Value.Value);
        using var bytes = new MemoryStream(); document.Save(bytes); Assert.Equal(File.ReadAllBytes(Fixture("rates")), bytes.ToArray());
        resource.Rates[1].StandardRate = 225m; resource.AvailabilityPeriods[1].Units = ProjectUnits.Percent(75);
        using var copy = document.Clone();
        Assert.Equal(225m, copy.Resources.GetByUid(1).Rates[1].StandardRate);
        Assert.Equal(0.75m, copy.Resources.GetByUid(1).AvailabilityPeriods[1].Units!.Value.Value);
        Assert.Contains(document.AnalyzeAssignments().Report.Diagnostics, d => d.Code == "PROJECT_RATE_ESTIMATE_UNSUPPORTED");
    }
    [Fact]
    public void AssignmentDelayUsesTenthsOfMinutesInXmlAndRequiresExplicitCalculationProfile() {
        using var document = ProjectDocument.Load(Fixture("calendars"));
        var assignment = document.Assignments.Single(a => a.Task?.Name == "Delayed assignment");
        Assert.Equal(120m, assignment.DelayMinutes); assignment.DelayMinutes = 90.5m;
        using var copy = document.Clone(); Assert.Equal(90.5m, copy.Assignments.GetByUid(assignment.Uid).DelayMinutes);
        Assert.Contains(copy.CalculateSchedule().Report.Diagnostics, d => d.Code == "PROJECT_ASSIGNMENT_SCHEDULING_PROFILE");
    }
    [Fact]
    public void CapacityAndRateOwnershipValidateBeforeOutputAndInvalidateScheduleResults() {
        using var document = ProjectDocument.Create(); document.Calendar = document.Calendars.AddStandardWorkingWeek();
        document.Settings.StartDate = new DateTime(2026, 10, 5, 8, 0, 0);
        var task = document.Tasks.Add("Task"); task.Duration = ProjectDuration.WorkingDays(1);
        var resource = document.Resources.AddWork("Engineer"); document.Assignments.Add(task, resource);
        var result = document.CalculateSchedule(); result.Report.ThrowIfErrors();
        var rate = resource.Rates.Add(); rate.From = document.Settings.StartDate; rate.To = rate.From.Value.AddDays(1); rate.StandardRate = 100;
        Assert.Throws<InvalidOperationException>(() => document.ApplySchedule(result));
        var overlap = resource.Rates.Add(); overlap.From = rate.From; overlap.To = rate.To; overlap.StandardRate = 120;
        Assert.Contains(document.Validate().Diagnostics, d => d.Code == "PROJECT_RATE_OVERLAP");
        using var bytes = new MemoryStream(); Assert.Throws<InvalidDataException>(() => document.Save(bytes)); Assert.Empty(bytes.ToArray());
        resource.Rates.Remove(overlap); Assert.Throws<InvalidOperationException>(() => overlap.StandardRate = 130);
    }
}

namespace OfficeIMO.Project.Tests;

public sealed class ProjectProgressBoundaryTests {
    private static readonly DateTime Monday = new(2026, 10, 5, 8, 0, 0);

    [Fact]
    public void InconsistentActualCostCurvesCannotProduceKnownHistoricalTotals() {
        using var document = ProjectDocument.Create(); document.Calendar = document.Calendars.AddStandardWorkingWeek();
        var task = document.Tasks.Add("Delivery"); var baseline = task.Baselines.Add(); baseline.Number = 0; baseline.Cost = 100m;
        var assignment = document.Assignments.Add(task, document.Resources.AddWork("Engineer")); assignment.ActualCost = 100m;
        var curve = assignment.TimephasedData.Add(); curve.Uid = assignment.Uid; curve.Type = 6; curve.Start = Monday; curve.Finish = Monday.AddHours(1); curve.Value = "5000";
        var result = document.AnalyzeEarnedValue(statusDate: Monday.AddHours(2));
        Assert.Null(Assert.Single(result.Tasks).ActualCost);
        Assert.Contains(result.Report.Diagnostics, d => d.Message.Contains("Actual cost curves differ"));
    }

    [Theory]
    [InlineData("task")]
    [InlineData("assignment")]
    [InlineData("fixed")]
    public void PlannedFinishCannotDateAnUndocumentedActualCost(string owner) {
        using var document = ProjectDocument.Create(); document.Calendar = document.Calendars.AddStandardWorkingWeek();
        var task = document.Tasks.Add("Late work"); task.Finish = Monday.AddHours(9); task.ActualStart = Monday.AddDays(1); task.ActualCost = 100m;
        var baseline = task.Baselines.Add(); baseline.Number = 0; baseline.Cost = 100m;
        baseline.Start = Monday; baseline.Finish = Monday.AddHours(9);
        if (owner != "task") {
            var assignment = document.Assignments.Add(task, document.Resources.AddWork("Engineer"));
            assignment.Finish = task.Finish; assignment.ActualStart = task.ActualStart;
            assignment.ActualCost = owner == "fixed" ? 0m : 100m;
        }
        var result = document.AnalyzeEarnedValue(statusDate: Monday.AddHours(10));
        Assert.Null(Assert.Single(result.Tasks).ActualCost);
        Assert.Contains(result.Report.Diagnostics, d => d.Code == "PROJECT_EARNED_VALUE_INCOMPLETE");
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(true, false)]
    [InlineData(false, true)]
    [InlineData(true, true)]
    public void ActualCostTimingFollowsTheSelectedRecalculationPolicy(bool recalculate, bool omitScalar) {
        using var document = ProjectDocument.Create(); document.Calendar = document.Calendars.AddStandardWorkingWeek();
        document.Settings.StartDate = Monday;
        var task = document.Tasks.Add("Work"); task.Duration = ProjectDuration.WorkingHours(1);
        task.ActualStart = Monday.AddDays(1); task.ActualFinish = Monday.AddDays(1).AddHours(1);
        task.ActualDuration = ProjectDuration.WorkingHours(1);
        var resource = document.Resources.AddWork("Engineer"); resource.StandardRate = 100m;
        var assignment = document.Assignments.Add(task, resource);
        assignment.Work = assignment.ActualWork = ProjectWork.Hours(1); assignment.ActualStart = task.ActualStart; assignment.ActualFinish = task.ActualFinish;
        if (!omitScalar) assignment.ActualCost = 100m;
        var old = assignment.TimephasedData.Add(); old.Uid = assignment.Uid; old.Type = 6; old.Start = Monday; old.Finish = Monday.AddHours(1); old.Value = "10000";
        string before = document.ToXml();
        var result = document.CalculateSchedule(new ProjectScheduleOptions { CalculateAssignments = true, RecalculateActualCosts = recalculate });
        if (!recalculate) {
            Assert.True(result.Report.HasErrors);
            Assert.Throws<InvalidDataException>(() => document.ApplySchedule(result));
            Assert.Equal(before, document.ToXml());
            return;
        }
        result.Report.ThrowIfErrors();
        var expected = Monday.AddDays(1);
        Assert.Equal(expected, Assert.Single(result.Assignments.Single().Costs, c => c.IsActual).Start);
        document.ApplySchedule(result);
        Assert.Equal(expected, Assert.Single(assignment.TimephasedData, c => c.Type == 6).Start);
        Assert.Equal(100m, assignment.ActualCost);
        using var copy = document.Clone();
        Assert.Equal(expected, Assert.Single(copy.Assignments.Single().TimephasedData, c => c.Type == 6).Start);
    }

    [Fact]
    public void CompletedMilestonesAndTheirZeroDurationSummaryStayCompleted() {
        using var document = ProjectDocument.Create(); document.Calendar = document.Calendars.AddStandardWorkingWeek(); document.Settings.StartDate = Monday;
        var summary = document.Tasks.AddSummary("Acceptance"); var task = summary.Children.Add("Accepted");
        task.Duration = ProjectDuration.WorkingMinutes(0); task.ActualStart = task.ActualFinish = Monday; task.PercentComplete = 100;
        var result = document.CalculateSchedule(new ProjectScheduleOptions { CalculateAssignments = true }); result.Report.ThrowIfErrors();
        Assert.All(result.Tasks, t => Assert.Equal(100, t.Calculation!.PercentComplete));
        document.ApplySchedule(result); Assert.Equal(100, task.PercentComplete); Assert.Equal(100, summary.PercentComplete);
        using var copy = document.Clone(); Assert.Equal(100, copy.Tasks.GetByUid(task.Uid).PercentComplete);
    }
}

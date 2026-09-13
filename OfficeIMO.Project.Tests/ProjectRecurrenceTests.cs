namespace OfficeIMO.Project.Tests;

public sealed class ProjectRecurrenceTests {
    [Fact]
    public void FiniteRulesRespectWeekAnchorsLeapYearsAndMissingMonthDays() {
        var weekly = new ProjectRecurrenceRule { Start = new DateTime(2026, 9, 11, 8, 0, 0), Frequency = ProjectRecurrenceFrequency.Weekly, Count = 3 };
        weekly.DaysOfWeek.Add(DayOfWeek.Monday);
        Assert.Equal(new[] { new DateTime(2026, 9, 14, 8, 0, 0), new DateTime(2026, 9, 21, 8, 0, 0), new DateTime(2026, 9, 28, 8, 0, 0) }, weekly.Expand());
        weekly.Interval = 2;
        Assert.Equal(new DateTime(2026, 9, 21, 8, 0, 0), weekly.Expand()[0]);
        var monthly = new ProjectRecurrenceRule { Start = new DateTime(2026, 1, 31, 9, 0, 0), Frequency = ProjectRecurrenceFrequency.Monthly, Count = 3 };
        Assert.Equal(new[] { monthly.Start, new DateTime(2026, 3, 31, 9, 0, 0), new DateTime(2026, 5, 31, 9, 0, 0) }, monthly.Expand());
        monthly.WeekDay = DayOfWeek.Monday; monthly.WeekOrdinal = -1;
        Assert.Equal(new DateTime(2026, 2, 23, 9, 0, 0), monthly.Expand()[0]);
        var yearly = new ProjectRecurrenceRule { Start = new DateTime(2024, 2, 29, 8, 0, 0), Frequency = ProjectRecurrenceFrequency.Yearly, Count = 2 };
        Assert.Equal(new[] { yearly.Start, new DateTime(2028, 2, 29, 8, 0, 0) }, yearly.Expand());
        Assert.Throws<InvalidOperationException>(() => yearly.Expand(maxCalendarDays: 365));
        Assert.Throws<OperationCanceledException>(() => yearly.Expand(cancellationToken: new CancellationToken(true)));
        yearly.Month = 13; Assert.Throws<ArgumentException>(() => yearly.Expand());
    }
    [Fact]
    public void ProducerExpandedOccurrencesRetainMarkersAndStoredDates() {
        string path = ProjectResourceCapacityTests.Fixture("recurrence");
        using var project = ProjectDocument.Load(path);
        var summary = project.Tasks.GetByUid(1);
        Assert.True(summary.IsRecurring); Assert.Equal(3, summary.Children.Count);
        Assert.All(summary.Children, child => Assert.True(child.IsRecurring));
        using var bytes = new MemoryStream(); project.Save(bytes);
        Assert.Equal(File.ReadAllBytes(path), bytes.ToArray());
        var schedule = project.CalculateSchedule(); schedule.Report.ThrowIfErrors();
        Assert.Contains(schedule.Report.Diagnostics, d => d.Code == "PROJECT_EXPANDED_RECURRENCE");
        foreach (var child in summary.Children) {
            var result = schedule.Tasks.Single(t => t.TaskUid == child.Uid);
            Assert.Equal(child.Start, result.Start); Assert.Equal(child.Finish, result.Finish);
        }
        summary.Children[0].IsRecurring = false;
        using var copy = project.Clone();
        Assert.False(copy.Tasks.GetByUid(2).IsRecurring);
    }
    [Fact]
    public void ExplicitSeriesRejectsInvalidInputBeforeMutationAndRoundTrips() {
        using var project = ProjectDocument.Load(ProjectResourceCapacityTests.Fixture("recurrence"));
        var first = new DateTime(2026, 10, 5, 8, 0, 0);
        long revision = project.Revision;
        Assert.Throws<ArgumentException>(() => project.AddRecurringTask("Review", new[] { first, first }, ProjectDuration.WorkingDays(1)));
        Assert.Throws<InvalidOperationException>(() => project.AddRecurringTask("Review", new[] { first, first.AddDays(7) }, ProjectDuration.WorkingDays(1), maxOccurrences: 1));
        Assert.Throws<OperationCanceledException>(() => project.AddRecurringTask("Review", new[] { first }, ProjectDuration.WorkingDays(1), cancellationToken: new CancellationToken(true)));
        Assert.Equal(revision, project.Revision);
        var summary = project.AddRecurringTask("Review", new[] { first, first.AddDays(7) }, ProjectDuration.WorkingDays(1));
        var schedule = project.CalculateSchedule(); schedule.Report.ThrowIfErrors();
        Assert.Equal(first, schedule.Tasks.Single(t => t.TaskUid == summary.Children[0].Uid).Start);
        Assert.Equal(first.AddDays(7), schedule.Tasks.Single(t => t.TaskUid == summary.Children[1].Uid).Start);
        using var copy = project.Clone(new ProjectSaveOptions { LossPolicy = OfficeConversionLossPolicy.Allow });
        Assert.True(copy.Tasks.GetByUid(summary.Uid).IsRecurring);
        Assert.Equal(first.AddDays(7), copy.Tasks.GetByUid(summary.Children[1].Uid).ConstraintDate);
    }
}

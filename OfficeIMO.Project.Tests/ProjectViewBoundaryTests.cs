namespace OfficeIMO.Project.Tests;

public sealed class ProjectViewBoundaryTests {
    private static readonly DateTime Wednesday = new(2026, 10, 7, 8, 0, 0);

    [Theory]
    [InlineData(ProjectViewKind.ResourceUsage, false)]
    [InlineData(ProjectViewKind.ResourceUsage, true)]
    [InlineData(ProjectViewKind.ResourceHistogram, false)]
    [InlineData(ProjectViewKind.ResourceHistogram, true)]
    public void ResourceRowLimitsAllowMultipleContributingTasks(ProjectViewKind kind, bool selectTasks) {
        using var document = ProjectDocument.Create(); document.Calendar = document.Calendars.AddStandardWorkingWeek(); document.Settings.StartDate = Wednesday;
        var resource = document.Resources.AddWork("Engineer"); resource.StandardRate = 100m;
        var first = document.Tasks.Add("First"); first.Duration = ProjectDuration.WorkingHours(1);
        var second = document.Tasks.Add("Second"); second.Duration = ProjectDuration.WorkingHours(2);
        document.Assignments.Add(first, resource); document.Assignments.Add(second, resource);
        var schedule = document.CalculateSchedule(new ProjectScheduleOptions { CalculateAssignments = true }); schedule.Report.ThrowIfErrors();
        var options = new ProjectViewOptions { Kind = kind, MaxRows = 1, TaskUids = selectTasks ? new[] { first.Uid, second.Uid } : null };
        var row = Assert.Single(document.CreateView(schedule, options).Rows);
        Assert.Equal(3m, row.WorkHours); Assert.Equal(300m, row.Cost);
        document.Resources.AddWork("Another");
        schedule = document.CalculateSchedule(new ProjectScheduleOptions { CalculateAssignments = true }); schedule.Report.ThrowIfErrors();
        Assert.Throws<InvalidOperationException>(() => document.CreateView(schedule, options));
    }

    [Fact]
    public void TaskRowLimitsAllowMultipleContributingResourceFilters() {
        using var document = ProjectDocument.Create(); document.Calendar = document.Calendars.AddStandardWorkingWeek(); document.Settings.StartDate = Wednesday;
        var task = document.Tasks.Add("Delivery"); task.Duration = ProjectDuration.WorkingHours(1);
        var first = document.Resources.AddWork("First"); var second = document.Resources.AddWork("Second");
        document.Assignments.Add(task, first); document.Assignments.Add(task, second);
        var schedule = document.CalculateSchedule(new ProjectScheduleOptions { CalculateAssignments = true }); schedule.Report.ThrowIfErrors();
        var row = Assert.Single(document.CreateView(schedule, new ProjectViewOptions { Kind = ProjectViewKind.TaskUsage, MaxRows = 1, ResourceUids = new[] { first.Uid, second.Uid } }).Rows);
        Assert.Equal(2m, row.WorkHours);
    }
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void MonthlyBucketStartUsesTheCalendarMonthUnlessExplicitlyClipped(bool explicitStart) {
        using var document = ProjectDocument.Create(); document.Calendar = document.Calendars.AddStandardWorkingWeek(); document.Settings.StartDate = Wednesday;
        document.Tasks.Add("Delivery").Duration = ProjectDuration.WorkingDays(30);
        var view = document.CreateView(document.CalculateSchedule(), new ProjectViewOptions { Timescale = ProjectViewTimescale.Month, Start = explicitStart ? Wednesday : null });
        Assert.Equal(explicitStart ? Wednesday : new DateTime(2026, 10, 1), view.Buckets[0].Start);
        Assert.Equal(new DateTime(2026, 11, 1), view.Buckets[0].Finish);
    }
    [Theory]
    [InlineData(ProjectViewKind.ResourceUsage)]
    [InlineData(ProjectViewKind.ResourceHistogram)]
    public void ResourceViewsHonorSelectedTasks(ProjectViewKind kind) {
        using var document = ProjectDocument.Create(); document.Calendar = document.Calendars.AddStandardWorkingWeek(); document.Settings.StartDate = Wednesday;
        var resource = document.Resources.AddWork("Engineer"); resource.StandardRate = 100m;
        var first = document.Tasks.Add("Selected"); first.Duration = ProjectDuration.WorkingHours(1);
        var other = document.Tasks.Add("Other"); other.Duration = ProjectDuration.WorkingDays(10);
        document.Assignments.Add(first, resource); document.Assignments.Add(other, resource);
        var schedule = document.CalculateSchedule(new ProjectScheduleOptions { CalculateAssignments = true }); schedule.Report.ThrowIfErrors();
        var view = document.CreateView(schedule, new ProjectViewOptions { Kind = kind, TaskUids = new[] { first.Uid }, Timescale = ProjectViewTimescale.Day, MaxBuckets = 1 });
        var row = Assert.Single(view.Rows); Assert.Equal(1m, row.WorkHours); Assert.Equal(100m, row.Cost);
        Assert.Equal(Wednesday.AddHours(1), row.Finish); Assert.Single(view.Buckets);
    }
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void WeeklyBucketStartIsInferredAtMondayOrExplicitlyClipped(bool explicitStart) {
        using var document = ProjectDocument.Create(); document.Calendar = document.Calendars.AddStandardWorkingWeek(); document.Settings.StartDate = Wednesday;
        document.Tasks.Add("Delivery").Duration = ProjectDuration.WorkingDays(8);
        var view = document.CreateView(document.CalculateSchedule(), new ProjectViewOptions { Timescale = ProjectViewTimescale.Week, Start = explicitStart ? Wednesday : null });
        Assert.Equal(explicitStart ? Wednesday : Wednesday.Date.AddDays(-2), view.Buckets[0].Start);
        Assert.Equal(DayOfWeek.Monday, view.Buckets[0].Finish.DayOfWeek);
    }
}

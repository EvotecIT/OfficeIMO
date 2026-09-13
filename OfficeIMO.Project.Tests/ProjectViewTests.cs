using OfficeIMO.Drawing;

namespace OfficeIMO.Project.Tests;

public sealed class ProjectViewTests {
    [Fact]
    public void ContentSizingPreservesFixedPageOptOutAndWrapsLongTimelineLabels() {
        using var project = Create();
        project.Tasks.Add("Design, implement, verify and document the complete shared service integration and operational handover").Duration = ProjectDuration.WorkingHours(1);
        var schedule = project.CalculateSchedule();
        var compact = project.CreateView(schedule, new ProjectViewOptions { Kind = ProjectViewKind.Timeline, PageWidth = 356, Columns = new[] { ProjectViewColumn.Name } });
        var fixedSize = project.CreateView(schedule, new ProjectViewOptions { Kind = ProjectViewKind.Timeline, PageWidth = 356, Columns = new[] { ProjectViewColumn.Name }, FitPageHeightToContent = false });
        Assert.True(Assert.Single(compact.Render()).Drawing.Height < compact.PageHeight);
        Assert.Equal(fixedSize.PageHeight, Assert.Single(fixedSize.Render()).Drawing.Height);
        string svg = OfficeDrawingSvgExporter.ToSvg(Assert.Single(compact.Render()).Drawing);
        Assert.Contains("handover", svg);
    }

    [Fact]
    public void NetworkWrapsLongNamesAndKeepsAllNodesAcrossMeasuredPages() {
        using var project = Create();
        ProjectTask? previous = null;
        for (int i = 0; i < 8; i++) {
            var task = project.Tasks.Add("Design, implement and verify the shared service integration and operational handover " + i);
            task.Duration = ProjectDuration.WorkingHours(1);
            if (previous != null) project.Dependencies.Add(previous, task);
            previous = task;
        }
        var view = project.CreateView(project.CalculateSchedule(), new ProjectViewOptions { Kind = ProjectViewKind.Network, PageWidth = 420, PageHeight = 420 });
        var pages = view.Render();
        Assert.True(pages.Count > 1); Assert.Equal(8, pages.Sum(page => page.RowCount));
        foreach (var page in pages) Assert.Contains("handover", OfficeDrawingSvgExporter.ToSvg(page.Drawing));
        Assert.Contains("FS", string.Join("", pages.Select(page => OfficeDrawingSvgExporter.ToSvg(page.Drawing))));
    }

    [Theory]
    [InlineData(ProjectViewKind.Gantt)]
    [InlineData(ProjectViewKind.Timeline)]
    public void InferredRangeIncludesTerminalMilestonesButExplicitFinishRemainsExclusive(ProjectViewKind kind) {
        using var project = Create();
        var task = project.Tasks.Add("Launch"); task.Duration = ProjectDuration.WorkingMinutes(0);
        var schedule = project.CalculateSchedule(); schedule.Report.ThrowIfErrors();
        var view = project.CreateView(schedule, new ProjectViewOptions { Kind = kind });
        Assert.True(view.Buckets.Last().Finish > view.Rows.Single().Finish);
        var explicitView = project.CreateView(schedule, new ProjectViewOptions {
            Kind = kind, Start = project.Settings.StartDate!.Value.Date, Finish = view.Rows.Single().Finish
        });
        Assert.Equal(view.Rows.Single().Finish, explicitView.Buckets.Last().Finish);
    }

    [Theory]
    [InlineData(ProjectViewKind.ResourceUsage)]
    [InlineData(ProjectViewKind.ResourceHistogram)]
    public void ResourceNameSelectionBoundsTheVisibleRange(ProjectViewKind kind) {
        using var project = Create();
        var selected = project.Resources.AddWork("Selected"); var other = project.Resources.AddWork("Other");
        var shortTask = project.Tasks.Add("Short"); shortTask.Duration = ProjectDuration.WorkingHours(8);
        var longTask = project.Tasks.Add("Long"); longTask.Duration = ProjectDuration.WorkingDays(20);
        project.Assignments.Add(shortTask, selected); project.Assignments.Add(longTask, other);
        var schedule = project.CalculateSchedule(new ProjectScheduleOptions { CalculateAssignments = true }); schedule.Report.ThrowIfErrors();
        var view = project.CreateView(schedule, new ProjectViewOptions { Kind = kind, NameContains = "selected", Timescale = ProjectViewTimescale.Day, MaxBuckets = 1 });
        Assert.Equal(selected.Uid, Assert.Single(view.Rows).Uid);
        Assert.Equal(new DateTime(2026, 10, 5, 17, 0, 0), Assert.Single(view.Buckets).Finish);
    }

    [Fact]
    public void ResourceSelectionUsesOnlyItsAssignmentTimeRange() {
        using var project = Create();
        var selected = project.Resources.AddWork("Selected"); var other = project.Resources.AddWork("Other");
        var shortTask = project.Tasks.Add("Short"); shortTask.Duration = ProjectDuration.WorkingHours(8);
        var longTask = project.Tasks.Add("Long"); longTask.Duration = ProjectDuration.WorkingDays(20);
        project.Assignments.Add(shortTask, selected); project.Assignments.Add(longTask, other);
        var schedule = project.CalculateSchedule(new ProjectScheduleOptions { CalculateAssignments = true }); schedule.Report.ThrowIfErrors();
        var view = project.CreateView(schedule, new ProjectViewOptions { Kind = ProjectViewKind.ResourceUsage, ResourceUids = new[] { selected.Uid }, Timescale = ProjectViewTimescale.Day });
        Assert.Equal(8m, Assert.Single(view.Rows).WorkHours);
        Assert.Equal(new DateTime(2026, 10, 5, 17, 0, 0), Assert.Single(view.Buckets).Finish);
    }

    [Theory]
    [InlineData(DateTimeKind.Utc)]
    [InlineData(DateTimeKind.Local)]
    public void VisibleDatesCannotSilentlyChangeTimezoneKind(DateTimeKind kind) {
        using var project = Create(); var task = project.Tasks.Add("Task"); task.Duration = ProjectDuration.WorkingHours(1);
        Assert.Throws<ArgumentException>(() => project.CreateView(project.CalculateSchedule(), new ProjectViewOptions { Start = new DateTime(2026, 10, 5, 8, 0, 0, kind) }));
    }

    [Fact]
    public void UsageBucketsPreserveEffortAndReportIsAnIndependentSnapshot() {
        using var project = Create();
        var task = project.Tasks.Add("Zażółć — engineering"); task.Duration = ProjectDuration.WorkingMinutes(960);
        var resource = project.Resources.AddWork("Engineer"); project.Assignments.Add(task, resource);
        var schedule = project.CalculateSchedule(new ProjectScheduleOptions { CalculateAssignments = true }); schedule.Report.ThrowIfErrors();
        var options = new ProjectViewOptions { Kind = ProjectViewKind.TaskUsage, Timescale = ProjectViewTimescale.Day };
        var view = project.CreateView(schedule, options);
        var row = Assert.Single(view.Rows);
        Assert.Equal(16m, row.WorkHours); Assert.Equal(new[] { 8m, 8m }, row.BucketWorkHours);
        Assert.Equal(16m, project.CreateView(schedule, new ProjectViewOptions { Kind = ProjectViewKind.ResourceUsage }).Rows.Single().BucketWorkHours.Sum());
        options.PageWidth = 1; task.Name = "Later edit";
        Assert.Equal("Zażółć — engineering", row.Name);
        Assert.Contains("Zażółć", OfficeDrawingSvgExporter.ToSvg(Assert.Single(view.Render()).Drawing));
        Assert.Throws<InvalidOperationException>(() => project.CreateView(schedule));
    }

    [Fact]
    public void SelectionClippingAndLimitsAreExplicit() {
        using var project = Create();
        var task = project.Tasks.Add("Work"); task.Duration = ProjectDuration.WorkingMinutes(960);
        var resource = project.Resources.AddWork("Engineer"); project.Assignments.Add(task, resource);
        var schedule = project.CalculateSchedule(new ProjectScheduleOptions { CalculateAssignments = true }); schedule.Report.ThrowIfErrors();
        var view = project.CreateView(schedule, new ProjectViewOptions { Kind = ProjectViewKind.TaskUsage,
            Start = new DateTime(2026, 10, 5, 10, 0, 0), Finish = new DateTime(2026, 10, 5, 12, 0, 0) });
        Assert.Equal(2m, view.Rows.Single().BucketWorkHours.Single());
        Assert.Equal(16m, view.Rows.Single().WorkHours);
        Assert.Empty(project.CreateView(schedule, new ProjectViewOptions { TaskUids = Array.Empty<int>() }).Rows);
        Assert.Throws<ArgumentException>(() => project.CreateView(schedule, new ProjectViewOptions { TaskUids = new[] { 999 } }));
        Assert.Throws<InvalidOperationException>(() => project.CreateView(schedule, new ProjectViewOptions { Timescale = ProjectViewTimescale.Day, MaxBuckets = 1 }));
        Assert.Throws<ArgumentOutOfRangeException>(() => project.CreateView(schedule, new ProjectViewOptions { PageWidth = double.NaN }));
        using var cancel = new CancellationTokenSource(); cancel.Cancel();
        Assert.Throws<OperationCanceledException>(() => project.CreateView(schedule, cancellationToken: cancel.Token));
    }

    [Theory]
    [InlineData(ProjectViewKind.Gantt)]
    [InlineData(ProjectViewKind.TaskUsage)]
    [InlineData(ProjectViewKind.ResourceUsage)]
    [InlineData(ProjectViewKind.ResourceHistogram)]
    [InlineData(ProjectViewKind.Network)]
    [InlineData(ProjectViewKind.Timeline)]
    [InlineData(ProjectViewKind.Table)]
    public void LayoutsPaginateWithoutDroppingRows(ProjectViewKind kind) {
        using var project = Create();
        for (int i = 0; i < 20; i++) {
            var task = project.Tasks.Add("Task " + i); task.Duration = ProjectDuration.WorkingMinutes(480);
            project.Assignments.Add(task, project.Resources.AddWork("Resource " + i));
        }
        var schedule = project.CalculateSchedule(new ProjectScheduleOptions { CalculateAssignments = true }); schedule.Report.ThrowIfErrors();
        var view = project.CreateView(schedule, new ProjectViewOptions { Kind = kind, PageHeight = 300 });
        var pages = view.Render();
        Assert.True(pages.Count > 1); Assert.Equal(20, pages.Sum(p => p.RowCount));
        Assert.All(pages, p => Assert.Contains("<svg", OfficeDrawingSvgExporter.ToSvg(p.Drawing)));
        var limited = project.CreateView(schedule, new ProjectViewOptions { Kind = kind, PageHeight = 300, MaxPages = 1 });
        Assert.Throws<InvalidOperationException>(() => limited.Render());
    }

    private static ProjectDocument Create() {
        var project = ProjectDocument.Create(); project.Calendar = project.Calendars.AddStandardWorkingWeek();
        project.Settings.StartDate = new DateTime(2026, 10, 5, 8, 0, 0); return project;
    }
}

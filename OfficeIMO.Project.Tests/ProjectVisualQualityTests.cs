using OfficeIMO.Drawing;

namespace OfficeIMO.Project.Tests;

public sealed class ProjectVisualQualityTests {
    [Fact]
    public void DependencyNetworkPlacesParallelTasksTogetherAndRetainsSourceMapping() {
        using var project = ProjectDocument.Create(); project.Calendar = project.Calendars.AddStandardWorkingWeek();
        project.Settings.StartDate = new DateTime(2026, 10, 5, 8, 0, 0);
        var start = project.Tasks.Add("Start"); var left = project.Tasks.Add("Parallel A");
        var right = project.Tasks.Add("Parallel B"); var finish = project.Tasks.Add("Finish");
        foreach (var task in project.Tasks) task.Duration = ProjectDuration.WorkingHours(8);
        project.Dependencies.Add(start, left); project.Dependencies.Add(start, right);
        project.Dependencies.Add(left, finish); project.Dependencies.Add(right, finish);
        var view = project.CreateView(project.CalculateSchedule(), new ProjectViewOptions { Kind = ProjectViewKind.Network, PageWidth = 1200, PageHeight = 900 });
        var page = Assert.Single(view.Render());
        var labels = page.Drawing.Elements.OfType<OfficeDrawingText>().ToArray();
        OfficeDrawingText Label(string name) => Assert.Single(labels, text => text.Text == name);
        Assert.True(Label("Start").X < Label("Parallel A").X);
        Assert.Equal(Label("Parallel A").X, Label("Parallel B").X);
        Assert.NotEqual(Label("Parallel A").Y, Label("Parallel B").Y);
        Assert.True(Label("Finish").X > Label("Parallel B").X);
        Assert.Equal(Enumerable.Range(0, 4), page.RowIndices.OrderBy(index => index));
        Assert.Equal(4, page.RowCount);
    }

    [Fact]
    public void ProgressAndStatusDateAreRenderedWithoutMutatingTheSnapshot() {
        using var project = ProjectDocument.Create(); project.Calendar = project.Calendars.AddStandardWorkingWeek();
        project.Settings.StartDate = new DateTime(2026, 10, 5, 8, 0, 0);
        var task = project.Tasks.Add("Half complete"); task.Duration = ProjectDuration.WorkingHours(8);
        task.ActualDuration = ProjectDuration.WorkingHours(4); task.RemainingDuration = ProjectDuration.WorkingHours(4);
        var schedule = project.CalculateSchedule(new ProjectScheduleOptions { CalculateAssignments = true }); schedule.Report.ThrowIfErrors();
        var withProgress = project.CreateView(schedule, new ProjectViewOptions { StatusDate = new DateTime(2026, 10, 5, 12, 0, 0) });
        var withoutProgress = project.CreateView(schedule, new ProjectViewOptions { ShowProgress = false });
        Assert.Equal(50, withProgress.Rows.Single().PercentComplete);
        var drawing = Assert.Single(withProgress.Render()).Drawing;
        Assert.True(drawing.Shapes.Count > Assert.Single(withoutProgress.Render()).Drawing.Shapes.Count);
        Assert.Contains(drawing.Shapes, shape => shape.Shape.StrokeDashStyle == OfficeStrokeDashStyle.Dash);
        Assert.Equal(50, withProgress.Rows.Single().PercentComplete);
    }
}

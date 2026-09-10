using OfficeIMO.Project.Fluent;

namespace OfficeIMO.Project.Tests;

public class ModelContracts {
    [Fact]
    public void DeepCloneRequiresExplicitLimitsAndKeepsTheWholeHierarchy() {
        using var document = ProjectDocument.Create();
        var task = document.Tasks.Add("Root");
        for (int i = 1; i < 130; i++) task = task.Children.Add("Level " + i);
        Assert.Throws<InvalidDataException>(() => document.Clone());
        using var clone = document.Clone(loadOptions: new ProjectLoadOptions { MaxOutlineDepth = 130 });
        Assert.Equal(130, clone.AllTasks.Count());
        Assert.Equal(task.Uid, clone.AllTasks.Last().Uid);
        Assert.Equal(129, clone.AllTasks.Last().Parent!.Uid);
    }

    [Fact]
    public void NormalAndFluentAuthorTheSameProjectAndAllowMixedEdits() {
        using var normal = ProjectDocument.Create();
        normal.Name = "Delivery";
        normal.Settings.StartDate = new DateTime(2026, 10, 5, 8, 0, 0);
        normal.Settings.ScheduleFromStart = true;
        normal.Calendar = normal.Calendars.AddStandardWorkingWeek();
        var resource = normal.Resources.AddWork("Engineer");
        resource.StandardRate = 125;
        var summary = normal.Tasks.AddSummary("Delivery");
        var design = summary.Children.Add("Design");
        design.Duration = ProjectDuration.WorkingDays(3);
        var build = summary.Children.Add("Build");
        build.Duration = ProjectDuration.WorkingDays(5);
        normal.Dependencies.Add(design, build);
        normal.Assignments.Add(build, resource, ProjectUnits.Percent(50));

        using var fluent = ProjectDocument.Create().AsFluent()
            .Info(i => i.Name("Delivery"))
            .StartsOn(new DateTime(2026, 10, 5, 8, 0, 0))
            .StandardWorkingWeek()
            .Resource("engineer", r => r.Name("Engineer").Work().StandardRate(125))
            .Summary("delivery", "Delivery", tasks => tasks
                .Task("design", "Design", t => t.Duration(ProjectDuration.WorkingDays(3)))
                .Task("build", "Build", t => t.Duration(ProjectDuration.WorkingDays(5)).After("design").Assign("engineer", ProjectUnits.Percent(50))))
            .End();
        Assert.Equal(normal.ToXml(), fluent.ToXml());
        fluent.Tasks.GetByUid(build.Uid).Notes = "Mixed normal/fluent edit";
        using var copy = ProjectDocument.Parse(fluent.ToXml());
        Assert.Equal("Mixed normal/fluent edit", copy.Tasks.GetByUid(build.Uid).Notes);
        Assert.Equal(0.5m, copy.Assignments.Single().Units!.Value.Value);
    }

    [Fact]
    public void StableIdentitiesSurviveMoveCloneAndCascadeRemoval() {
        using var document = ProjectDocument.Create();
        var summary = document.Tasks.AddSummary("Phase");
        var child = summary.Children.Add("Repeated name");
        var sibling = document.Tasks.Add("Repeated name");
        var resource = document.Resources.AddWork("Engineer");
        document.Assignments.Add(child, resource);
        document.Dependencies.Add(child, sibling);
        int uid = child.Uid;
        Assert.Throws<InvalidOperationException>(() => document.Tasks.Remove(summary));
        Assert.Throws<ArgumentException>(() => summary.MoveTo(child));
        child.MoveTo(null);
        Assert.Equal(uid, child.Uid);
        Assert.Same(child, document.Tasks.GetByUid(uid));
        using var clone = document.Clone();
        clone.Tasks.GetByUid(uid).Name = "Clone only";
        Assert.Equal("Repeated name", child.Name);
        Assert.True(document.Tasks.Remove(child, ProjectRemovalMode.Cascade));
        Assert.Empty(document.Assignments);
        Assert.Empty(document.Dependencies);
        Assert.Throws<InvalidOperationException>(() => child.Name = "Removed");
        Assert.Throws<System.Collections.Generic.KeyNotFoundException>(() => document.Tasks.GetByUid(uid));
        Assert.True(document.Tasks.Add("Next").Uid > uid);
    }

    [Fact]
    public void RemovedOwnersInvalidateNestedCollectionsAndPreviouslyReturnedChildren() {
        using var document = ProjectDocument.Create();
        var task = document.Tasks.Add("Task");
        var baseline = task.Baselines.Add();
        baseline.Number = 0;
        var interval = baseline.TimephasedData.Add();
        document.Tasks.Remove(task);
        Assert.Throws<InvalidOperationException>(() => baseline.Cost = 2);
        Assert.Throws<InvalidOperationException>(() => interval.Value = "PT8H0M0S");
        Assert.Throws<InvalidOperationException>(() => baseline.TimephasedData.Add());
        Assert.Throws<InvalidOperationException>(() => task.Baselines.Add());
        var calendar = document.Calendars.AddStandardWorkingWeek();
        var weekday = calendar.WeekDays.First(d => d.Day == DayOfWeek.Monday);
        var range = weekday.WorkingTimes[0];
        calendar.WeekDays.Remove(weekday);
        Assert.Throws<InvalidOperationException>(() => range.From = TimeSpan.Zero);
        document.Calendars.Remove(calendar);
        Assert.Throws<InvalidOperationException>(() => calendar.SetWorkingDay(DayOfWeek.Monday));
    }

    [Fact]
    public void CrossDocumentReferencesAndCalendarCyclesAreRejectedWithoutPartialChanges() {
        using var document = ProjectDocument.Create();
        using var other = ProjectDocument.Create();
        var task = document.Tasks.Add("Task");
        var foreign = other.Tasks.Add("Foreign");
        Assert.Throws<ArgumentException>(() => document.Dependencies.Add(task, foreign));
        Assert.Throws<ArgumentException>(() => document.Assignments.Add(task, other.Resources.AddWork("Foreign")));
        Assert.Throws<ArgumentException>(() => task.MoveTo(foreign));
        var parent = document.Calendars.Add("Base");
        var child = document.Calendars.Add("Child", parent);
        Assert.Throws<ArgumentException>(() => parent.BaseCalendar = child);
        Assert.Null(parent.BaseCalendar);
        Assert.Empty(document.Dependencies);
        Assert.Empty(document.Assignments);
    }

    [Fact]
    public void ForwardReferencesResolveTogetherAndMissingAliasesDoNotAddPartialLinks() {
        using var document = ProjectDocument.Create();
        var builder = document.AsFluent()
            .Task("first", "First", t => t.After("second"))
            .Task("second", "Second", t => t.After("missing"));
        Assert.Throws<InvalidOperationException>(() => builder.End());
        Assert.Empty(document.Dependencies);
        builder.Task("missing", "Third").End();
        Assert.Equal(2, document.Dependencies.Count);
        Assert.Same(document.Tasks[1], document.Dependencies[0].Predecessor);
    }

    [Fact]
    public void BatchRevisionsAndReportsDescribeStoredValuesWithoutCalculation() {
        using var document = ProjectDocument.Create();
        long initial = document.Revision;
        using (document.BeginUpdate()) {
            var task = document.Tasks.Add("Task");
            task.Duration = ProjectDuration.WorkingDays(5);
            using (document.BeginUpdate()) task.Notes = "Batch";
            Assert.Equal(initial, document.Revision);
            Assert.Throws<InvalidOperationException>(() => document.ToXml());
        }
        Assert.Equal(initial + 1, document.Revision);
        Assert.True(document.IsScheduleStale);
        Assert.Null(document.Tasks[0].Start);
        var report = document.Validate();
        document.Tasks[0].PercentComplete = 101;
        Assert.Equal(initial + 1, report.ModelRevision);
        Assert.False(report.HasErrors);
        Assert.True(document.Validate().HasErrors);
        Assert.Throws<InvalidDataException>(() => document.ToXml());
    }
}

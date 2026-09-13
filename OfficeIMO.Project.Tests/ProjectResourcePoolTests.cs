namespace OfficeIMO.Project.Tests;

public sealed class ProjectResourcePoolTests {
    [Fact]
    public void PoolCombinesProjectQualifiedAssignmentsUsingExplicitPoolAvailability() {
        using var pool = CreateProject();
        var resource = pool.Resources.AddWork("Shared engineer"); resource.MaxUnits = ProjectUnits.Fraction(1);
        var reduced = resource.AvailabilityPeriods.Add(); reduced.From = pool.Settings.StartDate;
        reduced.Through = reduced.From!.Value.AddHours(2).AddMinutes(-1); reduced.Units = ProjectUnits.Fraction(1);
        var increased = resource.AvailabilityPeriods.Add(); increased.From = reduced.From.Value.AddHours(2);
        increased.Through = reduced.From.Value.AddDays(1); increased.Units = ProjectUnits.Fraction(2);
        using var first = CreateProject(); using var second = CreateProject();
        var one = AddDemand(first, resource.Uid); var two = AddDemand(second, resource.Uid);
        long revision = pool.Revision;
        var result = pool.AnalyzeResourcePool(new[] { one, two });
        Assert.Equal(revision, pool.Revision);
        Assert.Equal(2, result.Intervals.Count);
        var overloaded = Assert.Single(result.Overallocations);
        Assert.Equal(pool.Settings.StartDate, overloaded.Start);
        Assert.Equal(overloaded.Start.AddHours(2), overloaded.Finish);
        Assert.Equal(2m, overloaded.Units); Assert.Equal(1m, overloaded.ExcessUnits);
        Assert.Equal(2, overloaded.Assignments.Count);
        Assert.Same(first, overloaded.Assignments[0].Binding.Project);
        Assert.Same(second, overloaded.Assignments[1].Binding.Project);
        Assert.Equal(overloaded.Assignments[0].AssignmentUid, overloaded.Assignments[1].AssignmentUid);
        Assert.Empty(pool.AnalyzeResourcePool(new[] { one }).Overallocations);
        Assert.Throws<ArgumentException>(() => pool.AnalyzeResourcePool(new[] { one, one }));
        Assert.Throws<InvalidOperationException>(() => pool.AnalyzeResourcePool(new[] { one, two }, maxBindings: 1));
        Assert.Throws<InvalidOperationException>(() => pool.AnalyzeResourcePool(new[] { one, two }, maxIntervals: 1));
        using var cancel = new CancellationTokenSource(); cancel.Cancel();
        Assert.Throws<OperationCanceledException>(() => pool.AnalyzeResourcePool(new[] { one }, cancellationToken: cancel.Token));
        first.Tasks.First().Name = "Changed";
        Assert.Throws<InvalidOperationException>(() => pool.AnalyzeResourcePool(new[] { one, two }));
    }

    [Fact]
    public void ProducerSharedResourceSnapshotsCanBeExplicitlyBoundWithoutNativePoolLinks() {
        using var owner = ProjectDocument.Load(ProjectResourceCapacityTests.Fixture("external-owner"));
        using var consumer = ProjectDocument.Load(ProjectResourceCapacityTests.Fixture("external-consumer"));
        var settings = new ProjectScheduleOptions { CalculateAssignments = true, ExternalProjectResolver = (_, _, _) => owner };
        var schedule = consumer.CalculateSchedule(settings); schedule.Report.ThrowIfErrors();
        var local = consumer.Resources.Single(r => r.Uid > 0);
        var pooled = owner.Resources.Single(r => r.Uid > 0);
        var binding = new ProjectResourcePoolBinding(consumer, schedule, local.Uid, pooled.Uid);
        var result = owner.AnalyzeResourcePool(new[] { binding });
        Assert.NotEmpty(result.Intervals); Assert.Empty(result.Overallocations);
        owner.Tasks.First(t => t.Uid > 0).Name = "Updated source";
        Assert.Throws<InvalidOperationException>(() => owner.AnalyzeResourcePool(new[] { binding }));
        Assert.Throws<InvalidOperationException>(() => consumer.AnalyzeResourceAllocation(schedule));
    }

    private static ProjectDocument CreateProject() {
        var document = ProjectDocument.Create(); document.Calendar = document.Calendars.AddStandardWorkingWeek();
        document.Settings.StartDate = new DateTime(2026, 10, 5, 8, 0, 0); return document;
    }
    private static ProjectResourcePoolBinding AddDemand(ProjectDocument project, int poolUid) {
        var task = project.Tasks.Add("Work"); task.Duration = ProjectDuration.WorkingMinutes(240);
        var resource = project.Resources.AddWork("Different local label"); project.Assignments.Add(task, resource);
        var schedule = project.CalculateSchedule(new ProjectScheduleOptions { CalculateAssignments = true }); schedule.Report.ThrowIfErrors();
        return new ProjectResourcePoolBinding(project, schedule, resource.Uid, poolUid);
    }
}

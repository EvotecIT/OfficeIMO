namespace OfficeIMO.Project.Tests;

public sealed class ProjectExternalSchedulingTests {
    [Fact]
    public void ProducerCrossProjectLinksResolveDisplayIdsAndKeepBorrowedSourcesUnchanged() {
        using var owner = ProjectDocument.Load(ProjectResourceCapacityTests.Fixture("external-owner"));
        using var consumer = ProjectDocument.Load(ProjectResourceCapacityTests.Fixture("external-consumer"));
        Assert.Equal(3, owner.AllTasks.Single(t => t.DisplayId == 1).Uid);
        Assert.True(consumer.CalculateSchedule(new ProjectScheduleOptions { CalculateAssignments = true }).Report.HasErrors);
        int calls = 0; long revision = owner.Revision;
        var options = new ProjectScheduleOptions { CalculateAssignments = true, ExternalProjectResolver = (origin, reference, token) => {
            Assert.Same(consumer, origin); Assert.Equal("owner.mpp", reference); calls++; return owner;
        } };
        var result = consumer.CalculateSchedule(options); result.Report.ThrowIfErrors();
        Assert.Equal(1, calls); Assert.Single(result.ExternalSources); Assert.Equal(revision, owner.Revision);
        var local = result.Tasks.Single(t => t.TaskUid == 1);
        Assert.Equal(consumer.Tasks.GetByUid(1).Start, local.Start); Assert.Equal(consumer.Tasks.GetByUid(1).Finish, local.Finish);
        consumer.ApplySchedule(result); Assert.Equal(revision, owner.Revision);
        var next = consumer.CalculateSchedule(options); next.Report.ThrowIfErrors();
        owner.Tasks.GetByUid(3).ConstraintType = ProjectConstraintType.StartNoEarlierThan;
        owner.Tasks.GetByUid(3).ConstraintDate = new DateTime(2026, 10, 6, 8, 0, 0);
        Assert.Throws<InvalidOperationException>(() => consumer.ApplySchedule(next));
        var recalculated = consumer.CalculateSchedule(options); recalculated.Report.ThrowIfErrors();
        Assert.Equal(new DateTime(2026, 10, 8, 8, 0, 0), recalculated.Tasks.Single(t => t.TaskUid == 1).Start);
    }
    [Fact]
    public void ResolverCyclesMissingTargetsAndEmptyUpdateScopesCannotBeApplied() {
        using var owner = ProjectDocument.Load(ProjectResourceCapacityTests.Fixture("external-owner"));
        using var consumer = ProjectDocument.Load(ProjectResourceCapacityTests.Fixture("external-consumer"));
        var options = new ProjectScheduleOptions { CalculateAssignments = true, ExternalProjectResolver = (origin, reference, token) => consumer };
        Assert.True(consumer.CalculateSchedule(options).Report.HasErrors);
        options.ExternalProjectResolver = (origin, reference, token) => null;
        Assert.True(consumer.CalculateSchedule(options).Report.HasErrors);
        options.ExternalProjectResolver = (origin, reference, token) => owner;
        using (owner.BeginUpdate()) Assert.True(consumer.CalculateSchedule(options).Report.HasErrors);
        var result = consumer.CalculateSchedule(options); result.Report.ThrowIfErrors();
        using (owner.BeginUpdate()) Assert.Throws<InvalidOperationException>(() => consumer.ApplySchedule(result));
        Assert.Throws<OperationCanceledException>(() => consumer.CalculateSchedule(options, new CancellationToken(true)));
    }

    [Fact]
    public void ExternalTaskBudgetRejectsBeforeCalculatingTheSource() {
        using var owner = ProjectDocument.Load(ProjectResourceCapacityTests.Fixture("external-owner"));
        using var consumer = ProjectDocument.Load(ProjectResourceCapacityTests.Fixture("external-consumer"));
        var options = new ProjectScheduleOptions {
            CalculateAssignments = true, MaxExternalTasks = 1,
            ExternalProjectResolver = (origin, reference, token) => owner
        };
        var start = consumer.Tasks.GetByUid(1).Start;
        var result = consumer.CalculateSchedule(options);
        Assert.True(result.Report.HasErrors);
        Assert.Contains(result.Report.Diagnostics, d => d.Message.Contains("MaxExternalTasks"));
        Assert.Throws<InvalidDataException>(() => consumer.ApplySchedule(result));
        Assert.Equal(start, consumer.Tasks.GetByUid(1).Start);
        options.MaxExternalTasks = owner.AllTasks.Count();
        consumer.CalculateSchedule(options).Report.ThrowIfErrors();
    }

    [Fact]
    public void IntervalBudgetIncludesLocalAndExternalSchedulesTogether() {
        using var owner = ProjectDocument.Load(ProjectResourceCapacityTests.Fixture("external-owner"));
        using var consumer = ProjectDocument.Load(ProjectResourceCapacityTests.Fixture("external-consumer"));
        var options = new ProjectScheduleOptions {
            CalculateAssignments = true, ExternalProjectResolver = (origin, reference, token) => owner
        };
        var source = owner.CalculateSchedule(options); source.Report.ThrowIfErrors();
        var local = consumer.CalculateSchedule(options); local.Report.ThrowIfErrors();
        int sourceCount = source.Assignments.Sum(a => a.Intervals.Count + a.Costs.Count);
        int localCount = local.Assignments.Sum(a => a.Intervals.Count + a.Costs.Count);
        Assert.True(sourceCount > 0); Assert.True(localCount > 0);
        options.MaxIntervals = Math.Max(sourceCount, localCount);
        var limited = consumer.CalculateSchedule(options);
        Assert.Contains(limited.Report.Diagnostics, d => d.Code == "PROJECT_CALCULATION_INTERVAL_LIMIT");
        Assert.Throws<InvalidDataException>(() => consumer.ApplySchedule(limited));
        options.MaxIntervals = sourceCount + localCount;
        consumer.CalculateSchedule(options).Report.ThrowIfErrors();
    }
}

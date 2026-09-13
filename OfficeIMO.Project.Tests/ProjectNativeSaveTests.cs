using System.Collections.Generic;
using OfficeIMO.Core.Internal;

namespace OfficeIMO.Project.Tests;

public sealed class ProjectNativeSaveTests {
    private static ProjectSaveOptions Allow(ProjectFileFormat format = ProjectFileFormat.Mpp14) => new ProjectSaveOptions { Format = format, LossPolicy = OfficeConversionLossPolicy.Allow };

    [Fact]
    public void ScalarGrowthPreservesUnrelatedCompoundStreamsAndRichTaskContent() {
        byte[] source = File.ReadAllBytes(ProjectNativeTests.Fixture("delivery.mpp"));
        using var document = ProjectDocument.Load(new MemoryStream(source));
        var task = document.Tasks.GetByUid(3); string? notes = task.Notes;
        task.Name = "Build revised café / Łódź / 日本語 with a longer label"; document.Title = "Edited native title";
        Assert.False(document.AssessSave().HasLoss);
        using var output = new MemoryStream(); document.Save(output);
        using var reload = ProjectDocument.Load(new MemoryStream(output.ToArray()));
        Assert.Equal(task.Name, reload.Tasks.GetByUid(3).Name); Assert.Equal(notes, reload.Tasks.GetByUid(3).Notes);
        Assert.Equal("Edited native title", reload.Title); Assert.Equal(5000m, reload.Tasks.GetByUid(3).Baselines.Single(b => b.Number == 0).Cost);
        var before = Compound(source); var after = Compound(output.ToArray());
        Assert.Equal(before.Streams.Keys.OrderBy(k => k), after.Streams.Keys.OrderBy(k => k));
        foreach (string name in before.Streams.Keys.Where(k => !k.StartsWith("   114/TBkndTask/", StringComparison.Ordinal)
            && k != "   114/Props" && k != OfficeOlePropertySetWriter.SummaryInformationStreamName))
            Assert.Equal(before.Streams[name], after.Streams[name]);
    }

    [Fact]
    public void ReparentRequiresExplicitLossPolicyAndWritesProducerSortPositions() {
        using var document = ProjectDocument.Load(ProjectNativeTests.Fixture("delivery.mpp"));
        document.Tasks.GetByUid(3).MoveTo(null);
        Assert.Contains(document.AssessSave().Diagnostics, d => d.Code == "PROJECT_NATIVE_STRUCTURAL_OPAQUE" && d.RepresentsLoss);
        using var destination = new MemoryStream();
        Assert.Throws<InvalidOperationException>(() => document.Save(destination)); Assert.Empty(destination.ToArray());
        document.Save(destination, Allow());
        using var reload = ProjectDocument.Load(new MemoryStream(destination.ToArray()));
        Assert.Equal(new[] { 0, 1, 2, 4, 3 }, reload.AllTasks.Select(t => t.Uid));
        Assert.Null(reload.Tasks.GetByUid(3).Parent); Assert.Equal(1, reload.Tasks.GetByUid(4).Parent!.Uid);
        var table = Table(destination.ToArray(), "Task", 0x14, 0x0b400056);
        var positions = table.Records.ToDictionary(r => r.Integer(0x0b400056)!.Value, r => r.Number(0x0b400479));
        // Project uses this independent position field on open; changing display IDs alone is insufficient.
        Assert.Equal(5m, positions[3]); Assert.Equal(4m, positions[4]);
        Assert.Contains(reload.Dependencies, d => d.Predecessor!.Uid == 3 && d.Successor.Uid == 4);
    }

    [Fact]
    public void CascadeDeletionRemovesNativeRowsAndTheirRelationships() {
        using var document = ProjectDocument.Load(ProjectNativeTests.Fixture("delivery.mpp"));
        var task = document.Tasks.GetByUid(3); task.Parent!.Children.Remove(task, ProjectRemovalMode.Cascade);
        var resource = document.Resources.GetByUid(1); var calendar = resource.Calendar;
        document.Resources.Remove(resource, ProjectRemovalMode.Cascade); document.Calendars.Remove(calendar!);
        using var output = new MemoryStream(); document.Save(output, Allow());
        using var reload = ProjectDocument.Load(new MemoryStream(output.ToArray()));
        Assert.DoesNotContain(reload.AllTasks, t => t.Uid == 3); Assert.DoesNotContain(reload.Resources, r => r.Uid == 1);
        Assert.DoesNotContain(reload.Assignments, a => a.Task?.Uid == 3 || a.Resource?.Uid == 1);
        Assert.Equal(new[] { 2, 5 }, reload.Assignments.Select(a => a.Uid)); // Untouched producer placeholders use resource UID 0.
        Assert.Empty(reload.Dependencies); Assert.DoesNotContain(reload.Calendars, c => c.Uid == calendar!.Uid);
        Assert.Equal(new[] { 0, 1, 2, 4 }, reload.AllTasks.Select(t => t.Uid));
    }

    [Fact]
    public void CloneUsesEditedModelAndXmlSaveDoesNotReviveStaleNativeBytes() {
        using var document = ProjectDocument.Load(ProjectNativeTests.Fixture("delivery.mpp"));
        document.Tasks.GetByUid(3).Name = "Current edited task";
        using var clone = document.Clone(); Assert.Equal("Current edited task", clone.Tasks.GetByUid(3).Name);
        Assert.True(document.IsModified);
        using var xml = new MemoryStream(); document.Save(xml, Allow(ProjectFileFormat.Xml)); Assert.False(document.IsModified);
        using var native = new MemoryStream(); document.Save(native, Allow());
        using var reloaded = ProjectDocument.Load(new MemoryStream(native.ToArray()));
        Assert.Equal("Current edited task", reloaded.Tasks.GetByUid(3).Name);
        document.Tasks.GetByUid(3).Name = "Second edit";
        using var second = new MemoryStream(); document.Save(second, Allow());
        using var again = ProjectDocument.Load(new MemoryStream(second.ToArray())); Assert.Equal("Second edit", again.Tasks.GetByUid(3).Name);
    }

    [Fact]
    public void AssessmentDoesNotAuthorizeLaterUnsupportedMutations() {
        using var document = ProjectDocument.Load(ProjectNativeTests.Fixture("delivery.mpp"));
        document.Tasks.GetByUid(3).Name = "Assessed"; var report = document.AssessSave(); Assert.False(report.HasErrors);
        document.Tasks.GetByUid(3).Notes = "Not qualified as a native edit";
        Assert.NotEqual(report.ModelRevision, document.Revision);
        using var target = new MemoryStream(new byte[] { 1, 2, 3 }, true);
        Assert.Throws<InvalidDataException>(() => document.Save(target, Allow())); Assert.Equal(new byte[] { 1, 2, 3 }, target.ToArray());
        Assert.True(document.IsModified);
    }

    [Fact]
    public void NativePrecisionAndIdentityFailuresAreDiagnosedBeforeCommit() {
        using var document = ProjectDocument.Load(ProjectNativeTests.Fixture("delivery.mpp"));
        var task = document.Tasks.GetByUid(3); task.Start = task.Start!.Value.AddSeconds(1);
        Assert.Contains(document.AssessSave(Allow()).Diagnostics, d => d.Code == "PROJECT_NATIVE_VALUE_UNREPRESENTABLE");
        task.Start = task.Start.Value.AddSeconds(-1); task.Guid = document.Tasks.GetByUid(2).Guid;
        Assert.Contains(document.AssessSave(Allow()).Diagnostics, d => d.Code == "PROJECT_NATIVE_GUID_INVALID");
        using var target = new MemoryStream(); Assert.Throws<InvalidDataException>(() => document.Save(target, Allow())); Assert.Equal(0, target.Length);
    }

    [Fact]
    public void NativeAssessmentRejectsConflictingDisplayIdsAndMissingProjectCalendar() {
        using var document = ProjectDocument.Load(ProjectNativeTests.Fixture("delivery.mpp"));
        Assert.True(document.AssessSave(Allow(ProjectFileFormat.Xml)).HasLoss);
        var task = document.Tasks.GetByUid(3); task.DisplayId = 99;
        Assert.Contains(document.AssessSave(Allow()).Diagnostics, d => d.Code == "PROJECT_NATIVE_DISPLAY_ORDER");
        task.DisplayId = 3; document.Calendar = null;
        Assert.Contains(document.AssessSave(Allow()).Diagnostics, d => d.Code == "PROJECT_NATIVE_CALENDAR_REQUIRED");
    }

    [Fact]
    public void ChangedTaskAndResourceGuidsPropagateToNativeRelationshipRecords() {
        using var document = ProjectDocument.Load(ProjectNativeTests.Fixture("delivery.mpp"));
        var taskGuid = Guid.NewGuid(); var resourceGuid = Guid.NewGuid(); var parentGuid = Guid.NewGuid();
        document.Tasks.GetByUid(3).Guid = taskGuid; document.Resources.GetByUid(1).Guid = resourceGuid; document.Tasks.GetByUid(1).Guid = parentGuid;
        using var output = new MemoryStream(); document.Save(output, Allow());
        var assignment = Table(output.ToArray(), "Assn", 0x17, 0x0f400000).Records.Single(r => r.Integer(0x0f400000) == 4);
        Assert.Equal(taskGuid, new Guid(assignment.Value(0x0f40027d)!.Value.Copy()));
        Assert.Equal(resourceGuid, new Guid(assignment.Value(0x0f40027e)!.Value.Copy()));
        var task = Table(output.ToArray(), "Task", 0x14, 0x0b400056).Records.Single(r => r.Integer(0x0b400056) == 3);
        Assert.Equal(parentGuid, new Guid(task.Value(0x0b40047f)!.Value.Copy()));
    }

    [Fact]
    public void RejectedNativeSaveLeavesExistingFileAndAssociationUntouched() {
        string path = Path.Combine(Path.GetTempPath(), "OfficeIMO-native-" + Guid.NewGuid().ToString("N") + ".mpp");
        try {
            File.WriteAllBytes(path, new byte[] { 9, 8, 7 });
            using var document = ProjectDocument.Load(ProjectNativeTests.Fixture("delivery.mpp"));
            document.Title = "Pending metadata";
            Assert.Throws<InvalidDataException>(() => document.Save(path, new ProjectSaveOptions { MaxOutputBytes = 128 }));
            Assert.Equal(new byte[] { 9, 8, 7 }, File.ReadAllBytes(path)); Assert.True(document.IsModified);
            Assert.Throws<OperationCanceledException>(() => document.Save(path, cancellationToken: new CancellationToken(true)));
            Assert.Equal(new byte[] { 9, 8, 7 }, File.ReadAllBytes(path));
        } finally { File.Delete(path); }
    }

    [Fact]
    public void ChangingProjectGuidKeepsCommittedImplicitEntityReferencesStable() {
        using var document = ProjectNativeAuthoringTests.Create(); document.Guid = Guid.NewGuid();
        using var first = new MemoryStream(); document.Save(first, Allow());
        using var original = ProjectDocument.Load(new MemoryStream(first.ToArray()));
        var originalFile = Compound(first.ToArray());
        var originalProperties = ProjectNativeProperties.Read(originalFile.Streams["   114/Props"], default);
        var originalTask = Table(first.ToArray(), "Task", 0x14, 0x0b400056).Records.Single(r => r.Uid == 1);
        var originalResource = Table(first.ToArray(), "Rsc", 0x15, 0x0c40001b).Records.Single(r => r.Uid == 1);
        var originalCalendarGuid = new Guid(originalProperties[0x024013c2].Copy());
        var originalTaskGuid = new Guid(originalTask.Value(0x0b400477)!.Value.Copy());
        var originalResourceGuid = new Guid(originalResource.Value(0x0c4002d8)!.Value.Copy());
        Assert.Null(original.Calendar!.Guid); Assert.Null(original.Tasks.GetByUid(1).Guid); Assert.Null(original.Resources.GetByUid(1).Guid);
        document.Guid = Guid.NewGuid();
        var child = document.Tasks.GetByUid(1).Children.Add("Later child"); child.Duration = ProjectDuration.WorkingDays(1);
        var assignment = document.Assignments.Add(child, document.Resources.GetByUid(1)); assignment.Units = ProjectUnits.Percent(100);
        using var second = new MemoryStream(); document.Save(second, Allow());
        using var loaded = ProjectDocument.Load(new MemoryStream(second.ToArray()));
        Assert.Equal(document.Guid, loaded.Guid);
        Assert.Null(loaded.Calendar!.Guid);
        var file = Compound(second.ToArray()); var properties = ProjectNativeProperties.Read(file.Streams["   114/Props"], default);
        Assert.Equal(originalCalendarGuid, new Guid(properties[0x024013c2].Copy()));
        var taskRow = Table(second.ToArray(), "Task", 0x14, 0x0b400056).Records.Single(r => r.Uid == child.Uid);
        Assert.Equal(originalTaskGuid, new Guid(taskRow.Value(0x0b40047f)!.Value.Copy()));
        var assignmentRow = Table(second.ToArray(), "Assn", 0x17, 0x0f400000).Records.Single(r => r.Uid == assignment.Uid);
        Assert.Equal(originalResourceGuid, new Guid(assignmentRow.Value(0x0f40027e)!.Value.Copy()));
        Assert.Equal(new Guid(taskRow.Value(0x0b400477)!.Value.Copy()), new Guid(assignmentRow.Value(0x0f40027d)!.Value.Copy()));
        Assert.Null(loaded.Tasks.GetByUid(child.Uid).Guid);
    }

    private static OfficeCompoundFile Compound(byte[] bytes) {
        Assert.True(OfficeCompoundFileReader.TryRead(bytes, out OfficeCompoundFile? file, out var error), error); return file!;
    }
    private static ProjectNativeTable Table(byte[] bytes, string name, uint id, uint uid) {
        var file = Compound(bytes); var props = ProjectNativeProperties.Read(file.Streams["   114/Props"], default);
        var table = new ProjectNativeTable(file, "TBknd" + name, props[0x03000000 | id], props[0x00020000 | id], 300000, default);
        foreach (var row in table.Records) row.Uid = row.Integer(uid)!.Value;
        return table;
    }
}

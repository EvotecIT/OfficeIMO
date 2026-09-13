using OfficeIMO.Core.Internal;
using System.Collections.Generic;

namespace OfficeIMO.Project.Tests;

public sealed class ProjectNative98Tests {
    private static ProjectSaveOptions Options(bool allow = true) => new ProjectSaveOptions {
        Format = ProjectFileFormat.Mpp8, LossPolicy = allow ? OfficeConversionLossPolicy.Allow : OfficeConversionLossPolicy.Block
    };

    [Theory]
    [InlineData(0, 100)]
    [InlineData(100, 100)]
    [InlineData(500, 500)]
    [InlineData(549, 500)]
    [InlineData(550, 600)]
    [InlineData(1000, 1000)]
    public void PriorityClassesReportAnyQuantization(int requested, int represented) {
        using var document = ProjectNativeAuthoringTests.Create();
        document.Tasks.GetByUid(2).Priority = requested;
        var report = document.AssessSave(Options());
        Assert.Equal(requested != represented, report.Diagnostics.Any(d => d.Code == "PROJECT_NATIVE_PRIORITY_QUANTIZED"));
        using var output = new MemoryStream(); document.Save(output, Options());
        using var reopened = ProjectDocument.Load(new MemoryStream(output.ToArray()));
        Assert.Equal(represented, reopened.Tasks.GetByUid(2).Priority);
    }

    [Fact]
    public void DeferredTextGrowthAndDeletionRetainResourceCalendarAndAssignmentWork() {
        using var document = ProjectNativeAuthoringTests.Create();
        using var original = new MemoryStream(); document.Save(original, Options());
        var task = document.Tasks.GetByUid(2); task.Name = string.Concat(Enumerable.Repeat("Łódź 日本語 / ", 15));
        var removed = document.Tasks.Add("Removed after save");
        using var intermediate = new MemoryStream(); document.Save(intermediate, Options());
        document.Tasks.Remove(removed, ProjectRemovalMode.Cascade);
        using var output = new MemoryStream(); document.Save(output, Options());
        using var reopened = ProjectDocument.Load(new MemoryStream(output.ToArray()));
        Assert.Equal(task.Name, reopened.Tasks.GetByUid(2).Name);
        Assert.DoesNotContain(reopened.AllTasks, t => t.Uid == removed.Uid);
        Assert.Equal(480m, reopened.Assignments.Single().Work!.Value.Minutes);
        Assert.Equal(document.Resources.First(r => r.Uid > 0).Calendar!.Uid, reopened.Resources.First(r => r.Uid > 0).Calendar!.Uid);
        Assert.Equal(document.Calendar!.GetWorkingIntervals(new DateTime(2026, 10, 12)), reopened.Calendar!.GetWorkingIntervals(new DateTime(2026, 10, 12)));
    }

    [Fact]
    public void SixByteWorkRejectsFractionalStorageUnitsBeforeWriting() {
        using var document = ProjectNativeAuthoringTests.Create();
        document.Assignments.Single().Work = new ProjectWork(0.0001m);
        using var output = new MemoryStream();
        Assert.Contains(document.AssessSave(Options()).Diagnostics, d => d.Code == "PROJECT_NATIVE_VALUE_UNREPRESENTABLE" && d.Severity == ProjectDiagnosticSeverity.Error);
        Assert.Throws<InvalidDataException>(() => document.Save(output, Options())); Assert.Empty(output.ToArray());
    }

    [Fact]
    public void CalendarRejectsMoreThanThreeIntervalsBeforeWriting() {
        using var document = ProjectNativeAuthoringTests.Create();
        document.Calendar!.SetWorkingDay(DayOfWeek.Monday, ProjectWorkingTime.Hours(1, 2), ProjectWorkingTime.Hours(3, 4),
            ProjectWorkingTime.Hours(5, 6), ProjectWorkingTime.Hours(7, 8));
        using var output = new MemoryStream();
        Assert.Contains(document.AssessSave(Options()).Diagnostics, d => d.Code == "PROJECT_NATIVE_CALENDAR_VALUE");
        Assert.Throws<InvalidDataException>(() => document.Save(output, Options())); Assert.Empty(output.ToArray());
    }

    [Fact]
    public void CyclicDeferredChainIsRejected() {
        using var document = ProjectNativeAuthoringTests.Create();
        using var output = new MemoryStream(); document.Save(output, Options());
        Assert.True(OfficeCompoundFileReader.TryRead(output.ToArray(), out OfficeCompoundFile? file, out var error), error);
        const string path = "   1/TBkndTask/FixDeferFix   0";
        var stream = (byte[])file!.Streams[path].Clone();
        // Every deferred value begins at a 36-byte allocation boundary after the free-list head.
        Buffer.BlockCopy(BitConverter.GetBytes(4), 0, stream, 4, 4);
        Buffer.BlockCopy(BitConverter.GetBytes(64), 0, stream, 8, 4);
        byte[] corrupt = OfficeCompoundFileWriter.Rewrite(file, new Dictionary<string, byte[]> { [path] = stream });
        Assert.Throws<InvalidDataException>(() => ProjectDocument.Load(new MemoryStream(corrupt)));
    }

    [Fact]
    public void LossAcceptanceDoesNotMakeUnrepresentedModelValuesDisappearFromLaterAssessments() {
        using var document = ProjectNativeAuthoringTests.Create();
        document.Tasks.GetByUid(2).Priority = 549;
        using var output = new MemoryStream(); document.Save(output, Options());
        var report = document.AssessSave(Options(false));
        Assert.Contains(report.Diagnostics, d => d.Code == "PROJECT_NATIVE_PRIORITY_QUANTIZED");
        Assert.Contains(report.Diagnostics, d => d.Code == "PROJECT_NATIVE_ALIAS_LOSS");
        using var second = new MemoryStream();
        Assert.Throws<InvalidOperationException>(() => document.Save(second, Options(false))); Assert.Empty(second.ToArray());
    }

    [Fact]
    public void ConflictingLegacyFinishCacheIsRejectedBeforeOutput() {
        using var document = ProjectNativeAuthoringTests.Create();
        var task = document.Tasks.GetByUid(2); task.EarlyFinish = task.Finish!.Value.AddDays(1);
        using var output = new MemoryStream();
        Assert.Contains(document.AssessSave(Options()).Diagnostics, d => d.Code == "PROJECT_NATIVE_LEGACY_FINISH_CACHE");
        Assert.Throws<InvalidDataException>(() => document.Save(output, Options())); Assert.Empty(output.ToArray());
    }

    [Fact]
    public void ExternalPropertyRewriteRetainsPayloadAndTrailer() {
        var properties = new ProjectNativePropertySet(true);
        properties.Set(1, 0x10000, new byte[] { 4, 0, 0, 0, 10, 20, 30, 40 }); properties.Set(2, 4, BitConverter.GetBytes(17));
        byte[] source = properties.Serialize(1000, default).Concat(new byte[] { 91, 92 }).ToArray();
        var editing = new ProjectNativePropertySet(source, default, true); editing.Set(2, 4, BitConverter.GetBytes(19));
        byte[] saved = editing.Serialize(1000, default);
        var parsed = ProjectNativeProperties.Read(saved, default, true);
        Assert.Equal(new byte[] { 4, 0, 0, 0, 10, 20, 30, 40 }, parsed[1].Copy()); Assert.Equal(19, parsed[2].Int32());
        Assert.Equal(new byte[] { 91, 92 }, saved.Skip(saved.Length - 2));
    }
}

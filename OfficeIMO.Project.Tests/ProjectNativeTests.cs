using System.Collections.Generic;
using OfficeIMO.Core.Internal;

namespace OfficeIMO.Project.Tests;

public sealed class ProjectNativeTests {
    internal static string Fixture(string name, string folder = "Project2024") => Path.Combine(AppContext.BaseDirectory, "Fixtures", folder, name);
    [Theory]
    [InlineData("empty")]
    [InlineData("delivery")]
    [InlineData("relationships")]
    [InlineData("calendars")]
    [InlineData("resources")]
    [InlineData("actuals")]
    [InlineData("custom-fields")]
    [InlineData("constraints", "Project2024Semantics")]
    [InlineData("backward", "Project2024Semantics")]
    [InlineData("work-cost", "Project2024Semantics")]
    [InlineData("rich-fields", "Project2024Semantics")]
    [InlineData("working-weeks", "Project2024WorkWeeks")]
    public void NativeReadAndCloneRetainWholeSourceAndTaskIdentities(string name, string folder = "Project2024") {
        var bytes = File.ReadAllBytes(Fixture(name + ".mpp", folder));
        using var native = ProjectDocument.Load(new MemoryStream(bytes));
        using var xml = ProjectDocument.Load(Fixture(name + ".xml", folder));
        Assert.Equal("MPP14", native.NativeInfo!.Generation);
        Assert.Equal(xml.Title, native.Title); Assert.Equal("OfficeIMO", native.Author);
        Assert.Equal("OfficeIMO", native.Company); Assert.Equal("OfficeIMO", native.Manager);
        Assert.False(native.IsModified);
        Assert.Equal(xml.AllTasks.Select(t => t.Uid), native.AllTasks.Select(t => t.Uid));
        foreach (var expected in xml.AllTasks) {
            var actual = native.Tasks.GetByUid(expected.Uid);
            Assert.Equal(expected.Name, actual.Name); Assert.Equal(expected.Parent?.Uid, actual.Parent?.Uid);
            Assert.Equal(expected.Start, actual.Start); Assert.Equal(expected.Finish, actual.Finish);
            Assert.Equal(expected.IsMilestone, actual.IsMilestone); Assert.Equal(expected.IsManual, actual.IsManual);
        }
        using var clone = native.Clone(); using var output = new MemoryStream(); clone.Save(output);
        Assert.Equal(bytes, output.ToArray());
        using var saved = new MemoryStream(); native.Save(saved); Assert.Equal(bytes, saved.ToArray());
    }
    [Theory]
    [InlineData(OfficeConversionLossPolicy.Block)]
    [InlineData(OfficeConversionLossPolicy.Allow)]
    public void NativeEditsNeverProduceAnUnqualifiedRewrite(OfficeConversionLossPolicy policy) {
        using var document = ProjectDocument.Load(Fixture("delivery.mpp"));
        document.Tasks.GetByUid(3).Name = "Changed";
        using var target = new MemoryStream(new byte[16], true);
        Assert.Throws<InvalidDataException>(() => document.Save(target, new ProjectSaveOptions { LossPolicy = policy }));
        Assert.Equal(new byte[16], target.ToArray());
        Assert.Throws<InvalidDataException>(() => document.Clone(new ProjectSaveOptions { LossPolicy = policy }));
        Assert.Throws<NotSupportedException>(() => document.ToXml());
    }
    [Fact]
    public void NativeInputAndOutputBudgetsAndCancellationAreEnforced() {
        var path = Fixture("delivery.mpp");
        Assert.Throws<InvalidDataException>(() => ProjectDocument.Load(path, new ProjectLoadOptions { MaxInputBytes = 128 }));
        Assert.Throws<InvalidDataException>(() => ProjectDocument.Load(path, new ProjectLoadOptions { MaxTasks = 2 }));
        Assert.Throws<InvalidDataException>(() => ProjectDocument.Load(path, new ProjectLoadOptions { MaxEntities = 4 }));
        Assert.Throws<OperationCanceledException>(() => ProjectDocument.Load(path, cancellationToken: new CancellationToken(true)));
        using var document = ProjectDocument.Load(path, new ProjectLoadOptions { MaxDiagnostics = 1 });
        Assert.True(document.ReadDiagnostics.Count <= 2);
        Assert.Contains(document.ReadDiagnostics, d => d.Code == "PROJECT_DIAGNOSTICS_TRUNCATED");
        using var destination = new MemoryStream();
        Assert.Throws<InvalidDataException>(() => document.Save(destination, new ProjectSaveOptions { MaxOutputBytes = 128 }));
        Assert.Throws<InvalidDataException>(() => document.Clone(new ProjectSaveOptions { MaxOutputBytes = 128 }));
        Assert.Equal(0, destination.Length);
    }
    [Fact]
    public void AllBaselineSlotsAndScalarCustomFieldsReadFromProducerRecords() {
        using var native = ProjectDocument.Load(Fixture("rich-fields.mpp", "Project2024Semantics"));
        using var xml = ProjectDocument.Load(Fixture("rich-fields.xml", "Project2024Semantics"));
        var task = native.Tasks.GetByUid(3); var expected = xml.Tasks.GetByUid(3);
        Assert.Equal(11, task.Baselines.Count);
        foreach (var baseline in expected.Baselines) {
            var actual = task.Baselines.Single(b => b.Number == baseline.Number);
            Assert.Equal(baseline.Start, actual.Start); Assert.Equal(baseline.Finish, actual.Finish);
            Assert.Equal(baseline.Work, actual.Work); Assert.Equal(baseline.Cost, actual.Cost);
            Assert.Equal(baseline.Duration?.Value, actual.Duration?.Value);
            Assert.Equal(baseline.Duration?.Unit, actual.Duration?.Unit);
            Assert.Equal(baseline.Duration?.IsElapsed, actual.Duration?.IsElapsed);
        }
        foreach (var field in task.CustomFields) {
            var reference = expected.CustomFields.Single(f => f.FieldId == field.FieldId);
            Assert.Equal(reference.Value, field.Value); Assert.Equal(reference.DurationFormat, field.DurationFormat);
        }
        Assert.Equal("Work package", native.CustomFields.Single(f => f.FieldId == "188743734").Alias);
        Assert.Equal("Native custom text", task.CustomFields.Single(f => f.FieldId == "188743734").Value);
        Assert.Equal("12345", task.CustomFields.Single(f => f.FieldId == "188743786").Value);
        Assert.Equal("1", task.CustomFields.Single(f => f.FieldId == "188743752").Value);
        Assert.Equal("1", task.CustomFields.Single(f => f.FieldId == "188743981").Value);
        Assert.Equal("1", native.Resources.GetByUid(1).CustomFields.Single(f => f.FieldId == "205521023").Value);
        Assert.Equal(11, native.Resources.GetByUid(1).Baselines.Count);
        Assert.Equal(11, native.Assignments.GetByUid(4).Baselines.Count);
    }
    [Fact]
    public void NativeCalendarExceptionsAndWorkWeeksRemainDistinctAndCalculate() {
        using var native = ProjectDocument.Load(Fixture("working-weeks.mpp", "Project2024WorkWeeks"));
        using var xml = ProjectDocument.Load(Fixture("working-weeks.xml", "Project2024WorkWeeks"));
        var calendar = native.Calendars.Single(c => c.Name == "Workshop");
        Assert.Equal(new[] { "Maintenance", "Shutdown" }, calendar.Exceptions.Select(e => e.Name));
        Assert.Single(calendar.WorkWeeks); Assert.Equal("Four-day week", calendar.WorkWeeks[0].Name);
        var schedule = native.CalculateSchedule(); schedule.Report.ThrowIfErrors();
        foreach (var task in schedule.Tasks) {
            var expected = xml.Tasks.GetByUid(task.TaskUid); Assert.Equal(expected.Start, task.Start); Assert.Equal(expected.Finish, task.Finish);
        }
    }
    [Theory]
    [InlineData("   114/TBkndTask/FixedMeta", 8, int.MaxValue)]
    [InlineData("   114/TBkndTask/Fixed2Meta", 8, 0)]
    [InlineData("   114/TBkndTask/VarMeta", 28, int.MaxValue)]
    [InlineData("   114/Props", 12, int.MaxValue)]
    public void MalformedProducerTablesFailBeforeReturningAPartialDocument(string stream, int offset, int value) {
        var bytes = Rewrite(stream, data => Buffer.BlockCopy(BitConverter.GetBytes(value), 0, data, offset, 4));
        Assert.Throws<InvalidDataException>(() => ProjectDocument.Load(new MemoryStream(bytes)));
    }
    [Fact]
    public void MetadataUsesUnsignedUtf8CodePagesAndUnpaddedDictionaryNames() {
        using var project = ProjectDocument.Load(Fixture("delivery.mpp"));
        Assert.Equal("OfficeIMO", project.Author);
        var bytes = File.ReadAllBytes(Fixture("delivery.mpp"));
        Assert.True(OfficeCompoundFileReader.TryRead(bytes, out OfficeCompoundFile? file, out var error), error);
        var sections = OfficeOlePropertySetReader.ReadSections(file!.Streams[OfficeOlePropertySetWriter.DocumentSummaryInformationStreamName]);
        var custom = sections.Single(s => s.Dictionary.Count != 0);
        Assert.Contains("Scheduled Start", custom.Dictionary.Values);
        Assert.Contains(custom.Properties.Values, p => p.AsString()?.Contains("zł") == true);
    }
    private static byte[] Rewrite(string stream, Action<byte[]> mutation) {
        Assert.True(OfficeCompoundFileReader.TryRead(File.ReadAllBytes(Fixture("delivery.mpp")), out OfficeCompoundFile? compound, out var error), error);
        var changed = (byte[])compound!.Streams[stream].Clone(); mutation(changed);
        return OfficeCompoundFileWriter.Rewrite(compound, new Dictionary<string, byte[]> { [stream] = changed });
    }
    [Theory]
    [InlineData("read-protected.mpp")]
    [InlineData("write-reserved.mpp")]
    public void IndependentlyProtectedInputsAreRejectedExplicitly(string name) {
        var error = Assert.Throws<NotSupportedException>(() => ProjectDocument.Load(Fixture(name, "Project2024Protection")));
        Assert.Contains("Protected MPP", error.Message);
    }
    [Fact]
    public async Task NativeAsyncPathAndStreamLifecycleRetainsBytesAndCallerOwnership() {
        string path = Fixture("delivery.mpp"); byte[] bytes = File.ReadAllBytes(path);
        using var file = await ProjectDocument.LoadAsync(path);
        using var input = new MemoryStream(bytes); input.Position = 17;
        using var streamed = await ProjectDocument.LoadAsync(input);
        Assert.Equal(17, input.Position); Assert.NotNull(file.NativeInfo); Assert.NotNull(streamed.NativeInfo);
        Assert.Equal(file.AllTasks.Select(t => t.Name), streamed.AllTasks.Select(t => t.Name));
        using var output = new MemoryStream(); await streamed.SaveAsync(output);
        Assert.Equal(bytes, output.ToArray());
        streamed.Dispose(); Assert.True(input.CanRead); Assert.True(output.CanWrite);
        await Assert.ThrowsAsync<OperationCanceledException>(() => ProjectDocument.LoadAsync(path, cancellationToken: new CancellationToken(true)));
    }
}

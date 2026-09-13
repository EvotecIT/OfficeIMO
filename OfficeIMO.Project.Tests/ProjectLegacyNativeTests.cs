using OfficeIMO.Core.Internal;
using System.Collections.Generic;

namespace OfficeIMO.Project.Tests;

public sealed class ProjectLegacyNativeTests {
    private static ProjectSaveOptions Native(ProjectFileFormat format, bool allow = false) => new ProjectSaveOptions {
        Format = format, LossPolicy = allow ? OfficeConversionLossPolicy.Allow : OfficeConversionLossPolicy.Block
    };

    [Fact]
    public void LegacyCurrencyCanRemainAbsentUntilXmlConversionIsRequested() {
        using var document = ProjectNativeAuthoringTests.Create();
        using var initial = new MemoryStream(); document.Save(initial, Native(ProjectFileFormat.Mpp9, true));
        document.Settings.CurrencyCode = null;
        Assert.DoesNotContain(document.Validate().Diagnostics, d => d.Code == "PROJECT_CURRENCY_CODE");
        using var updated = new MemoryStream(); document.Save(updated, Native(ProjectFileFormat.Mpp9, true));
        using var reopened = ProjectDocument.Load(new MemoryStream(updated.ToArray()));
        Assert.Null(reopened.Settings.CurrencyCode);
        Assert.Contains(reopened.AssessSave(Native(ProjectFileFormat.Xml)).Diagnostics,
            d => d.Code == "PROJECT_CURRENCY_CODE" && d.Severity == ProjectDiagnosticSeverity.Error);
        reopened.Settings.CurrencyCode = "PLN";
        Assert.DoesNotContain(reopened.AssessSave(Native(ProjectFileFormat.Xml)).Diagnostics, d => d.Code == "PROJECT_CURRENCY_CODE");
    }

    [Fact]
    public void LegacyStructuralEditsKeepOptionalSecondaryMetadataInStepWithPrimaryRecords() {
        using var document = ProjectNativeAuthoringTests.Create();
        using var initial = new MemoryStream(); document.Save(initial, Native(ProjectFileFormat.Mpp9, true));
        Assert.True(OfficeCompoundFileReader.TryRead(initial.ToArray(), out OfficeCompoundFile? file, out var error), error);
        // Later producers can include empty secondary record storage in MPP9.
        const string prefix = "   19/TBkndTask/";
        int rows = BitConverter.ToInt32(file!.Streams[prefix + "FixedMeta"], 8);
        var metadata = new byte[16 + rows * 9];
        Buffer.BlockCopy(BitConverter.GetBytes(0xfadfadbau), 0, metadata, 0, 4);
        metadata[4] = 4; Buffer.BlockCopy(BitConverter.GetBytes(rows), 0, metadata, 8, 4);
        byte[] source = OfficeCompoundFileWriter.Rewrite(file, new Dictionary<string, byte[]> {
            [prefix + "Fixed2Meta"] = metadata, [prefix + "Fixed2Data"] = Array.Empty<byte>()
        });
        using var editing = ProjectDocument.Load(new MemoryStream(source));
        var added = editing.Tasks.Add("Added task");
        var removed = editing.Tasks.GetByUid(2);
        (removed.Parent?.Children ?? editing.Tasks).Remove(removed, ProjectRemovalMode.Cascade);
        editing.Tasks.Add("Second added task");
        using var saved = new MemoryStream(); editing.Save(saved, Native(ProjectFileFormat.Mpp9, true));
        using var reopened = ProjectDocument.Load(new MemoryStream(saved.ToArray()));
        Assert.Equal("Added task", reopened.Tasks.GetByUid(added.Uid).Name);
        Assert.Contains(reopened.AllTasks, t => t.Name == "Second added task");
        Assert.DoesNotContain(reopened.AllTasks, t => t.Uid == removed.Uid);
    }

    [Theory]
    [InlineData("delivery")]
    [InlineData("empty")]
    [InlineData("actuals")]
    [InlineData("calendars")]
    [InlineData("custom-fields")]
    [InlineData("relationships")]
    [InlineData("resources")]
    public void Project2007ExportRetainsItsGenerationAndProducerBytes(string name) {
        string path = Path.Combine(AppContext.BaseDirectory, "Fixtures", "Project2024Mpp12", name + ".mpp");
        using var document = ProjectDocument.Load(path);
        Assert.Equal("MPP12", document.NativeInfo!.Generation);
        using var output = new MemoryStream(); document.Save(output);
        Assert.Equal(File.ReadAllBytes(path), output.ToArray());
        Assert.All(document.AllTasks, task => { Assert.False(task.IsManual); Assert.True(task.IsActive); });
        document.Title = "MPP12 metadata edit";
        document.Save(output);
        using var reopened = ProjectDocument.Load(new MemoryStream(output.ToArray()));
        Assert.Equal("MPP12", reopened.NativeInfo!.Generation);
        Assert.Equal("MPP12 metadata edit", reopened.Title);
        Assert.Equal(document.AllTasks.Select(t => (t.Uid, t.Name, t.Parent?.Uid, t.Start, t.Finish)),
            reopened.AllTasks.Select(t => (t.Uid, t.Name, t.Parent?.Uid, t.Start, t.Finish)));
    }

    [Theory]
    [InlineData(ProjectFileFormat.Mpp8, false)]
    [InlineData(ProjectFileFormat.Mpt8, true)]
    [InlineData(ProjectFileFormat.Mpp9, false)]
    [InlineData(ProjectFileFormat.Mpt9, true)]
    [InlineData(ProjectFileFormat.Mpp12, false)]
    [InlineData(ProjectFileFormat.Mpt12, true)]
    public void SeedFreeLegacyAuthoringUsesLegacyStorage(ProjectFileFormat format, bool template) {
        using var document = ProjectNativeAuthoringTests.Create();
        bool legacy8 = format == ProjectFileFormat.Mpp8 || format == ProjectFileFormat.Mpt8;
        bool legacy9 = legacy8 || format == ProjectFileFormat.Mpp9 || format == ProjectFileFormat.Mpt9;
        string generation = legacy8 ? "MPP8" : legacy9 ? "MPP9" : "MPP12";
        using var output = new MemoryStream(); document.Save(output, Native(format, legacy9));
        using var read = ProjectDocument.Load(new MemoryStream(output.ToArray()));
        Assert.Equal(generation, read.NativeInfo!.Generation); Assert.Equal(template, read.NativeInfo.IsTemplate);
        Assert.Equal("Design café / Łódź / 日本語", read.Tasks.GetByUid(2).Name);
        Assert.Equal(800m, read.Tasks.GetByUid(2).Baselines.Single().Cost);
        Assert.Equal(480m, read.Assignments.Single().Work!.Value.Minutes);
        if (legacy9) {
            Assert.Contains(read.Calendar!.Exceptions, e => e.FromDate == new DateTime(2026, 10, 12) && e.IsWorking == false);
            Assert.All(read.Calendar.Exceptions, e => Assert.Null(e.Name));
            Assert.Empty(read.Calendar.WorkWeeks);
        }
        else Assert.Equal("Maintenance", read.Calendar!.Exceptions.Single().Name);
        using var instance = ProjectDocument.CreateFromTemplate(new MemoryStream(output.ToArray()));
        using var instanceBytes = new MemoryStream(); instance.Save(instanceBytes);
        using var instanceRead = ProjectDocument.Load(new MemoryStream(instanceBytes.ToArray()));
        Assert.Equal(generation, instanceRead.NativeInfo!.Generation); Assert.False(instanceRead.NativeInfo.IsTemplate);
    }

    [Theory]
    [InlineData(ProjectFileFormat.Mpp14, ProjectFileFormat.Mpp12)]
    [InlineData(ProjectFileFormat.Mpp12, ProjectFileFormat.Mpp14)]
    public void GenerationConversionRebuildsStorageAndRequiresExplicitLossAcceptance(ProjectFileFormat source, ProjectFileFormat target) {
        using var document = ProjectNativeAuthoringTests.Create();
        using var first = new MemoryStream(); document.Save(first, Native(source));
        using var next = new MemoryStream();
        Assert.Contains(document.AssessSave(Native(target)).Diagnostics, d => d.Code == "PROJECT_NATIVE_GENERATION_LOSS");
        Assert.Throws<InvalidOperationException>(() => document.Save(next, Native(target))); Assert.Empty(next.ToArray());
        document.Save(next, Native(target, true));
        using var read = ProjectDocument.Load(new MemoryStream(next.ToArray()));
        Assert.Equal(target == ProjectFileFormat.Mpp12 ? "MPP12" : "MPP14", read.NativeInfo!.Generation);
        Assert.Equal("Design café / Łódź / 日本語", read.Tasks.GetByUid(2).Name);
        Assert.Equal(800m, read.Tasks.GetByUid(2).Baselines.Single().Cost);
    }

    [Fact]
    public void LegacyCreationReportsManualAndInactiveTaskLoss() {
        using var document = ProjectNativeAuthoringTests.Create();
        document.Tasks.GetByUid(2).IsManual = true; document.Tasks.GetByUid(2).IsActive = false;
        var assessment = document.AssessSave(Native(ProjectFileFormat.Mpp12));
        Assert.Contains(assessment.Diagnostics, d => d.RepresentsLoss && d.Location.EndsWith("/IsManual", StringComparison.Ordinal));
        Assert.Contains(assessment.Diagnostics, d => d.RepresentsLoss && d.Location.EndsWith("/IsActive", StringComparison.Ordinal));
        using var output = new MemoryStream(); Assert.Throws<InvalidOperationException>(() => document.Save(output, Native(ProjectFileFormat.Mpp12)));
        Assert.Empty(output.ToArray());
        document.Save(output, Native(ProjectFileFormat.Mpp12, true));
        using var read = ProjectDocument.Load(new MemoryStream(output.ToArray()));
        Assert.False(read.Tasks.GetByUid(2).IsManual); Assert.True(read.Tasks.GetByUid(2).IsActive);
    }
}

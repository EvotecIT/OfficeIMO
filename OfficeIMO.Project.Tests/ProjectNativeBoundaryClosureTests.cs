using System.Collections.Generic;
using OfficeIMO.Core.Internal;

namespace OfficeIMO.Project.Tests;

public sealed class ProjectNativeBoundaryClosureTests {
    [Fact]
    public void NativeVariableValuesShareOneConfiguredBudgetAcrossTables() {
        byte[] source = File.ReadAllBytes(ProjectNativeTests.Fixture("delivery.mpp"));
        var compound = Compound(source);
        int count = new[] { "Cal", "Task", "Rsc", "Assn", "Cons" }
            .Sum(table => BitConverter.ToInt32(compound.Streams["   114/TBknd" + table + "/VarMeta"], 8));
        Assert.True(count > 1);
        using var accepted = ProjectDocument.Load(new MemoryStream(source), new ProjectLoadOptions { MaxNativeValues = count });
        var error = Assert.Throws<InvalidDataException>(() => ProjectDocument.Load(new MemoryStream(source),
            new ProjectLoadOptions { MaxNativeValues = count - 1 }));
        Assert.Contains("variable value budget", error.Message);
        Assert.Throws<ArgumentOutOfRangeException>(() => ProjectDocument.Load(new MemoryStream(source),
            new ProjectLoadOptions { MaxNativeValues = 0 }));
    }

    [Fact]
    public void UnmappedNativeVariableFieldsFailBeforeTheirDataOffsetsAreRead() {
        byte[] source = File.ReadAllBytes(ProjectNativeTests.Fixture("delivery.mpp"));
        var compound = Compound(source);
        const string path = "   114/TBkndTask/VarMeta";
        byte[] metadata = (byte[])compound.Streams[path].Clone();
        Assert.True(BitConverter.ToInt32(metadata, 8) > 0);
        Buffer.BlockCopy(BitConverter.GetBytes(int.MaxValue), 0, metadata, 28, 4);
        Buffer.BlockCopy(BitConverter.GetBytes(uint.MaxValue), 0, metadata, 32, 4);
        byte[] corrupt = OfficeCompoundFileWriter.Rewrite(compound, new Dictionary<string, byte[]> { [path] = metadata });
        var error = Assert.Throws<InvalidDataException>(() => ProjectDocument.Load(new MemoryStream(corrupt)));
        Assert.Contains("no storage mapping", error.Message);
    }

    [Fact]
    public void NegativeNativeAssignmentRecordsDoNotEnterTheTypedOrXmlModels() {
        byte[] source = NewNative(ProjectFileFormat.Mpp14);
        var compound = Compound(source);
        var profile = ProjectNativeProfile.Detect(compound);
        var properties = ProjectNativeProperties.Read(compound.Streams[profile.Properties], default);
        var table = new ProjectNativeTable(compound, "TBkndAssn", properties[0x03000017], properties[0x00020017],
            int.MaxValue, default, profile);
        var record = Assert.Single(table.Records); var field = table.Fields[0x0f400000]; Assert.Equal(10, field.Source);
        string prefix = profile.DataRoot + "/TBkndAssn/";
        byte[] metadata = compound.Streams[prefix + "FixedMeta"];
        byte[] data = (byte[])compound.Streams[prefix + "FixedData"].Clone();
        int recordOffset = BitConverter.ToInt32(metadata, 16 + record.MetadataIndex * table.MetadataWidth + 4);
        Buffer.BlockCopy(BitConverter.GetBytes(-1), 0, data, recordOffset + field.Offset, 4);
        byte[] corrupt = OfficeCompoundFileWriter.Rewrite(compound, new Dictionary<string, byte[]> { [prefix + "FixedData"] = data });
        using var document = ProjectDocument.Load(new MemoryStream(corrupt));
        Assert.Empty(document.Assignments); Assert.NotEmpty(document.AllTasks); Assert.NotEmpty(document.Resources);
        using var xml = new MemoryStream();
        document.Save(xml, new ProjectSaveOptions { Format = ProjectFileFormat.Xml, LossPolicy = OfficeConversionLossPolicy.Allow });
        using var reopened = ProjectDocument.Load(new MemoryStream(xml.ToArray())); Assert.Empty(reopened.Assignments);
    }

    [Theory]
    [InlineData(ProjectFileFormat.Mpp8)]
    [InlineData(ProjectFileFormat.Mpp9)]
    [InlineData(ProjectFileFormat.Mpp12)]
    [InlineData(ProjectFileFormat.Mpp14)]
    public void NewNativeOutputDoesNotLeakStagingProperties(ProjectFileFormat format) {
        using var document = ProjectNativeAuthoringTests.Create();
        document.Name = null; document.Title = null; document.Settings.ScheduleFromStart = null;
        using var output = new MemoryStream(); document.Save(output, Native(format));
        using var reopened = ProjectDocument.Load(new MemoryStream(output.ToArray()));
        Assert.Null(reopened.Name); Assert.Null(reopened.Title); Assert.Null(reopened.Settings.ScheduleFromStart);
    }

    [Fact]
    public void MissingNativeSchedulingDirectionRemainsAbsentThroughEditsAndXmlConversion() {
        byte[] source = RewriteProperties(NewNative(ProjectFileFormat.Mpp14), properties => properties.Remove(0x02400004));
        using var document = ProjectDocument.Load(new MemoryStream(source));
        Assert.Null(document.Settings.ScheduleFromStart);
        Assert.DoesNotContain("ScheduleFromStart", document.ToXml(new ProjectSaveOptions {
            Format = ProjectFileFormat.Xml, LossPolicy = OfficeConversionLossPolicy.Allow
        }));
        document.Tasks.First(task => task.Uid > 0).Name = "Edited without direction";
        using var output = new MemoryStream(); document.Save(output, Native(ProjectFileFormat.Mpp14));
        using var reopened = ProjectDocument.Load(new MemoryStream(output.ToArray()));
        Assert.Null(reopened.Settings.ScheduleFromStart);
    }

    [Fact]
    public void UnknownNativeSchedulingDirectionIsRejected() {
        byte[] source = RewriteProperties(NewNative(ProjectFileFormat.Mpp14),
            properties => properties.Set(0x02400004, 2, BitConverter.GetBytes((short)2)));
        var error = Assert.Throws<InvalidDataException>(() => ProjectDocument.Load(new MemoryStream(source)));
        Assert.Contains("scheduling direction", error.Message);
    }

    private static ProjectSaveOptions Native(ProjectFileFormat format) => new() {
        Format = format, LossPolicy = OfficeConversionLossPolicy.Allow
    };

    private static byte[] NewNative(ProjectFileFormat format) {
        using var document = ProjectNativeAuthoringTests.Create();
        using var output = new MemoryStream(); document.Save(output, Native(format)); return output.ToArray();
    }

    private static byte[] RewriteProperties(byte[] source, Action<ProjectNativePropertySet> change) {
        var compound = Compound(source); var profile = ProjectNativeProfile.Detect(compound);
        var properties = new ProjectNativePropertySet(compound.Streams[profile.Properties], default, profile == ProjectNativeProfile.Mpp8);
        change(properties);
        return OfficeCompoundFileWriter.Rewrite(compound, new Dictionary<string, byte[]> {
            [profile.Properties] = properties.Serialize(source.Length * 2L, default)
        });
    }

    private static OfficeCompoundFile Compound(byte[] source) {
        Assert.True(OfficeCompoundFileReader.TryRead(source, out OfficeCompoundFile? compound, out var error), error);
        return compound!;
    }
}

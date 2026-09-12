using System.Collections.Generic;
using OfficeIMO.Core.Internal;
using System.Xml.Linq;

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

    [Theory]
    [InlineData(ProjectFileFormat.Mpp8)]
    [InlineData(ProjectFileFormat.Mpp9)]
    [InlineData(ProjectFileFormat.Mpp12)]
    [InlineData(ProjectFileFormat.Mpp14)]
    public void NativeVariableValueBudgetAppliesToEverySupportedGeneration(ProjectFileFormat format) {
        byte[] source = NewNative(format);
        using var accepted = ProjectDocument.Load(new MemoryStream(source));
        var error = Assert.Throws<InvalidDataException>(() => ProjectDocument.Load(new MemoryStream(source),
            new ProjectLoadOptions { MaxNativeValues = 1 }));
        Assert.Contains("variable value budget", error.Message);
    }

    [Fact]
    public void Mpp9SourceZeroVariableFieldsRoundTrip() {
        using var document = ProjectNativeAuthoringTests.Create();
        document.Tasks.GetByUid(2).Wbs = "DELIVERY.DESIGN";
        using var output = new MemoryStream(); document.Save(output, Native(ProjectFileFormat.Mpp9));
        using var reopened = ProjectDocument.Load(new MemoryStream(output.ToArray()));
        Assert.Equal("DELIVERY.DESIGN", reopened.Tasks.GetByUid(2).Wbs);
    }

    [Theory]
    [InlineData(ProjectFileFormat.Mpp8)]
    [InlineData(ProjectFileFormat.Mpp9)]
    [InlineData(ProjectFileFormat.Mpp12)]
    [InlineData(ProjectFileFormat.Mpp14)]
    public void GeneratedNativeProjectSummaryDoesNotEnterTheTypedModel(ProjectFileFormat format) {
        using var document = ProjectNativeAuthoringTests.Create();
        int[] expected = document.AllTasks.Select(task => task.Uid).ToArray(); Assert.DoesNotContain(0, expected);
        using var output = new MemoryStream(); document.Save(output, Native(format));
        using var reopened = ProjectDocument.Load(new MemoryStream(output.ToArray()));
        Assert.Equal(expected, reopened.AllTasks.Select(task => task.Uid));
    }

    [Theory]
    [InlineData(ProjectFileFormat.Mpp12)]
    [InlineData(ProjectFileFormat.Mpp14)]
    public void GeneratedNativeIdentitiesDoNotMaterializeAbsentPublicGuids(ProjectFileFormat format) {
        using var document = ProjectNativeAuthoringTests.Create();
        var explicitGuid = new Guid("57e42c89-4898-47e8-b184-d9130eb2d3bb");
        document.Tasks.GetByUid(2).Guid = explicitGuid;
        Assert.Null(document.Guid);
        Assert.All(document.AllTasks.Where(task => task.Uid != 2), task => Assert.Null(task.Guid));
        Assert.All(document.Resources, resource => Assert.Null(resource.Guid));
        Assert.All(document.Calendars, calendar => Assert.Null(calendar.Guid));
        Assert.All(document.Assignments, assignment => Assert.Null(assignment.Guid));

        using var output = new MemoryStream(); document.Save(output, Native(format));
        using var reopened = ProjectDocument.Load(new MemoryStream(output.ToArray()));

        Assert.Null(reopened.Guid);
        Assert.Equal(explicitGuid, reopened.Tasks.GetByUid(2).Guid);
        Assert.All(reopened.AllTasks.Where(task => task.Uid != 2), task => Assert.Null(task.Guid));
        Assert.All(reopened.Resources, resource => Assert.Null(resource.Guid));
        Assert.All(reopened.Calendars, calendar => Assert.Null(calendar.Guid));
        Assert.All(reopened.Assignments, assignment => Assert.Null(assignment.Guid));

        var later = reopened.Tasks.GetByUid(1).Children.Add("Later"); later.Duration = ProjectDuration.WorkingDays(1);
        var laterAssignment = reopened.Assignments.Add(later, reopened.Resources.GetByUid(1)); laterAssignment.Units = ProjectUnits.Percent(100);
        using var second = new MemoryStream(); reopened.Save(second, Native(format));
        using var twiceReopened = ProjectDocument.Load(new MemoryStream(second.ToArray()));
        Assert.Null(twiceReopened.Tasks.GetByUid(later.Uid).Guid);
        Assert.Null(twiceReopened.Assignments.GetByUid(laterAssignment.Uid).Guid);
    }

    [Theory]
    [InlineData(ProjectFileFormat.Mpp12)]
    [InlineData(ProjectFileFormat.Mpp14)]
    public void GeneratedNativeProjectSummaryTracksAnExplicitProjectGuid(ProjectFileFormat format) {
        using var document = ProjectNativeAuthoringTests.Create();
        document.Guid = new Guid("429757aa-d385-46b1-9161-8a6046c584a4");
        using var first = new MemoryStream(); document.Save(first, Native(format));
        document.Guid = new Guid("01c67d06-90f8-49a9-bfbd-c5b0665d84af");
        using var second = new MemoryStream(); document.Save(second, Native(format));
        using var reopened = ProjectDocument.Load(new MemoryStream(second.ToArray()));
        Assert.DoesNotContain(reopened.AllTasks, task => task.Uid == 0);
    }

    [Theory]
    [InlineData(ProjectFileFormat.Mpp8)]
    [InlineData(ProjectFileFormat.Mpp9)]
    [InlineData(ProjectFileFormat.Mpp12)]
    [InlineData(ProjectFileFormat.Mpp14)]
    public void GeneratedNativeProjectSummaryTracksAProjectRename(ProjectFileFormat format) {
        using var document = ProjectNativeAuthoringTests.Create();
        using var first = new MemoryStream(); document.Save(first, Native(format));
        document.Name = "Renamed native project";
        using var second = new MemoryStream(); document.Save(second, Native(format));
        using var reopened = ProjectDocument.Load(new MemoryStream(second.ToArray()));
        Assert.Equal("Renamed native project", reopened.Name);
        Assert.DoesNotContain(reopened.AllTasks, task => task.Uid == 0);
    }

    [Theory]
    [InlineData(ProjectFileFormat.Mpp8)]
    [InlineData(ProjectFileFormat.Mpp9)]
    [InlineData(ProjectFileFormat.Mpp12)]
    [InlineData(ProjectFileFormat.Mpp14)]
    public void ExplicitProjectSummaryRemainsInTheTypedModel(ProjectFileFormat format) {
        using var document = ExplicitSummary();
        using var output = new MemoryStream(); document.Save(output, Native(format));
        using var reopened = ProjectDocument.Load(new MemoryStream(output.ToArray()));
        Assert.Equal("Explicit project summary", reopened.Tasks.GetByUid(0).Name);
    }

    [Theory]
    [InlineData(ProjectFileFormat.Mpp8)]
    [InlineData(ProjectFileFormat.Mpp9)]
    [InlineData(ProjectFileFormat.Mpp12)]
    [InlineData(ProjectFileFormat.Mpp14)]
    public void ExplicitNativeProjectSummaryDeletionIsRejected(ProjectFileFormat format) {
        using var source = ExplicitSummary(); using var output = new MemoryStream(); source.Save(output, Native(format));
        using var document = ProjectDocument.Load(new MemoryStream(output.ToArray()));
        Assert.True(document.Tasks.Remove(document.Tasks.GetByUid(0)));
        var report = document.AssessSave(Native(format));
        Assert.Contains(report.Diagnostics, diagnostic => diagnostic.Code == "PROJECT_NATIVE_PROJECT_SUMMARY_REQUIRED"
            && diagnostic.Severity == ProjectDiagnosticSeverity.Error);
        Assert.Throws<InvalidDataException>(() => document.Save(new MemoryStream(), Native(format)));
    }

    [Theory]
    [InlineData(ProjectFileFormat.Mpp8)]
    [InlineData(ProjectFileFormat.Mpp9)]
    [InlineData(ProjectFileFormat.Mpp12)]
    [InlineData(ProjectFileFormat.Mpp14)]
    public void ReservedNativeResourceDeletionIsRejected(ProjectFileFormat format) {
        using var source = ProjectNativeAuthoringTests.Create(); using var output = new MemoryStream(); source.Save(output, Native(format));
        using var document = ProjectDocument.Load(new MemoryStream(output.ToArray()));
        Assert.True(document.Resources.Remove(document.Resources.GetByUid(0)));
        var report = document.AssessSave(Native(format));
        Assert.Contains(report.Diagnostics, diagnostic => diagnostic.Code == "PROJECT_NATIVE_RESERVED_RESOURCE_REQUIRED"
            && diagnostic.Severity == ProjectDiagnosticSeverity.Error);
        Assert.Throws<InvalidDataException>(() => document.Save(new MemoryStream(), Native(format)));
    }

    [Theory]
    [InlineData(ProjectFileFormat.Mpp8)]
    [InlineData(ProjectFileFormat.Mpp9)]
    [InlineData(ProjectFileFormat.Mpp12)]
    [InlineData(ProjectFileFormat.Mpp14)]
    public void NativeCalendarKindUsesItsBaseReferenceWhenTheFlagIsAbsent(ProjectFileFormat format) {
        using var document = ProjectNativeAuthoringTests.Create();
        var resource = document.Resources.First(item => item.Uid > 0); var derived = resource.Calendar!;
        Assert.NotNull(derived.BaseCalendar); derived.IsBaseCalendar = null;
        using var output = new MemoryStream(); document.Save(output, Native(format));
        using var reopened = ProjectDocument.Load(new MemoryStream(output.ToArray()));
        var copied = reopened.Resources.GetByUid(resource.Uid).Calendar!;
        Assert.False(copied.IsBaseCalendar); Assert.NotNull(copied.BaseCalendar);
    }

    [Theory]
    [InlineData(ProjectFileFormat.Mpp8)]
    [InlineData(ProjectFileFormat.Mpp9)]
    [InlineData(ProjectFileFormat.Mpp12)]
    [InlineData(ProjectFileFormat.Mpp14)]
    public void ResourceCalendarWithoutBaseOrExplicitDerivedKindIsRejected(ProjectFileFormat format) {
        using var document = ProjectNativeAuthoringTests.Create();
        var calendar = document.Resources.First(item => item.Uid > 0).Calendar!;
        calendar.BaseCalendar = null; calendar.IsBaseCalendar = null;
        var report = document.AssessSave(Native(format));
        Assert.Contains(report.Diagnostics, diagnostic => diagnostic.Code == "PROJECT_NATIVE_RESOURCE_CALENDAR"
            && diagnostic.Severity == ProjectDiagnosticSeverity.Error);
        Assert.Throws<InvalidDataException>(() => document.Save(new MemoryStream(), Native(format)));
    }

    [Theory]
    [InlineData(ProjectFileFormat.Mpp8)]
    [InlineData(ProjectFileFormat.Mpp9)]
    [InlineData(ProjectFileFormat.Mpp12)]
    [InlineData(ProjectFileFormat.Mpp14)]
    public void CrossProjectDependenciesAreRejectedBeforeNativeSerialization(ProjectFileFormat format) {
        using var document = ProjectDocument.Load(ProjectResourceCapacityTests.Fixture("external-consumer"));
        var report = document.AssessSave(Native(format));
        Assert.Contains(report.Diagnostics, diagnostic => diagnostic.Code == "PROJECT_NATIVE_CROSS_PROJECT_DEPENDENCY"
            && diagnostic.Severity == ProjectDiagnosticSeverity.Error);
        Assert.Throws<InvalidDataException>(() => document.Save(new MemoryStream(), Native(format)));
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
        Assert.Contains("storage mapping", error.Message);
    }

    [Fact]
    public void FixedNativeFieldsCannotMasqueradeAsVariableStorage() {
        byte[] source = NewNative(ProjectFileFormat.Mpp14);
        var compound = Compound(source); var profile = ProjectNativeProfile.Detect(compound);
        var properties = ProjectNativeProperties.Read(compound.Streams[profile.Properties], default);
        var table = new ProjectNativeTable(compound, "TBkndTask", properties[0x03000014], properties[0x00020014],
            int.MaxValue, default, profile);
        uint fixedField = table.Fields.Values.First(field => field.Source == 10).Id;
        string path = profile.DataRoot + "/TBkndTask/VarMeta";
        byte[] metadata = (byte[])compound.Streams[path].Clone(); Assert.True(BitConverter.ToInt32(metadata, 8) > 0);
        Buffer.BlockCopy(BitConverter.GetBytes(int.MaxValue), 0, metadata, 28, 4);
        Buffer.BlockCopy(BitConverter.GetBytes(fixedField), 0, metadata, 32, 4);
        byte[] corrupt = OfficeCompoundFileWriter.Rewrite(compound, new Dictionary<string, byte[]> { [path] = metadata });
        var error = Assert.Throws<InvalidDataException>(() => ProjectDocument.Load(new MemoryStream(corrupt)));
        Assert.Contains("qualified variable storage mapping", error.Message);
    }

    [Fact]
    public void NegativeNativeAssignmentRecordsDoNotEnterTheTypedOrXmlModels() {
        byte[] source = NewNative(ProjectFileFormat.Mpp14);
        byte[] corrupt = RewriteFixedUid(source, "Assn", 0x17, 0x0f400000);
        using var document = ProjectDocument.Load(new MemoryStream(corrupt));
        Assert.Empty(document.Assignments); Assert.NotEmpty(document.AllTasks); Assert.NotEmpty(document.Resources);
        using var xml = new MemoryStream();
        document.Save(xml, new ProjectSaveOptions { Format = ProjectFileFormat.Xml, LossPolicy = OfficeConversionLossPolicy.Allow });
        using var reopened = ProjectDocument.Load(new MemoryStream(xml.ToArray())); Assert.Empty(reopened.Assignments);
    }

    [Fact]
    public void NegativeNativeDependencyRecordsDoNotEnterTheTypedOrXmlModels() {
        using var sourceDocument = ProjectNativeAuthoringTests.Create();
        var predecessor = sourceDocument.Tasks.GetByUid(2); var successor = sourceDocument.Tasks.Add("Successor");
        sourceDocument.Dependencies.Add(predecessor, successor);
        using var source = new MemoryStream(); sourceDocument.Save(source, Native(ProjectFileFormat.Mpp14));
        byte[] corrupt = RewriteFixedUid(source.ToArray(), "Cons", 0x18, 0x0e400000);
        using var document = ProjectDocument.Load(new MemoryStream(corrupt)); Assert.Empty(document.Dependencies);
        using var xml = new MemoryStream();
        document.Save(xml, new ProjectSaveOptions { Format = ProjectFileFormat.Xml, LossPolicy = OfficeConversionLossPolicy.Allow });
        using var reopened = ProjectDocument.Load(new MemoryStream(xml.ToArray())); Assert.Empty(reopened.Dependencies);
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

    [Theory]
    [InlineData(ProjectFileFormat.Mpp8)]
    [InlineData(ProjectFileFormat.Mpp9)]
    [InlineData(ProjectFileFormat.Mpp12)]
    [InlineData(ProjectFileFormat.Mpp14)]
    public void MissingNativeDefaultClocksRemainAbsentThroughEditsAndXmlConversion(ProjectFileFormat format) {
        byte[] source = RewriteProperties(NewNative(format), properties => {
            properties.Remove(0x0240001c);
            properties.Remove(0x02400021);
        });
        using var document = ProjectDocument.Load(new MemoryStream(source));
        Assert.Null(document.Settings.DefaultStartTime);
        Assert.Null(document.Settings.DefaultFinishTime);
        string xml = document.ToXml(new ProjectSaveOptions { Format = ProjectFileFormat.Xml, LossPolicy = OfficeConversionLossPolicy.Allow });
        Assert.DoesNotContain("DefaultStartTime", xml);
        Assert.DoesNotContain("DefaultFinishTime", xml);
        document.Tasks.First(task => task.Uid > 0).Name = "Edited without default clocks";
        using var output = new MemoryStream(); document.Save(output, Native(format));
        using var reopened = ProjectDocument.Load(new MemoryStream(output.ToArray()));
        Assert.Null(reopened.Settings.DefaultStartTime);
        Assert.Null(reopened.Settings.DefaultFinishTime);
    }

    [Fact]
    public void UnknownNativeSchedulingDirectionIsRejected() {
        byte[] source = RewriteProperties(NewNative(ProjectFileFormat.Mpp14),
            properties => properties.Set(0x02400004, 2, BitConverter.GetBytes((short)2)));
        var error = Assert.Throws<InvalidDataException>(() => ProjectDocument.Load(new MemoryStream(source)));
        Assert.Contains("scheduling direction", error.Message);
    }

    [Fact]
    public void WrongWidthNativeSchedulingDirectionIsRejected() {
        byte[] source = RewriteProperties(NewNative(ProjectFileFormat.Mpp14),
            properties => properties.Set(0x02400004, 4, BitConverter.GetBytes(65536)));
        var error = Assert.Throws<InvalidDataException>(() => ProjectDocument.Load(new MemoryStream(source)));
        Assert.Contains("invalid width", error.Message);
    }

    private static ProjectSaveOptions Native(ProjectFileFormat format) => new() {
        Format = format, LossPolicy = OfficeConversionLossPolicy.Allow
    };

    private static ProjectDocument ExplicitSummary() {
        using var seed = ProjectNativeAuthoringTests.Create(); var xml = XDocument.Parse(seed.ToXml());
        XNamespace ns = XmlContracts.Ns; var tasks = xml.Root!.Element(ns + "Tasks")!;
        tasks.AddFirst(new XElement(ns + "Task", new XElement(ns + "UID", 0), new XElement(ns + "ID", 0),
            new XElement(ns + "Name", "Explicit project summary"), new XElement(ns + "OutlineLevel", 0), new XElement(ns + "Summary", 1)));
        return ProjectDocument.Parse(xml.ToString(SaveOptions.DisableFormatting));
    }

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

    private static byte[] RewriteFixedUid(byte[] source, string tableName, uint tableId, uint uidField) {
        var compound = Compound(source); var profile = ProjectNativeProfile.Detect(compound);
        var properties = ProjectNativeProperties.Read(compound.Streams[profile.Properties], default);
        var table = new ProjectNativeTable(compound, "TBknd" + tableName, properties[0x03000000 | tableId],
            properties[0x00020000 | tableId], int.MaxValue, default, profile);
        var record = Assert.Single(table.Records); var field = table.Fields[uidField]; Assert.Equal(10, field.Source);
        string prefix = profile.DataRoot + "/TBknd" + tableName + "/";
        byte[] metadata = compound.Streams[prefix + "FixedMeta"];
        byte[] data = (byte[])compound.Streams[prefix + "FixedData"].Clone();
        int recordOffset = BitConverter.ToInt32(metadata, 16 + record.MetadataIndex * table.MetadataWidth + 4);
        Buffer.BlockCopy(BitConverter.GetBytes(-1), 0, data, recordOffset + field.Offset, 4);
        return OfficeCompoundFileWriter.Rewrite(compound, new Dictionary<string, byte[]> { [prefix + "FixedData"] = data });
    }

    private static OfficeCompoundFile Compound(byte[] source) {
        Assert.True(OfficeCompoundFileReader.TryRead(source, out OfficeCompoundFile? compound, out var error), error);
        return compound!;
    }
}

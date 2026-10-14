using System.IO;
using OfficeIMO.Core.Internal;

namespace OfficeIMO.Security.Tests;

public sealed class OfficeVbaStreamInspectionTests {
    [Theory]
    [InlineData(false, false)]
    [InlineData(true, false)]
    [InlineData(false, true)]
    [InlineData(true, true)]
    public void StreamInspectionMatchesNativePathsWithAnyDictionaryComparer(bool changeDirectoryCase, bool changeModuleCase) {
        var project = OfficeVbaProject.Create("Inventory");
        project.AddModule("Helpers", "'source\r\n");
        project.AddRegisteredReference("Office", new Guid("2DF8D04C-5BFA-101B-BDE5-00AA0044DE52"), 2, 8);
        var streams = ReadStreams(project.Write().GetBytes()).ToDictionary(
            pair => pair.Key == "VBA/dir" && changeDirectoryCase ? "vBa/DiR"
                : pair.Key == "VBA/Helpers" && changeModuleCase ? "vBa/hElPeRs" : pair.Key,
            pair => pair.Value, StringComparer.Ordinal);

        var inspection = OfficeVbaProjectInspector.Inspect(streams, 20000);

        Assert.Null(inspection.Limitation);
        Assert.Equal("Inventory", inspection.Name);
        Assert.Equal(project.GetModule("Helpers").Source, Assert.Single(inspection.Modules).Source);
        Assert.Equal("Office", Assert.Single(inspection.References).Name);
    }

    [Theory]
    [InlineData("VBA/dir", "vba/DIR")]
    [InlineData("VBA/Helpers", "vba/HELPERS")]
    public void StreamInspectionRejectsAmbiguousCaseAliases(string canonicalPath, string alias) {
        var project = OfficeVbaProject.Create();
        project.AddModule("Helpers", "'source\r\n");
        var streams = ReadStreams(project.Write().GetBytes()).ToDictionary(pair => pair.Key, pair => pair.Value, StringComparer.Ordinal);
        streams.Add(alias, streams[canonicalPath]);

        var inspection = OfficeVbaProjectInspector.Inspect(streams, 20000);

        Assert.Contains("repeats a stream path", inspection.Limitation);
        Assert.Empty(inspection.Modules);
    }

    [Fact]
    public void StreamInspectionRetainsDeclaredModulesWhenSourceIsMissing() {
        var project = OfficeVbaProject.Create("Inventory");
        project.AddModule("Helpers", "Public Function Value() As Long\r\nValue = 42\r\nEnd Function\r\n");
        project.AddRegisteredReference("Office", new Guid("2DF8D04C-5BFA-101B-BDE5-00AA0044DE52"), 2, 8);
        var streams = ReadStreams(project.Write().GetBytes());
        streams.Remove("VBA/Helpers");

        var inspection = OfficeVbaProjectInspector.Inspect(streams, 64 * 1024 * 1024);

        Assert.Null(inspection.Limitation);
        Assert.Equal("Inventory", inspection.Name);
        var module = Assert.Single(inspection.Modules);
        Assert.Equal("Helpers", module.Name); Assert.Equal("Helpers", module.StreamName);
        Assert.True(module.IsProcedural); Assert.Null(module.Source); Assert.NotNull(module.Limitation);
        var reference = Assert.Single(inspection.References);
        Assert.Equal("Office", reference.Name); Assert.Equal((ushort)0x000d, reference.NativeKind);
        Assert.Contains("{2DF8D04C-5BFA-101B-BDE5-00AA0044DE52}", reference.LibraryId);
    }

    [Fact]
    public void StreamInspectionUsesOneExpandedBudgetAcrossDirectoryAndSources() {
        var project = OfficeVbaProject.Create();
        project.AddModule("First", "'" + new string('a', 1000));
        project.AddModule("Second", "'" + new string('b', 1000));
        var streams = ReadStreams(project.Write().GetBytes());
        Assert.True(OfficeVbaCompression.TryDecompress(streams["VBA/dir"], 10000, out byte[] directory, out _));
        int budget = directory.Length + Encoding.ASCII.GetByteCount(project.GetModule("First").Source) + 5;

        var inspection = OfficeVbaProjectInspector.Inspect(streams, budget);

        Assert.Null(inspection.Limitation);
        Assert.Equal(project.GetModule("First").Source, inspection.Modules[0].Source);
        Assert.Null(inspection.Modules[1].Source); Assert.NotNull(inspection.Modules[1].Limitation);
        Assert.NotNull(OfficeVbaProjectInspector.Inspect(streams, directory.Length - 1).Limitation);
        Assert.NotNull(OfficeVbaProjectInspector.Inspect(new Dictionary<string, byte[]>(), 100).Limitation);
    }

    [Fact]
    public void StreamInspectionDoesNotChargeSignatureTranscriptsToTheSourceBudget() {
        var project = OfficeVbaProject.Create("Inventory");
        project.AddRegisteredReference("Library", new Guid("11111111-2222-3333-4444-555555555555"), path: new string('x', 4000));
        var streams = ReadStreams(project.Write().GetBytes());
        Assert.True(OfficeVbaCompression.TryDecompress(streams["VBA/dir"], 20000, out byte[] directory, out _));
        var inspection = OfficeVbaProjectInspector.Inspect(streams, directory.Length);
        Assert.Null(inspection.Limitation); Assert.Single(inspection.References);
        Assert.NotNull(OfficeVbaProjectInspector.Inspect(streams, directory.Length - 1).Limitation);
    }

    [Fact]
    public void StreamInspectionRetainsInventoryForAnOverexpandedModuleChunk() {
        var project = OfficeVbaProject.Create("Inventory");
        project.AddModule("Helpers", "'original");
        var streams = ReadStreams(project.Write().GetBytes());
        streams["VBA/Helpers"] = new byte[] { 1, 4, 0xb0, 2, 0x41, 0xfc, 0x0f, 0x42 };
        var inspection = OfficeVbaProjectInspector.Inspect(streams, 20000);
        Assert.Null(inspection.Limitation);
        var module = Assert.Single(inspection.Modules);
        Assert.Equal("Helpers", module.Name); Assert.Null(module.Source); Assert.Contains("4096", module.Limitation);
    }

    [Fact]
    public void Windows1250SourceCanBeCreatedReadAndEditedWithStrictCharacterChecks() {
        const string text = "' Zażółć gęślą jaźń\r\n";
        var project = OfficeVbaProject.Create("Polish", 1250);
        project.AddModule("Helpers", text);
        byte[] bytes = project.Write().GetBytes();
        var loaded = OfficeVbaProject.Load(bytes);
        Assert.Equal(1250, loaded.CodePage); Assert.Contains(text, loaded.GetModule("Helpers").Source);
        var inspection = OfficeVbaProjectInspector.Inspect(ReadStreams(bytes), 10000);
        Assert.Null(inspection.Limitation); Assert.Equal(loaded.GetModule("Helpers").Source, Assert.Single(inspection.Modules).Source);
        loaded.SetModuleSource("Helpers", text + "' Łódź\r\n");
        Assert.Contains("Łódź", OfficeVbaProject.Load(loaded.Write().GetBytes()).GetModule("Helpers").Source);
        Assert.Throws<EncoderFallbackException>(() => loaded.SetModuleSource("Helpers", "' 漢字"));
        Assert.Contains("Łódź", loaded.GetModule("Helpers").Source);
    }

    private static Dictionary<string, byte[]> ReadStreams(byte[] bytes) {
        Assert.True(OfficeCompoundFileReader.TryRead(bytes, out OfficeCompoundFile? compound, out _));
        return compound!.Streams.ToDictionary(x => x.Key, x => x.Value, StringComparer.OrdinalIgnoreCase);
    }
}

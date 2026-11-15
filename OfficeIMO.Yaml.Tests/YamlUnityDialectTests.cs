using Xunit;

namespace OfficeIMO.Yaml.Tests;

// Unity's serialization files break the specification in two ways that the Unity dialect accepts.
public class YamlUnityDialectTests {
    private const string Scene =
        "%YAML 1.1\n" +
        "%TAG !u! tag:unity3d.com,2011:\n" +
        "--- !u!29 &1\n" +
        "OcclusionCullingSettings:\n" +
        "  m_ObjectHideFlags: 0\n" +
        "--- !u!1 &519420028\n" +
        "GameObject:\n" +
        "  m_Component:\n" +
        "  - component: {fileID: 519420032}\n" +
        "  m_Name: Main Camera\n" +
        "  m_IsActive: 1\n" +
        "--- !u!1 &700001 stripped\n" +
        "GameObject:\n" +
        "  m_CorrespondingSourceObject: {fileID: 100, guid: 9fc0d4010bbf28b4594072e72b8655ab, type: 3}\n" +
        "  m_PrefabInstance: {fileID: 700000}\n";

    private static readonly YamlLoadOptions Unity = new YamlLoadOptions { Dialect = YamlDialect.Unity };

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void StandardReadsRejectUnityFiles(bool strict) {
        YamlException error = Assert.Throws<YamlException>(() => YamlDocument.Parse(Scene, new YamlLoadOptions { Strict = strict }));
        Assert.Equal(YamlErrorCodes.InvalidTag, error.Code);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void UnityFilesLoadWithTagsAnchorsAndStrippedMarkers(bool strict) {
        using YamlDocument document = YamlDocument.Parse(Scene, new YamlLoadOptions { Dialect = YamlDialect.Unity, Strict = strict });

        Assert.Equal(3, document.Documents.Count);
        Assert.Equal(new[] { "tag:unity3d.com,2011:29", "tag:unity3d.com,2011:1", "tag:unity3d.com,2011:1" }, document.Documents.Select(p => p.Root.Tag));
        Assert.Equal(new[] { "1", "519420028", "700001" }, document.Documents.Select(p => p.Root.Anchor));
        Assert.Equal(new[] { false, false, true }, document.Documents.Select(p => p.IsStripped));
        // Each part reports only the directives it declares, and all of them read YAML 1.1.
        Assert.Equal(new[] { 1, 0, 0 }, document.Documents.Select(p => p.TagDirectives.Count));
        Assert.All(document.Documents, p => Assert.Equal(YamlSchema.Yaml11, p.Schema));

        YamlMapping camera = document.Documents[1].Root.AsMapping()["GameObject"].AsMapping();
        Assert.Equal("Main Camera", camera["m_Name"].AsString());
        Assert.Equal(519420032L, camera.Document!.Documents[1].Root.AsMapping()["GameObject"].AsMapping()["m_Component"].AsSequence()[0].AsMapping()["component"].AsMapping()["fileID"].GetValue<long>());
        Assert.Equal(Scene, document.ToString());
    }

    [Fact]
    public void EditsChangeOnlyTheirLinesAndKeepTheStrippedMarker() {
        using YamlDocument document = YamlDocument.Parse(Scene, Unity);
        document.Documents[1].Root.AsMapping()["GameObject"].AsMapping().Set("m_Name", "Camera 2");
        document.Documents[2].Root.AsMapping()["GameObject"].AsMapping()["m_PrefabInstance"].AsMapping().Set("fileID", 700002);

        string expected = Scene.Replace("m_Name: Main Camera", "m_Name: Camera 2").Replace("{fileID: 700000}", "{fileID: 700002}");
        Assert.Equal(expected, document.ToString());
    }

    [Fact]
    public void RenamingAStrippedAnchorKeepsTheMarker() {
        using YamlDocument document = YamlDocument.Parse(Scene, Unity);
        document.Documents[2].Root.SetAnchor("700009");

        Assert.Contains("--- !u!1 &700009 stripped\n", document.ToString());
    }

    [Fact]
    public void EditedTagsUseTheDocumentsTagHandles() {
        using YamlDocument document = YamlDocument.Parse("%TAG !e! tag:example.com,2026:\n--- !e!item &a\nname: x\n");
        document.Root!.SetTag("tag:example.com,2026:widget");

        Assert.Equal("%TAG !e! tag:example.com,2026:\n--- !e!widget &a\nname: x\n", document.ToString());
    }

    [Fact]
    public void ReaderReportsStrippedNodes() {
        using var reader = new YamlReader(Scene, new YamlReaderOptions { Dialect = YamlDialect.Unity });
        var stripped = new List<string>();
        while (reader.Read()) {
            if (reader.IsStripped) stripped.Add(reader.Anchor.ToString());
        }
        Assert.Equal(new[] { "700001" }, stripped);
    }

    [Theory]
    [InlineData("--- !u!1 &5\nstripped: 1\n")]                  // a key on the next line, not a marker
    [InlineData("--- !u!1 &5 strippedX\nGameObject: {}\n")]       // a longer word
    [InlineData("a: &5 stripped\n")]                             // not on a '---' line: an ordinary scalar
    public void StrippedIsOnlyAMarkerAtTheEndOfTheDocumentStartLine(string body) {
        string yaml = "%TAG !u! tag:unity3d.com,2011:\n" + body;
        bool standard = TryParse(yaml, new YamlLoadOptions(), out bool standardStripped);
        bool unity = TryParse(yaml, Unity, out bool unityStripped);

        Assert.Equal(standard, unity);
        Assert.False(standardStripped);
        Assert.False(unityStripped);
    }

    [Fact]
    public void StrippedMayCarryATrailingComment() {
        using YamlDocument document = YamlDocument.Parse("%TAG !u! tag:unity3d.com,2011:\n--- !u!1 &5 stripped # placeholder\nGameObject: {}\n", Unity);
        Assert.True(document.Documents[0].IsStripped);
    }

    [Fact]
    public void ALaterTagDirectiveReplacesTheInheritedOnes() {
        const string yaml = "%TAG !u! tag:unity3d.com,2011:\n--- !u!1 &1\na: 1\n...\n%TAG !x! tag:example.com,2026:\n--- !x!2 &2\nb: 2\n--- !u!3 &3\nc: 3\n";
        YamlException error = Assert.Throws<YamlException>(() => YamlDocument.Parse(yaml, Unity));
        Assert.Equal(YamlErrorCodes.InvalidTag, error.Code);
    }

    [Fact]
    public void UnityFilesCannotBeReformatted() {
        using YamlDocument document = YamlDocument.Parse(Scene, Unity);
        var output = new MemoryStream();
        Assert.Throws<NotSupportedException>(() => document.Save(output, new YamlSaveOptions { Reformat = true }));
    }

    private static bool TryParse(string yaml, YamlLoadOptions options, out bool stripped) {
        stripped = false;
        try {
            using YamlDocument document = YamlDocument.Parse(yaml, options);
            stripped = document.Documents.Any(p => p.IsStripped);
            return true;
        } catch (YamlException) {
            return false;
        }
    }
}

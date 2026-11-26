using OfficeIMO.PowerPoint;
using Xunit;
using static OfficeIMO.Tests.PowerPointUnchangedSaveTests;

namespace OfficeIMO.Tests;

/// <summary>
/// Replacing text in a presentation whose Open XML SDK package is not open yet edits its slides without opening it, and
/// gives the replacement count and saved package a replacement through the slide model gives.
/// </summary>
public class PowerPointReplaceWithoutSdkTests {
    public static IEnumerable<object[]> CorpusPresentations() => OpenXmlSlideTextReaderTests.CorpusPresentations();

    private static readonly (string Text, string? Replacement, PowerPointTextSearchOptions? Options)[] Searches = {
        ("e", "E", null),
        ("the", null, null),
        ("Slide", "Page", new PowerPointTextSearchOptions { MatchCase = true }),
        ("a", "[a]", new PowerPointTextSearchOptions { WholeWords = true }),
        ("o", "0", new PowerPointTextSearchOptions { IncludeTables = false })
    };

    [Theory]
    [MemberData(nameof(CorpusPresentations))]
    public void ReplacesInCorpusPresentationsAsTheModelDoes(string relativePath) {
        byte[] source = File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, relativePath));
        AssertNoDifferences(Differences(source));
    }

    [Fact]
    public void ReplacesAcrossRunsWithoutOpeningTheSdk() {
        byte[] source = Rewrite(CreateDeck(), "ppt/slides/slide1.xml", xml => xml.Replace("<a:t>First slide</a:t></a:r>",
            "<a:t>Fir</a:t></a:r><a:r><a:rPr lang=\"en-US\" b=\"1\"/><a:t>st sl</a:t></a:r><a:r><a:rPr lang=\"en-US\"/><a:t>ide</a:t></a:r>"));
        Assert.Contains("<a:t>st sl</a:t>", System.Text.Encoding.UTF8.GetString(ReadParts(source)["ppt/slides/slide1.xml"]));

        using var presentation = PowerPointDocument.Load(new MemoryStream(source));
        Assert.Equal(1, presentation.ReplaceAll("first slide", "Opening"));
        Assert.False(presentation.IsSdkLoaded);

        byte[] saved = presentation.ToBytes();
        Assert.False(presentation.IsSdkLoaded);
        AssertNoDifferences(Differences(source, ("first slide", "Opening", null)));
        using var reopened = PowerPointDocument.Load(new MemoryStream(saved));
        Assert.Contains("Opening", string.Concat(reopened.ExtractMarkdownChunks().Select(chunk => chunk.Text)));
    }

    [Fact]
    public void LaterReadsSeeTheEditsBeforeTheSave() {
        using var presentation = PowerPointDocument.Load(new MemoryStream(CreateDeck()));
        presentation.ReplaceAll("Second", "Middle");
        presentation.ReplaceAll("Middle slide", "Centre slide");

        Assert.Contains("Centre slide", string.Concat(presentation.ExtractMarkdownChunks().Select(chunk => chunk.Text)));
        Assert.False(presentation.IsSdkLoaded);
        Assert.Contains(presentation.Slides[1].TextBoxes, box => box.Text == "Centre slide");
        Assert.False(presentation.IsSdkLoaded);
    }

    [Fact]
    public void SaveOnDisposeWritesTheEdits() {
        var destination = new MemoryStream();
        byte[] source = CreateDeck();
        destination.Write(source, 0, source.Length);
        destination.Position = 0;

        var presentation = PowerPointDocument.Load(destination, new PowerPointLoadOptions { PersistenceMode = DocumentPersistenceMode.SaveOnDispose });
        presentation.ReplaceAll("Third", "Last");
        presentation.Dispose();

        Assert.False(presentation.IsSdkLoaded);
        using var reopened = PowerPointDocument.Load(new MemoryStream(destination.ToArray()));
        Assert.Contains(reopened.Slides[2].TextBoxes, box => box.Text == "Last slide");
    }

    // A save the fast path declines goes through the SDK, which opens the package with the edits in it.
    [Fact]
    public void ASaveThroughTheSdkKeepsTheEdits() {
        byte[] source = Rewrite(CreateDeck(), "ppt/presentation.xml", xml => xml.Replace("</p:presentation>",
            "<p:extLst><p:ext uri=\"urn:officeimo:powerpoint:footer-definitions:1\"/></p:extLst></p:presentation>"));
        using var presentation = PowerPointDocument.Load(new MemoryStream(source));
        presentation.ReplaceAll("First", "Opening");
        Assert.False(presentation.IsSdkLoaded);

        byte[] saved = presentation.ToBytes();

        Assert.True(presentation.IsSdkLoaded);
        using var reopened = PowerPointDocument.Load(new MemoryStream(saved));
        Assert.Contains(reopened.Slides[0].TextBoxes, box => box.Text == "Opening slide");
    }

    [Theory]
    [InlineData("line")]
    [InlineData("comment")]
    [InlineData("notes")]
    [InlineData("split text")]
    [InlineData("cdata between runs")]
    public void ReplacesWhatTheFastPathLeavesAsTheModelDoes(string variant) {
        byte[] source = CreateDeck();
        string replacement = "Opening";
        PowerPointTextSearchOptions? options = null;
        if (variant == "line") replacement = "Opening\nline";
        if (variant == "comment") source = Rewrite(source, "ppt/slides/slide1.xml", xml => xml.Replace("<a:t>First slide</a:t></a:r>", "<a:t>First slide</a:t></a:r><!-- note -->"));
        if (variant == "split text") source = Rewrite(source, "ppt/slides/slide1.xml", xml => xml.Replace("<a:t>First slide</a:t>", "<a:t>First<![CDATA[ slide]]></a:t>"));
        if (variant == "cdata between runs") source = Rewrite(source, "ppt/slides/slide1.xml", xml => xml.Replace("<a:t>First slide</a:t></a:r>", "<a:t>First slide</a:t></a:r><![CDATA[ ]]>"));
        if (variant == "notes") options = new PowerPointTextSearchOptions { IncludeNotes = true };

        using var presentation = PowerPointDocument.Load(new MemoryStream(source));
        int count = presentation.ReplaceAll("First", replacement, options);

        // Runs and paragraphs edit the package themselves; notes still go through the SDK.
        Assert.Equal(variant == "notes", presentation.IsSdkLoaded);
        Assert.Equal(1, count);
        if (variant == "split text") {
            // The SDK reads only the text before the CDATA section; the package keeps all of it.
            using var reopened = PowerPointDocument.Load(new MemoryStream(presentation.ToBytes()));
            Assert.Contains(reopened.Slides[0].TextBoxes, box => box.Text == "Opening slide");
            return;
        }
        AssertNoDifferences(Differences(source, ("First", replacement, options)));
    }

    // Opt-in, as the other corpus checks: OFFICEIMO_PPTX_CORPUS names a folder of presentations.
    [Fact]
    public void ExternalCorpusPresentationsReplaceAsTheModelDoes() {
        string? folder = Environment.GetEnvironmentVariable("OFFICEIMO_PPTX_CORPUS");
        if (string.IsNullOrEmpty(folder) || !Directory.Exists(folder)) return;
        var failures = new List<string>();
        foreach (string path in Directory.EnumerateFiles(folder, "*.pptx", SearchOption.AllDirectories).OrderBy(file => file, StringComparer.Ordinal)) {
            byte[] source = File.ReadAllBytes(path);
            try {
                using (PowerPointDocument.Load(new MemoryStream(source))) { }
            } catch (Exception) {
                continue;
            }
            failures.AddRange(Differences(source).Select(difference => path + ": " + difference));
        }
        Assert.True(failures.Count == 0, string.Join(Environment.NewLine, failures.Take(40)));
    }

    private static List<string> Differences(byte[] source, params (string Text, string? Replacement, PowerPointTextSearchOptions? Options)[] searches) {
        var differences = new List<string>();
        foreach ((string text, string? replacement, PowerPointTextSearchOptions? options) in searches.Length == 0 ? Searches : searches) {
            string label = text + " -> " + (replacement ?? "(removed)");
            int modelCount = -1, deferredCount = -1;
            byte[]? model = null, deferred = null;
            Exception? modelFailure = Record.Exception(() => {
                using var opened = PowerPointDocument.Load(new MemoryStream(source));
                _ = opened.OpenXmlDocument;
                modelCount = opened.ReplaceAll(text, replacement!, options);
                model = opened.ToBytes();
            });
            Exception? deferredFailure = Record.Exception(() => {
                using var loaded = PowerPointDocument.Load(new MemoryStream(source));
                deferredCount = loaded.ReplaceAll(text, replacement!, options);
                deferred = loaded.ToBytes();
            });
            if (modelFailure != null || deferredFailure != null) {
                if (modelFailure?.GetType() != deferredFailure?.GetType()) differences.Add(label + ": model " + modelFailure?.GetType().Name + ", without the SDK " + deferredFailure?.GetType().Name);
                continue;
            }
            if (modelCount != deferredCount) differences.Add(label + ": " + modelCount + " replaced by the model, " + deferredCount + " without the SDK");
            Exception? packageDifference = Record.Exception(() => AssertSamePackage(source, model!, deferred!, keptAsLoaded: false));
            if (packageDifference != null) differences.Add(label + ": " + packageDifference.Message);
        }
        return differences;
    }

    private static byte[] CreateDeck() {
        using var presentation = PowerPointDocument.Create(new MemoryStream());
        presentation.AddSlide().AddTextBox("First slide");
        presentation.AddSlide().AddTextBox("Second slide");
        presentation.AddSlide().AddTextBox("Third slide");
        return presentation.ToBytes();
    }

    // The whole of each difference, which the collection display would cut short.
    private static void AssertNoDifferences(List<string> differences) =>
        Assert.True(differences.Count == 0, string.Join(Environment.NewLine, differences));
}

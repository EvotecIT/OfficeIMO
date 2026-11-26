using OfficeIMO.PowerPoint;
using Xunit;
using ExtractOptions = OfficeIMO.PowerPoint.PowerPointExtractionExtensions.PowerPointExtractOptions;

namespace OfficeIMO.Tests;

/// <summary>
/// A presentation whose Open XML SDK package is not open yet counts its slides and extracts its text without opening
/// it, and gives the same count and chunks as the slide model.
/// </summary>
public class PowerPointTextWithoutSdkTests {
    public static IEnumerable<object[]> CorpusPresentations() => OpenXmlSlideTextReaderTests.CorpusPresentations();

    private static readonly (ExtractOptions Extract, int MaxChars)[] Variants = {
        (new ExtractOptions(), 8000),
        (new ExtractOptions { IncludeMarkdown = false }, 8000),
        (new ExtractOptions { IncludeNotes = false, IncludeTables = false, IncludeHiddenShapes = false }, 8000),
        (new ExtractOptions(), 256),
        (new ExtractOptions { IncludeMarkdown = false }, 256)
    };

    [Theory]
    [MemberData(nameof(CorpusPresentations))]
    public void CountsAndExtractsCorpusPresentationsAsTheModelDoes(string relativePath) {
        byte[] source = File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, relativePath));
        AssertNoDifferences(Differences(source));
    }

    [Fact]
    public void CountsAndExtractsWithoutOpeningTheSdk() {
        byte[] source = CreateDeck();
        using var presentation = PowerPointDocument.Load(new MemoryStream(source));

        Assert.Equal(3, presentation.Slides.Count);
        List<PowerPointExtractChunk> chunks = presentation.ExtractMarkdownChunks().ToList();

        Assert.False(presentation.IsSdkLoaded);
        Assert.Contains(chunks, chunk => chunk.Text.Contains("Second slide"));
        Assert.Contains(chunks, chunk => chunk.Text.Contains("Speaker note"));
        AssertNoDifferences(Differences(source));
    }

    [Fact]
    public void ReadsTheOptionsWhenTheExtractionIsRequested() {
        using var presentation = PowerPointDocument.Load(new MemoryStream(CreateDeck()));
        var extract = new ExtractOptions { IncludeMarkdown = false };
        var chunking = new PowerPointExtractChunkingOptions { MaxChars = 256 };
        IEnumerable<PowerPointExtractChunk> pending = presentation.ExtractMarkdownChunks(extract, chunking);

        extract.IncludeMarkdown = true;
        chunking.MaxChars = 100_000;
        PowerPointExtractChunk[] chunks = pending.ToArray();

        Assert.False(presentation.IsSdkLoaded);
        Assert.All(chunks, chunk => Assert.Null(chunk.Markdown));
    }

    [Fact]
    public void CancellingAfterTheFirstChunkStopsTheExtraction() {
        using var presentation = PowerPointDocument.Load(new MemoryStream(CreateDeck()));
        using var cancellation = new CancellationTokenSource();
        using IEnumerator<PowerPointExtractChunk> chunks = presentation.ExtractMarkdownChunks(cancellationToken: cancellation.Token).GetEnumerator();
        Assert.True(chunks.MoveNext());
        Assert.False(presentation.IsSdkLoaded);

        cancellation.Cancel();

        Assert.Throws<OperationCanceledException>(() => chunks.MoveNext());
    }

    [Fact]
    public void TheCountFollowsAddedSlidesBeforeAndAfterTheSdkOpens() {
        using var presentation = PowerPointDocument.Load(new MemoryStream(CreateDeck()));
        IReadOnlyList<PowerPointSlide> slides = presentation.Slides;
        Assert.Equal(3, slides.Count);

        presentation.AddSlide();
        Assert.False(presentation.IsSdkLoaded);
        Assert.Equal(4, slides.Count);
        _ = presentation.OpenXmlDocument;

        Assert.True(presentation.IsSdkLoaded);
        Assert.Equal(4, slides.Count);
        Assert.Equal(4, presentation.Slides.Count);
        Assert.Equal(4, slides.Count(slide => slide != null));
    }

    // Opt-in, as the unchanged-save corpus check: OFFICEIMO_PPTX_CORPUS names a folder of presentations.
    [Fact]
    public void ExternalCorpusPresentationsCountAndExtractAsTheModelDoes() {
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

    // Each variant extracted from a deferred document and from one whose SDK is open; also the slide counts.
    private static List<string> Differences(byte[] source) {
        var differences = new List<string>();
        foreach ((ExtractOptions extract, int maxChars) in Variants) {
            var chunking = new PowerPointExtractChunkingOptions { MaxChars = maxChars };
            Exception? modelFailure = null, streamedFailure = null;
            List<string>? model = null, streamed = null;
            int modelCount = -1, streamedCount = -1;
            try {
                using var opened = PowerPointDocument.Load(new MemoryStream(source));
                _ = opened.OpenXmlDocument;
                modelCount = opened.Slides.Count;
                model = opened.ExtractMarkdownChunks(extract, chunking, "deck.pptx").Select(Describe).ToList();
            } catch (Exception exception) {
                modelFailure = exception;
            }
            try {
                using var deferred = PowerPointDocument.Load(new MemoryStream(source));
                streamedCount = deferred.Slides.Count;
                streamed = deferred.ExtractMarkdownChunks(extract, chunking, "deck.pptx").Select(Describe).ToList();
            } catch (Exception exception) {
                streamedFailure = exception;
            }
            string variant = (extract.IncludeMarkdown ? "markdown" : "text") + (extract.IncludeNotes ? "" : ", no notes") + ", " + maxChars;
            if (modelFailure != null || streamedFailure != null) {
                if (modelFailure?.GetType() != streamedFailure?.GetType()) differences.Add(variant + ": model " + modelFailure?.GetType().Name + ", streamed " + streamedFailure?.GetType().Name);
                continue;
            }
            if (modelCount != streamedCount) differences.Add(variant + ": " + modelCount + " slides, counted " + streamedCount);
            if (!model!.SequenceEqual(streamed!)) {
                int index = Enumerable.Range(0, Math.Min(model.Count, streamed!.Count)).FirstOrDefault(i => model[i] != streamed[i]);
                differences.Add(variant + ": chunk " + index + " of " + model.Count + "/" + streamed!.Count + Environment.NewLine
                    + "  model:    " + (index < model.Count ? Shorten(model[index]) : "-") + Environment.NewLine
                    + "  streamed: " + (index < streamed.Count ? Shorten(streamed[index]) : "-"));
            }
        }
        return differences;
    }

    private static string Describe(PowerPointExtractChunk chunk) =>
        chunk.Id + "|" + chunk.Location.Path + "|" + chunk.Location.Slide + "|" + chunk.Location.BlockIndex + "|" + chunk.Text + "|" + chunk.Markdown
        + "|" + string.Join(";", chunk.Warnings ?? Array.Empty<string>());

    private static string Shorten(string text) => text.Length <= 400 ? text : text.Substring(0, 400) + "...";

    private static byte[] CreateDeck() {
        using var presentation = PowerPointDocument.Create(new MemoryStream());
        presentation.AddSlide().AddTextBox("First slide");
        PowerPointSlide second = presentation.AddSlide();
        second.AddTextBox("Second slide");
        second.Notes.Text = "Speaker note";
        presentation.AddSlide().AddTextBox("Third slide");
        return presentation.ToBytes();
    }

    // The whole of each difference, which the collection display would cut short.
    private static void AssertNoDifferences(List<string> differences) =>
        Assert.True(differences.Count == 0, string.Join(Environment.NewLine, differences));
}

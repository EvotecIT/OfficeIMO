using System.Globalization;
using System.Text;
using System.Threading;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using OfficeIMO.IWork;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.IWork.Tests;

public sealed class IWorkPagesWritingTests {
    [Fact]
    public void Creation_matches_the_independently_decoded_package_edited_and_reopened_by_Pages() {
        byte[] bytes = Example().SaveBytes();
        Assert.Equal(File.ReadAllBytes(Fixture("pages-created-v15.4.pages")), bytes);
        Assert.Equal(bytes, Example().SaveBytes());
        IWorkPagesProjection projection = IWorkSourceDocument.Open(new MemoryStream(bytes)).ReadPages();
        Assert.True(projection.HasEditableContent);
        Assert.Empty(projection.Diagnostics);
        Assert.Empty(projection.SourceDeclarationIssues);
        Assert.Empty(projection.SourceReferenceIssues);
        Assert.Equal(84, projection.Paragraphs.Count); // Native terminating LF also projects an empty trailing paragraph.
        Assert.Equal("", projection.Paragraphs[2]);
        VerifyRichRuns(projection, " plain reset");
        using var result = WordIWorkConverter.ConvertPagesToWordResult(new MemoryStream(bytes),
            conversionOptions: new() { Mode = IWorkConversionMode.EditableOnly });
        result.RequireCompleteEditableReconstruction();
        using var output = new MemoryStream(); result.Value.Save(output); output.Position = 0;
        using var saved = WordprocessingDocument.Open(output, false);
        Assert.Empty(new OpenXmlValidator().Validate(saved));
        Assert.Equal(projection.Paragraphs, saved.MainDocumentPart!.Document!.Body!.Elements<W.Paragraph>().Select(Text));
        W.Run[] runs = saved.MainDocumentPart!.Document!.Body!.Elements<W.Paragraph>().ElementAt(1).Elements<W.Run>()
            .Where(run => Text(run).Length > 0).ToArray();
        Assert.Equal(new[] { "Bold 😀", " plain reset", " — Zażółć café" }, runs.Select(Text));
        Assert.True(Enabled(runs[0].RunProperties!.Bold)); Assert.True(Enabled(runs[0].RunProperties!.Italic));
        Assert.Equal(W.UnderlineValues.Single, runs[0].RunProperties!.Underline!.Val!.Value);
        Assert.True(Enabled(runs[0].RunProperties!.Strike));
        Assert.False(Enabled(runs[1].RunProperties!.Bold)); Assert.False(Enabled(runs[1].RunProperties!.Italic));
        Assert.False(Enabled(runs[1].RunProperties!.Strike));
        Assert.Equal("24", runs[1].RunProperties!.FontSize!.Val!.Value);
        Assert.Equal("000000", runs[1].RunProperties!.Color!.Val!.Value);
    }

    [Fact]
    public void Native_saved_creation_retains_the_edit_and_matches_Apple_DOCX_content_and_page_settings() {
        string path = Fixture("pages-created-edited-v15.4.pages");
        IWorkPagesProjection projection = IWorkSourceDocument.Open(path).ReadPages();
        Assert.Empty(projection.SourceDeclarationIssues); Assert.Empty(projection.SourceReferenceIssues);
        Assert.Empty(projection.Diagnostics);
        VerifyRichRuns(projection, " plain edited");
        using var result = WordIWorkConverter.ConvertPagesToWordResult(path,
            conversionOptions: new() { Mode = IWorkConversionMode.EditableOnly });
        result.RequireCompleteEditableReconstruction();
        using var output = new MemoryStream(); result.Value.Save(output); output.Position = 0;
        using var saved = WordprocessingDocument.Open(output, false);
        using var apple = WordprocessingDocument.Open(Fixture("pages-created-v15.4.docx"), false);
        Assert.Empty(new OpenXmlValidator().Validate(saved));
        W.Body body = saved.MainDocumentPart!.Document!.Body!, reference = apple.MainDocumentPart!.Document!.Body!;
        // Apple exports the authored paragraphs; OfficeIMO also retains the source's final LF paragraph.
        Assert.Equal(reference.Elements<W.Paragraph>().Select(Text), body.Elements<W.Paragraph>().Take(83).Select(Text));
        Assert.Equal("", Text(body.Elements<W.Paragraph>().Last()));
        W.SectionProperties section = body.Elements<W.SectionProperties>().Single();
        W.SectionProperties native = reference.Elements<W.SectionProperties>().Single();
        Assert.Equal(native.GetFirstChild<W.PageSize>()!.Width!.Value, section.GetFirstChild<W.PageSize>()!.Width!.Value);
        Assert.Equal(native.GetFirstChild<W.PageSize>()!.Height!.Value, section.GetFirstChild<W.PageSize>()!.Height!.Value);
        var actualMargin = section.GetFirstChild<W.PageMargin>()!; var nativeMargin = native.GetFirstChild<W.PageMargin>()!;
        Assert.Equal(nativeMargin.Top!.Value, actualMargin.Top!.Value); Assert.Equal(nativeMargin.Bottom!.Value, actualMargin.Bottom!.Value);
        Assert.Equal(nativeMargin.Left!.Value, actualMargin.Left!.Value); Assert.Equal(nativeMargin.Right!.Value, actualMargin.Right!.Value);
    }

    [Fact]
    public void Landscape_creation_matches_the_package_opened_and_exported_by_Pages() {
        var document = IWorkPagesDocument.Create(792, 612, 36, 54, 45, 63);
        document.AddParagraph("Landscape page settings", fontSizePoints: 24);
        Assert.Equal(File.ReadAllBytes(Fixture("pages-created-landscape-v15.4.pages")), document.SaveBytes());
    }

    [Theory]
    [InlineData(612f, 792f, false)]
    [InlineData(792f, 612f, true)]
    public void Custom_page_dimensions_and_margins_survive_native_readback_and_saved_Word(float width, float height, bool landscape) {
        var document = IWorkPagesDocument.Create(width, height, 36, 54, 45, 63);
        document.AddParagraph("Page settings");
        byte[] bytes = document.SaveBytes();
        IWorkPageLayout layout = IWorkSourceDocument.Open(new MemoryStream(bytes)).ReadPages().PageLayout!;
        Assert.Equal(width, layout.WidthPoints); Assert.Equal(height, layout.HeightPoints);
        Assert.Equal(landscape, layout.IsLandscape);
        Assert.Equal(36, layout.LeftMarginPoints); Assert.Equal(54, layout.RightMarginPoints);
        Assert.Equal(45, layout.TopMarginPoints); Assert.Equal(63, layout.BottomMarginPoints);
        using var result = WordIWorkConverter.ConvertPagesToWordResult(new MemoryStream(bytes));
        result.RequireCompleteEditableReconstruction();
        using var saved = new MemoryStream(); result.Value.Save(saved); saved.Position = 0;
        using var package = WordprocessingDocument.Open(saved, false);
        var section = package.MainDocumentPart!.Document!.Body!.Elements<W.SectionProperties>().Single();
        Assert.Equal((uint)width * 20, section.GetFirstChild<W.PageSize>()!.Width!.Value);
        Assert.Equal((uint)height * 20, section.GetFirstChild<W.PageSize>()!.Height!.Value);
        Assert.Equal(landscape ? W.PageOrientationValues.Landscape : W.PageOrientationValues.Portrait,
            section.GetFirstChild<W.PageSize>()!.Orient!.Value);
    }

    [Fact]
    public void Paragraphs_snapshot_runs_and_keep_editable_blank_paragraph_metrics() {
        var document = IWorkPagesDocument.Create();
        var input = new List<IWorkTextRunData> { new("", fontSizePoints: 80), new("visible", fontSizePoints: 14), new("") };
        IWorkPagesParagraphData paragraph = document.AddRichParagraph(input); input.Clear();
        Assert.Single(paragraph.Runs); Assert.Equal("visible", paragraph.Text);
        Assert.Throws<NotSupportedException>(() => ((IList<IWorkTextRunData>)paragraph.Runs).Clear());
        document.AddParagraph("", "Times New Roman", 20);
        IWorkPagesProjection projection = IWorkSourceDocument.Open(new MemoryStream(document.SaveBytes())).ReadPages();
        Assert.Equal(new[] { "visible", "", "" }, projection.Paragraphs);
        Assert.Equal(14, projection.Body!.Paragraphs[0].Runs[0].Style!.FontSizePoints);
        Assert.Equal(20, projection.Body.Paragraphs[1].Style.TextStyle.FontSizePoints);
        Assert.Equal("Times New Roman", projection.Body.Paragraphs[1].Style.TextStyle.FontName);
    }

    [Theory]
    [InlineData("paragraphs")]
    [InlineData("runs")]
    [InlineData("characters")]
    [InlineData("bytes")]
    public void Encoding_budgets_reject_before_mutating_a_file_or_caller_stream(string budget) {
        var document = IWorkPagesDocument.Create();
        document.AddRichParagraph(new[] { new IWorkTextRunData("A😀", fontSizePoints: 12), new IWorkTextRunData("B", fontSizePoints: 12) });
        document.AddParagraph("");
        var limits = new IWorkPagesWriteOptions();
        switch (budget) {
            case "paragraphs": limits.MaximumParagraphs = 1; break;
            case "runs": limits.MaximumTextRuns = 2; break;
            case "characters": limits.MaximumTextCharacters = 5; break; // 4 UTF-16 units plus both terminators require 6.
            case "bytes": limits.MaximumPackageBytes = 128; break;
        }
        using var stream = new MemoryStream(new byte[] { 9, 8, 7 });
        Assert.Throws<InvalidDataException>(() => document.Save(stream, limits)); Assert.Equal(new byte[] { 9, 8, 7 }, stream.ToArray());
        string folder = TemporaryFolder();
        try {
            string path = Path.Combine(folder, "existing.pages"); File.WriteAllBytes(path, new byte[] { 9, 8, 7 });
            Assert.Throws<InvalidDataException>(() => document.Save(path, overwrite: true, options: limits));
            Assert.Equal(new byte[] { 9, 8, 7 }, File.ReadAllBytes(path)); Assert.Single(Directory.GetFiles(folder));
        } finally { Directory.Delete(folder, recursive: true); }
        if (budget == "characters") {
            limits.MaximumTextCharacters = 6;
            Assert.Equal("A😀B\n\n", IWorkSourceDocument.Open(new MemoryStream(document.SaveBytes(limits))).ReadPages().Body!.PlainText);
        }
    }

    [Fact]
    public void Saves_obey_stream_position_conflict_policy_cancellation_and_caller_ownership() {
        var document = IWorkPagesDocument.Create(); document.AddParagraph("save"); byte[] bytes = document.SaveBytes();
        using var stream = new MemoryStream(); stream.WriteByte(9); document.Save(stream);
        Assert.True(stream.CanWrite); Assert.Equal(bytes, stream.ToArray().Skip(1));
        using var cancelled = new CancellationTokenSource(); cancelled.Cancel();
        byte[] before = stream.ToArray();
        Assert.Throws<OperationCanceledException>(() => document.Save(stream, cancellationToken: cancelled.Token));
        Assert.Equal(before, stream.ToArray());
        string folder = TemporaryFolder();
        try {
            string path = Path.Combine(folder, "created.pages"); document.Save(path);
            Assert.Throws<IOException>(() => document.Save(path)); Assert.Equal(bytes, File.ReadAllBytes(path));
            Assert.Throws<OperationCanceledException>(() => document.Save(path, overwrite: true, cancellationToken: cancelled.Token));
            Assert.Equal(bytes, File.ReadAllBytes(path)); Assert.Single(Directory.GetFiles(folder));
        } finally { Directory.Delete(folder, recursive: true); }
        var large = IWorkPagesDocument.Create(); large.AddParagraph(new string('a', 150_000));
        using var copying = new CancellationTokenSource(); using var destination = new CancelAfterWriteStream(copying);
        Assert.Throws<OperationCanceledException>(() => large.Save(destination, cancellationToken: copying.Token));
        Assert.Equal(65_536, destination.Length); Assert.True(destination.CanWrite);
    }

    [Fact]
    public void Invalid_text_paragraphs_page_settings_and_options_fail_without_model_mutation() {
        var document = IWorkPagesDocument.Create();
        Assert.Throws<InvalidOperationException>(() => document.SaveBytes());
        Assert.Throws<ArgumentNullException>(() => document.AddRichParagraph(null!));
        Assert.Throws<ArgumentException>(() => document.AddRichParagraph(Array.Empty<IWorkTextRunData>()));
        Assert.Throws<ArgumentException>(() => document.AddRichParagraph(new IWorkTextRunData[] { null! }));
        Assert.Throws<ArgumentException>(() => document.AddParagraph("two\r\nparagraphs"));
        Assert.Throws<ArgumentException>(() => document.AddParagraph("tab\there"));
        Assert.Throws<EncoderFallbackException>(() => document.AddParagraph("invalid\ud800"));
        Assert.Empty(document.Paragraphs);
        Assert.Throws<ArgumentOutOfRangeException>(() => IWorkPagesDocument.Create(float.NaN));
        Assert.Throws<ArgumentOutOfRangeException>(() => IWorkPagesDocument.Create(leftMarginPoints: -1));
        Assert.Throws<ArgumentException>(() => IWorkPagesDocument.Create(leftMarginPoints: 400, rightMarginPoints: 300));
        document.AddParagraph("valid");
        Assert.Throws<ArgumentOutOfRangeException>(() => document.SaveBytes(new() { MaximumParagraphs = 0 }));
        Assert.Throws<ArgumentOutOfRangeException>(() => document.SaveBytes(new() { MaximumTextRuns = 0 }));
        Assert.Throws<ArgumentOutOfRangeException>(() => document.SaveBytes(new() { MaximumTextCharacters = 0 }));
        Assert.Throws<ArgumentOutOfRangeException>(() => document.SaveBytes(new() { MaximumPackageBytes = 0 }));
    }

    private static void VerifyRichRuns(IWorkPagesProjection projection, string plain) {
        IWorkTextRun[] runs = projection.Body!.Paragraphs[1].Runs.ToArray();
        Assert.Equal(new[] { "Bold 😀", plain, " — Zażółć café" }, runs.Select(run => run.Text));
        Assert.True(runs[0].Style!.Bold); Assert.True(runs[0].Style!.Italic);
        Assert.True(runs[0].Style!.Underline); Assert.True(runs[0].Style!.Strikethrough);
        Assert.Equal(16, runs[0].Style!.FontSizePoints); Assert.Equal("003366", runs[0].Style!.Color!.RgbHex);
        Assert.False(runs[1].Style!.Bold); Assert.False(runs[1].Style!.Italic);
        Assert.False(runs[1].Style!.Underline); Assert.False(runs[1].Style!.Strikethrough);
        Assert.Equal(12, runs[1].Style!.FontSizePoints); Assert.Equal("Arial", runs[1].Style!.FontName);
        Assert.Equal("000000", runs[1].Style!.Color!.RgbHex);
        Assert.True(runs[2].Style!.Italic); Assert.Equal("Times New Roman", runs[2].Style!.FontName);
        Assert.Equal(14, runs[2].Style!.FontSizePoints); Assert.Equal("008000", runs[2].Style!.Color!.RgbHex);
    }

    internal static IWorkPagesDocument Example() {
        var document = IWorkPagesDocument.Create();
        document.AddParagraph("OfficeIMO native body document", fontSizePoints: 24);
        document.AddRichParagraph(new[] {
            new IWorkTextRunData("Bold 😀", "Arial", 16, "003366", bold: true, italic: true, underline: true, strikethrough: true),
            new IWorkTextRunData(" plain reset", "Arial", 12), new IWorkTextRunData(" — Zażółć café", "Times New Roman", 14, "008000", italic: true)
        });
        document.AddParagraph("");
        for (int index = 1; index <= 80; index++)
            document.AddParagraph("Paragraph " + index.ToString("00", CultureInfo.InvariantCulture) + ": the body flows over pages with editable text.");
        return document;
    }

    private static bool Enabled(W.OnOffType? value) => value is not null && (value.Val?.Value ?? true);
    private static string Text(DocumentFormat.OpenXml.OpenXmlElement element) => string.Concat(element.Descendants<W.Text>().Select(text => text.Text));
    private static string Fixture(string name) => Path.Combine(AppContext.BaseDirectory, "Documents", "IWorkCorpus", "native-exports", name);
    private static string TemporaryFolder() {
        string path = Path.Combine(Path.GetTempPath(), "OfficeIMO-PagesWriting-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(path); return path;
    }
    private sealed class CancelAfterWriteStream(CancellationTokenSource cancellation) : MemoryStream {
        public override void Write(byte[] buffer, int offset, int count) { base.Write(buffer, offset, count); cancellation.Cancel(); }
    }
}

using System.Text.Json;
using OfficeIMO.Web.Converter.Engine;
using Xunit;

namespace OfficeIMO.Web.Converter.Tests;

/// <summary>Runs the engine's tool handlers in-process, the same way engine-worker.js calls them in the browser.</summary>
public sealed class EngineToolTests {
    [Theory]
    [InlineData("BT /F1 18 Tf 20 80 Td (Hello) Tj ET")]
    [InlineData("q UnknownPaint Q")]
    public void IncompletePdfComparisonDoesNotClaimIdenticalFilesDiffer(string content) {
        byte[] pdf = System.Text.Encoding.ASCII.GetBytes(string.Join("\n", new[] {
            "%PDF-1.4", "1 0 obj", "<< /Type /Catalog /Pages 2 0 R >>", "endobj",
            "2 0 obj", "<< /Type /Pages /Count 1 /Kids [3 0 R] >>", "endobj",
            "3 0 obj", "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 240 180] /Resources << /Font << /F1 5 0 R >> >> /Contents 4 0 R >>", "endobj",
            "4 0 obj", "<< /Length " + content.Length + " >>", "stream", content, "endstream", "endobj",
            "5 0 obj", "<< /Type /Font /Subtype /Type1 /BaseFont /CustomSans /Encoding /WinAnsiEncoding >>", "endobj",
            "trailer", "<< /Root 1 0 R /Size 6 >>", "%%EOF", ""
        }));
        var session = new ToolSession();
        session.Stage(0, pdf, "expected.pdf");
        session.Stage(1, pdf, "actual.pdf");

        ToolResultDocument result = PdfTool.Run(session, "compare", "run", Options());

        Assert.True(result.Ok);
        Assert.Equal("The comparison is incomplete", result.Verdict.Title);
        Assert.Contains(result.Facts, fact => fact.Label == "Incomplete pages" && fact.Value == "1");
        Assert.Contains("1 incomplete page", System.Text.Encoding.UTF8.GetString(session.Artifact(0)), StringComparison.Ordinal);
    }

    private static byte[] Sample(string name) => File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "samples", name));

    private static ToolOptions Options(params (string Key, string Value)[] values) =>
        new(values.ToDictionary(static pair => pair.Key, static pair => pair.Value, StringComparer.Ordinal));

    [Theory]
    [InlineData(300_000)]
    [InlineData(1_048_573)]
    public void TextFileReviewScansTheEntireAdvertisedByteAllowance(int prefixCharacters) {
        var session = new ToolSession();
        byte[] input = System.Text.Encoding.UTF8.GetBytes(new string('a', prefixCharacters) + "\u200b");
        session.Stage(0, input, "large-text.txt");

        ToolResultDocument result = TextTool.Run(session, "inspect", Options(("source", "file")));

        Assert.True(result.Ok);
        Assert.Single(result.Items);
        Assert.Equal(prefixCharacters + 1, result.Preview!.Text!.Length);
        Assert.EndsWith("\u200b", result.Preview.Text, StringComparison.Ordinal);
        Assert.Contains(result.Facts, fact => fact.Label == "Characters" && fact.Value ==
            (prefixCharacters + 1).ToString("N0", System.Globalization.CultureInfo.InvariantCulture));
    }

    [Fact]
    public void TextFileReviewAndCleanCopyPreserveNameEncodingAndPreamble() {
        var session = new ToolSession();
        byte[] input = System.Text.Encoding.Unicode.GetPreamble().Concat(System.Text.Encoding.Unicode.GetBytes("Visible\u200b text")).ToArray();
        session.Stage(0, input, "UTF16-original.txt");
        var review = TextTool.Run(session, "inspect", Options(("source", "file")));
        Assert.Contains(review.Facts, fact => fact.Label == "Encoding" && fact.Value == "UTF-16");
        Assert.Contains(review.Facts, fact => fact.Label == "File" && fact.Value == "UTF16-original.txt");
        var removed = TextTool.Run(session, "remove", Options(("source", "file"), ("remove", review.Items[0].Id)));
        var artifact = Assert.Single(removed.Artifacts, item => item.Role == "primary");
        Assert.Equal("UTF16-original-cleaned.txt", artifact.FileName);
        Assert.Equal("text/plain;charset=utf-16", artifact.ContentType);
        byte[] actual = session.Artifact(artifact.Index);
        Assert.Equal(System.Text.Encoding.Unicode.GetPreamble(), actual.Take(2));
        Assert.Equal("Visible text", System.Text.Encoding.Unicode.GetString(actual, 2, actual.Length - 2));
    }

    [Fact]
    public void WordToPdf_ReturnsAPreviewableFileAndAPlainVerdict() {
        var session = new ToolSession();
        session.Stage(0, Sample("business-summary.docx"), "summary.docx");
        ToolResultDocument result = ConvertTool.Run(session, "docx-pdf", "convert", Options(("profile", "faithful")));

        Assert.True(result.Ok);
        Assert.StartsWith("PDF ready", result.Verdict.Title, StringComparison.Ordinal);
        Assert.Equal("good", result.Verdict.Tone);
        Assert.DoesNotContain("Faithful", result.Verdict.Detail, StringComparison.Ordinal);
        ToolArtifact primary = Assert.Single(result.Artifacts, static artifact => artifact.Role == "primary");
        Assert.Equal("application/pdf", primary.ContentType);
        Assert.Contains(result.Artifacts, static artifact => artifact.Role == "report");
        Assert.Equal("pdf", result.Preview!.Kind);
        Assert.Contains(result.Facts, static fact => fact.Label == "PDF type" && fact.Value == "Standard");
        Assert.Equal((byte)'%', session.Artifact(primary.Index)[0]);
    }

    [Fact]
    public void MarkdownToHtml_ReturnsASandboxableHtmlPreviewWithSource() {
        var session = new ToolSession();
        ToolResultDocument result = ConvertTool.Run(session, "markdown-html", "convert", Options(("text", "# Title\n\n- one\n- two")));
        Assert.True(result.Ok);
        Assert.Equal("html", result.Preview!.Kind);
        Assert.Contains("<h1", result.Preview.Html!, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void ConversionWithoutInputExplainsWhatIsMissing() {
        var session = new ToolSession();
        Assert.Throws<InvalidOperationException>(() => ConvertTool.Run(session, "docx-pdf", "convert", Options()));
        Assert.Throws<NotSupportedException>(() => ConvertTool.Run(session, "no-such-route", "convert", Options()));
    }

    [Fact]
    public void FileOrigin_UnsignedCredentialDoesNotEstablishCertificateIdentity() {
        var session = new ToolSession();
        session.Stage(0, Sample("unsigned-credential-12-actions.png"), "unsigned.png");
        ToolResultDocument result = OriginTool.Run(session, "inspect", Options());
        Assert.Contains(result.Facts, fact => fact.Label == "Certificate subject" && fact.Value.Contains("unverified", StringComparison.OrdinalIgnoreCase));
        Assert.Contains("unverified", result.Verdict.Detail, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("signed by", result.Verdict.Detail, StringComparison.OrdinalIgnoreCase);
        Assert.Contains(result.Items, item => item.Id == OriginTool.Manifests && item.Detail.Contains("not verified", StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public void FileOrigin_InspectNamesTheRecordAndRemoveReportsWhatChanged() {
        var session = new ToolSession();
        session.Stage(0, Sample("provenance-demo.png"), "demo.png");

        ToolResultDocument inspect = OriginTool.Run(session, "inspect", Options());
        Assert.True(inspect.Ok);
        Assert.Equal("Found 1 origin record", inspect.Verdict.Title);
        ToolItem label = Assert.Single(inspect.Items, static item => item.Selectable);
        Assert.Equal(OriginTool.Declarations, label.Id);
        Assert.Contains("generative AI", label.Detail, StringComparison.Ordinal);
        Assert.DoesNotContain("iTXt", label.Detail, StringComparison.Ordinal);

        ToolResultDocument removed = OriginTool.Run(session, "remove", Options(("remove", OriginTool.Declarations)));
        Assert.True(removed.Ok);
        Assert.Equal("good", removed.Verdict.Tone);
        Assert.StartsWith("Clean copy ready", removed.Verdict.Title, StringComparison.Ordinal);
        Assert.Contains(removed.Items, static item => item.Id == OriginTool.Declarations && item.State == ToolState.Removed);
        Assert.Contains(removed.Items, static item => item.Id == "content" && item.State == ToolState.Kept);
        Assert.Equal("image", removed.Preview!.Kind);
    }

    [Fact]
    public void FileOrigin_KeepingARecordIsReportedAsTheVisitorsChoice() {
        var session = new ToolSession();
        session.Stage(0, Sample("provenance-demo.png"), "demo.png");
        Assert.Throws<InvalidOperationException>(() => OriginTool.Run(session, "remove", Options()));
        ToolResultDocument result = OriginTool.Run(session, "remove", Options(("remove", OriginTool.Manifests)));
        Assert.Contains(result.Items, static item => item.Id == OriginTool.Declarations && item.State == ToolState.Kept);
        Assert.Equal("Nothing was removed", result.Verdict.Title);
    }

    [Fact]
    public void HiddenCharacters_FindsDirectionOverridesWithOffsetsAndRemovesOnlySelected() {
        var session = new ToolSession();
        string text = "Invoice‮123‬ and a‍b";
        ToolResultDocument inspect = TextTool.Run(session, "inspect", Options(("text", text)));
        Assert.True(inspect.Ok);
        Assert.Equal("warn", inspect.Verdict.Tone);
        Assert.All(inspect.Items, static item => Assert.False(item.Selected));
        ToolItem first = inspect.Items[0];
        Assert.Equal(7, first.Start);
        Assert.Equal("Can change meaning", first.Group);

        Assert.Throws<InvalidOperationException>(() => TextTool.Run(session, "remove", Options(("remove", first.Id), ("text", "Edited text"))));
        ToolResultDocument removed = TextTool.Run(session, "remove", Options(("remove", first.Id), ("text", text)));
        Assert.True(removed.Ok);
        Assert.Equal(ToolState.Removed, removed.Items[0].State);
        Assert.Equal(ToolState.Kept, removed.Items[^1].State);
        // The preview keeps the original text so markers stay aligned; the download is the cleaned copy.
        Assert.Equal(text, removed.Preview!.Text);
        string cleaned = System.Text.Encoding.UTF8.GetString(session.Artifact(removed.Artifacts.Single(static artifact => artifact.Role == "primary").Index));
        Assert.DoesNotContain('‮', cleaned);
        Assert.Contains('‍', cleaned);
    }

    [Fact]
    public void InspectPdf_ReturnsAFactSheet() {
        var session = new ToolSession();
        session.Stage(0, Sample("showcase-dashboard.pdf"), "showcase.pdf");
        ToolResultDocument result = PdfTool.Run(session, "inspect", "run", Options());
        Assert.True(result.Ok);
        Assert.Contains(result.Facts, static fact => fact.Label == "Pages");
        Assert.Contains(result.Facts, static fact => fact.Label == "Password" && fact.Value == "No");
        Assert.Contains("PDF", result.Verdict.Title, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("Critical blockers", "Critical blockers")]
    [InlineData("  Critical blockers  ", "  Critical blockers  ")]
    [InlineData("\tCritical blockers\n", "Critical blockers")]
    public void RedactFindsMatchesBeforeRemovingAnything(string findText, string executeText) {
        var session = new ToolSession();
        session.Stage(0, Sample("showcase-dashboard.pdf"), "showcase.pdf");
        ToolResultDocument found = PdfTool.Run(session, "redact", "find", Options(("text", findText)));
        Assert.True(found.Ok);
        Assert.Empty(found.Artifacts);
        Assert.NotEmpty(found.Items);
        Assert.All(found.Items, static item => Assert.NotNull(item.Page));

        ToolResultDocument none = PdfTool.Run(session, "redact", "find", Options(("text", "zzz-not-in-this-file")));
        Assert.StartsWith("No matches", none.Verdict.Title, StringComparison.Ordinal);

        Assert.Throws<InvalidOperationException>(() => PdfTool.Run(session, "redact", "run", Options(("text", "Critical blockers"), ("confirm", "true"))));
        PdfTool.Run(session, "redact", "find", Options(("text", findText)));
        Assert.Throws<InvalidOperationException>(() => PdfTool.Run(session, "redact", "run", Options(("text", "different phrase"), ("confirm", "true"))));
        session.Stage(0, Sample("showcase-dashboard.pdf"), "replacement.pdf");
        Assert.Throws<InvalidOperationException>(() => PdfTool.Run(session, "redact", "run", Options(("text", executeText), ("confirm", "true"))));
        PdfTool.Run(session, "redact", "find", Options(("text", findText)));

        ToolResultDocument applied = PdfTool.Run(session, "redact", "run", Options(("text", executeText), ("confirm", "true")));
        Assert.True(applied.Ok);
        Assert.Equal("Text removed and checked", applied.Verdict.Title);
    }

    [Fact]
    public void ClearingInputsReleasesThePreviousResultsAndReviews() {
        var session = new ToolSession();
        TextTool.Run(session, "inspect", Options(("text", "a\u202Eb")));
        Assert.NotNull(session.LastReview);
        Assert.NotEmpty(session.Artifact(0));
        session.ClearInputs();
        Assert.Empty(session.Inputs);
        Assert.Null(session.LastReview);
        Assert.Null(session.LastConversion);
        Assert.Throws<ArgumentOutOfRangeException>(() => session.Artifact(0));
    }

    [Fact]
    public void ProbeReportsPagesAndProtection() {
        var session = new ToolSession();
        session.Stage(0, Sample("showcase-dashboard.pdf"), "showcase.pdf");
        PdfProbeDocument probe = PdfTool.Probe(session, 0);
        Assert.True(probe.Ok);
        Assert.True(probe.PageCount > 0);
        Assert.False(probe.NeedsPassword);
    }

    [Fact]
    public void PageThumbnailsRenderFromStagedInput() {
        var session = new ToolSession();
        session.Stage(0, Sample("showcase-dashboard.pdf"), "showcase.pdf");
        var image = session.Preview("input", 0).Render(1, 220);
        Assert.True(image.Succeeded);
        Assert.NotNull(image.Bytes);
        Assert.True(image.Bytes!.Length > 0);
    }

    [Fact]
    public void ResultDocumentsSerializeWithoutNullNoise() {
        string json = JsonSerializer.Serialize(
            new ToolResultDocument(true, new ToolVerdict("good", "Done", "Detail"), [], [new ToolItem("a", "Title", "Detail", ToolState.Info)], [], null, 5),
            EngineJsonContext.Default.ToolResultDocument);
        Assert.Contains("\"verdict\":{\"tone\":\"good\"", json, StringComparison.Ordinal);
        Assert.DoesNotContain("\"page\":null", json, StringComparison.Ordinal);
        Assert.DoesNotContain("\"preview\"", json, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("RightToLeftOverride", "Right to left override")]
    [InlineData("PDFVersion", "PDF version")]
    public void IdentifiersBecomeReadableWords(string identifier, string expected) =>
        Assert.Equal(expected, ToolFormat.Words(identifier));

    [Fact]
    public void PageListsReadNaturally() {
        Assert.Equal("page 3", ToolFormat.Pages([3]));
        Assert.Equal("pages 2 and 5", ToolFormat.Pages([5, 2, 5]));
        Assert.Equal("pages 1, 2 and 3", ToolFormat.Pages([3, 1, 2]));
    }

    [Theory]
    [InlineData("1e2")]
    [InlineData("1.5")]
    [InlineData("")]
    public void ExplicitMalformedNumericOptionsNeverSilentlyUseDefaults(string value) {
        Assert.Throws<ArgumentException>(() => Options(("perFile", value)).Number("perFile", 1));
        Assert.Equal(1, Options().Number("perFile", 1));
        Assert.Equal(100, Options(("perFile", "100")).Number("perFile", 1));
    }
}

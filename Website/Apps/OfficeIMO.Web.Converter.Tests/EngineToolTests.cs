using System.Text.Json;
using OfficeIMO.Web.Converter.Engine;
using Xunit;

namespace OfficeIMO.Web.Converter.Tests;

/// <summary>Runs the engine's tool handlers in-process, the same way engine-worker.js calls them in the browser.</summary>
public sealed class EngineToolTests {
    private static byte[] Sample(string name) => File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "samples", name));

    private static ToolOptions Options(params (string Key, string Value)[] values) =>
        new(values.ToDictionary(static pair => pair.Key, static pair => pair.Value, StringComparer.Ordinal));

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

        ToolResultDocument removed = TextTool.Run(session, "remove", Options(("remove", first.Id)));
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

    [Fact]
    public void RedactFindsMatchesBeforeRemovingAnything() {
        var session = new ToolSession();
        session.Stage(0, Sample("showcase-dashboard.pdf"), "showcase.pdf");
        ToolResultDocument found = PdfTool.Run(session, "redact", "find", Options(("text", "Critical blockers")));
        Assert.True(found.Ok);
        Assert.Empty(found.Artifacts);
        Assert.NotEmpty(found.Items);
        Assert.All(found.Items, static item => Assert.NotNull(item.Page));

        ToolResultDocument none = PdfTool.Run(session, "redact", "find", Options(("text", "zzz-not-in-this-file")));
        Assert.StartsWith("No matches", none.Verdict.Title, StringComparison.Ordinal);

        ToolResultDocument applied = PdfTool.Run(session, "redact", "run", Options(("text", "Critical blockers"), ("confirm", "true")));
        Assert.True(applied.Ok);
        Assert.Equal("Text removed and checked", applied.Verdict.Title);
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
}

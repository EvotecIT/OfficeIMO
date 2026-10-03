using System.IO.Compression;
using OfficeIMO.Mhtml;
using OfficeIMO.Pdf;
using OfficeIMO.Tool.Commands.Html;
using Xunit;

namespace OfficeIMO.Tool.Tests;

public sealed class HtmlPdfIntentCommandTests {
    private const string Styles = "@page{size:384px 288px;margin:0}html,body{margin:0}"
        + "p{height:280px;margin:0;font:16px/20px Arial}"
        + "@media print{.screen{display:none}}@media screen{.print{display:none}}";
    private const string Body = "<section class='print'><p>Print first</p><p>Print second</p></section>"
        + "<section class='screen'><p>Screen first</p><p>Screen second</p><p>Screen third</p></section>";

    [Theory]
    [InlineData("html", "print-paged", "Print", 2)]
    [InlineData("mhtml", "print-paged", "Print", 2)]
    [InlineData("zip", "print-paged", "Print", 2)]
    [InlineData("html", "screen-media-paged", "Screen", 3)]
    [InlineData("mhtml", "screen-media-paged", "Screen", 3)]
    [InlineData("zip", "screen-media-paged", "Screen", 3)]
    [InlineData("html", "screen-snapshot-paged", "Screen", 1)]
    [InlineData("mhtml", "screen-snapshot-paged", "Screen", 1)]
    [InlineData("zip", "screen-snapshot-paged", "Screen", 1)]
    public async Task PdfIntentPreservesSelectedMediaAndArchivedStyles(
        string format, string profile, string visibleMarker, int expectedPages) {
        await using var input = new MemoryStream(CreateInput(format));
        await using var output = new MemoryStream();
        using var errors = new StringWriter();

        int exit = await HtmlCommand.RunAsync(new[] {
            "convert", "-", "--input-format", format, "--output", "-", "--profile", profile,
            "--viewport-width", "384", "--viewport-height", "288"
        }, input, output, errors);

        Assert.True(exit == 0, errors.ToString());
        PdfReadDocument pdf = PdfReadDocument.Open(output.ToArray());
        Assert.Equal(expectedPages, pdf.Pages.Count);
        string text = pdf.ExtractText();
        Assert.Contains(visibleMarker + " first", text, StringComparison.Ordinal);
        Assert.Contains(visibleMarker + " second", text, StringComparison.Ordinal);
        Assert.DoesNotContain(visibleMarker == "Print" ? "Screen" : "Print", text, StringComparison.Ordinal);
        Assert.DoesNotContain("StylesheetResourceUnavailable", errors.ToString(), StringComparison.Ordinal);
    }

    [Fact]
    public async Task PdfPageSelectionUsesTheRequestedScreenSnapshot() {
        await using var input = new MemoryStream(CreateInput("mhtml", tallSnapshot: true));
        await using var output = new MemoryStream();
        using var errors = new StringWriter();

        int exit = await HtmlCommand.RunAsync(new[] {
            "convert", "-", "--input-format", "mhtml", "--output", "-",
            "--profile", "screen-snapshot-paged", "--pages", "2", "--viewport-width", "384"
        }, input, output, errors);

        Assert.True(exit == 0, errors.ToString());
        PdfReadDocument pdf = PdfReadDocument.Open(output.ToArray());
        Assert.Single(pdf.Pages);
        Assert.Contains("Screen second", pdf.ExtractText(), StringComparison.Ordinal);
        Assert.DoesNotContain("Screen first", pdf.ExtractText(), StringComparison.Ordinal);
        Assert.DoesNotContain("Screen third", pdf.ExtractText(), StringComparison.Ordinal);
    }

    private static byte[] CreateInput(string format, bool tallSnapshot = false) {
        string styles = Styles + (tallSnapshot ? ".screen p{height:1600px}" : string.Empty);
        if (format == "html") return Encoding.UTF8.GetBytes("<style>" + styles + "</style>" + Body);
        if (format == "mhtml") return new MhtmlDocument(
            "<link rel='stylesheet' href='cid:styles'>" + Body,
            new[] { new MhtmlResource(Encoding.UTF8.GetBytes(styles), "text/css", contentId: "styles") }).ToBytes();
        using var output = new MemoryStream();
        using (var zip = new ZipArchive(output, ZipArchiveMode.Create, leaveOpen: true)) {
            using (var writer = new StreamWriter(zip.CreateEntry("index.html").Open(), new UTF8Encoding(false)))
                writer.Write("<link rel='stylesheet' href='style.css'>" + Body);
            using (var writer = new StreamWriter(zip.CreateEntry("style.css").Open(), new UTF8Encoding(false)))
                writer.Write(styles);
        }
        return output.ToArray();
    }
}

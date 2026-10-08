using System.Globalization;
using System.IO.Compression;
using OfficeIMO.Pdf;
using OfficeIMO.Tool.Commands.Html;
using Xunit;

namespace OfficeIMO.Tool.Tests;

public sealed class HtmlSiteBundleCommandTests {
    [Theory]
    [InlineData("convert")]
    [InlineData("render")]
    public async Task BundledInputPreservesCssAndSelectedRootWithoutConfusingEncodedAndDecodedLimits(string command) {
        byte[] bytes = CreateBundle();
        using var input = new MemoryStream(bytes);
        using var output = new MemoryStream();
        using var error = new StringWriter();
        var args = new List<string> {
            command, "-", "--input-format", "site-bundle", "--output", "-",
            "--entry-path", "pages/article.html", "--base-uri", "https://docs.example.test/archive/",
            "--max-input-bytes", bytes.Length.ToString(CultureInfo.InvariantCulture)
        };
        if (command == "render") args.AddRange(new[] { "--encoder", "svg", "--profile", "print-paged" });

        int exit = await HtmlCommand.RunAsync(args.ToArray(), input, output, error);

        Assert.True(exit == 0, error.ToString());
        Assert.True(input.CanRead);
        Assert.DoesNotContain("StylesheetResourceUnavailable", error.ToString(), StringComparison.Ordinal);
        if (command == "convert") {
            string text = PdfReadDocument.Open(output.ToArray()).ExtractText();
            Assert.Contains("Embedded CLI bundle", text, StringComparison.Ordinal);
            Assert.DoesNotContain("Wrong root", text, StringComparison.Ordinal);
        } else {
            using var archive = new ZipArchive(new MemoryStream(output.ToArray()), ZipArchiveMode.Read);
            using var reader = new StreamReader(archive.GetEntry("pages/page-0001.svg")!.Open());
            string svg = await reader.ReadToEndAsync();
            Assert.Contains("CLI bundle", svg, StringComparison.Ordinal);
            Assert.Contains("#123456", svg, StringComparison.Ordinal);
            Assert.DoesNotContain("Wrong root", svg, StringComparison.Ordinal);
        }
    }

    [Fact]
    public async Task ZipFileSelectsBundleLoaderAndKeepsInputArtifact() {
        string path = Path.Combine(Path.GetTempPath(), "officeimo-site-" + Guid.NewGuid().ToString("N") + ".zip");
        byte[] bytes = CreateBundle();
        await File.WriteAllBytesAsync(path, bytes);
        try {
            using var input = new MemoryStream();
            using var output = new MemoryStream();
            using var error = new StringWriter();
            int exit = await HtmlCommand.RunAsync(new[] {
                "convert", path, "--entry-path", "pages/article.html", "--output", "-"
            }, input, output, error);
            Assert.True(exit == 0, error.ToString());
            Assert.Contains("Embedded CLI bundle", PdfReadDocument.Open(output.ToArray()).ExtractText(), StringComparison.Ordinal);
            Assert.Equal(bytes, await File.ReadAllBytesAsync(path));
        } finally {
            File.Delete(path);
        }
    }

    [Theory]
    [InlineData("--max-bundle-entry-bytes")]
    [InlineData("--max-bundle-decoded-bytes")]
    [InlineData("--max-bundle-entries")]
    public async Task DecodedBundleLimitsRejectBeforePublishingOutput(string option) {
        using var input = new MemoryStream(CreateBundle());
        using var output = new MemoryStream();
        using var error = new StringWriter();
        int exit = await HtmlCommand.RunAsync(new[] {
            "render", "-", "--input-format", "zip", "--entry-path", "pages/article.html",
            "--output", "-", option, "1"
        }, input, output, error);
        Assert.Equal((int)OfficeImoToolExitCode.UnsupportedInput, exit);
        Assert.Equal(0, output.Length);
        Assert.Contains("limit", error.ToString(), StringComparison.OrdinalIgnoreCase);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task AmbiguousOrUnsafeBundlesRejectBeforePublishingOutput(bool unsafePath) {
        byte[] bytes = unsafePath ? CreateArchive(("../index.html", "<p>Unsafe</p>")) : CreateBundle();
        using var input = new MemoryStream(bytes);
        using var output = new MemoryStream();
        using var error = new StringWriter();
        int exit = await HtmlCommand.RunAsync(new[] {
            "convert", "-", "--input-format", "zip", "--output", "-"
        }, input, output, error);
        Assert.Equal((int)OfficeImoToolExitCode.UnsupportedInput, exit);
        Assert.Equal(0, output.Length);
    }

    [Theory]
    [InlineData("html", "--entry-path", "index.html")]
    [InlineData("zip", "--base-uri", "file:///tmp/")]
    public async Task InapplicableBundleConfigurationFailsBeforeReadingInput(string format, string option, string value) {
        using var input = new MemoryStream(CreateBundle());
        using var output = new MemoryStream();
        using var error = new StringWriter();
        int exit = await HtmlCommand.RunAsync(new[] {
            "convert", "-", "--input-format", format, "--output", "-", option, value
        }, input, output, error);
        Assert.Equal((int)OfficeImoToolExitCode.Usage, exit);
        Assert.Equal(0, input.Position);
        Assert.Equal(0, output.Length);
    }

    private static byte[] CreateBundle() => CreateArchive(
        ("pages/article.html", "<link rel='stylesheet' href='../styles/site.css'>"
            + "<p class='marker'>CLI bundle</p><p>" + new string('x', 4000) + "</p>"),
        ("pages/other.html", "<p>Wrong root</p>"),
        ("styles/site.css", "p.marker::before{content:'Embedded ';}p.marker{color:#123456}"));

    private static byte[] CreateArchive(params (string Path, string Content)[] entries) {
        using var stream = new MemoryStream();
        using (var archive = new ZipArchive(stream, ZipArchiveMode.Create, leaveOpen: true)) {
            foreach (var item in entries) {
                using var writer = new StreamWriter(archive.CreateEntry(item.Path).Open(), new UTF8Encoding(false));
                writer.Write(item.Content);
            }
        }
        return stream.ToArray();
    }
}

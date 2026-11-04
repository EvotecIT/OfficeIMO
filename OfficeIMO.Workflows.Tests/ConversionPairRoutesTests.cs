using System.IO.Compression;
using System.Text;
using OfficeIMO.AsciiDoc;
using OfficeIMO.AsciiDoc.Markdown;
using OfficeIMO.Excel;
using OfficeIMO.Html;
using OfficeIMO.OneNote;
using OfficeIMO.OneNote.Html;
using OfficeIMO.OpenDocument;
using OfficeIMO.Pdf;
using OfficeIMO.PowerPoint;
using OfficeIMO.Rtf;
using OfficeIMO.Word;
using OfficeIMO.Word.OpenDocument;

namespace OfficeIMO.Workflows.Tests;

public sealed class ConversionPairRoutesTests {
    [Theory]
    [InlineData("asciidoc-markdown", ".adoc", ".md", "= Marker\n\nVisible *bold*.\n")]
    [InlineData("markdown-asciidoc", ".md", ".adoc", "# Marker\n\nVisible **bold**.\n")]
    [InlineData("latex-markdown", ".tex", ".md", "\\documentclass{article}\n\\begin{document}\n\\section{Marker}\nVisible.\n\\end{document}\n")]
    [InlineData("markdown-latex", ".md", ".tex", "# Marker\n\nVisible **bold**.\n")]
    [InlineData("rtf-markdown", ".rtf", ".md", "{\\rtf1\\ansi Marker\\par Visible}")]
    [InlineData("rtf-html", ".rtf", ".html", "{\\rtf1\\ansi Marker\\par Visible}")]
    public async Task TextPairRoutesPreserveVisibleContentInThePublishedArtifact(string route, string sourceExtension,
        string targetExtension, string text) {
        var (bytes, report) = await ConvertFileAsync(route, sourceExtension, Encoding.UTF8.GetBytes(text), targetExtension);
        Assert.Contains("Marker", Encoding.UTF8.GetString(bytes), StringComparison.Ordinal);
        Assert.DoesNotContain(report.Diagnostics, finding => finding.LossKind == OfficeConversionLossKind.Failure);
        Assert.NotNull(report.OutputPath);
    }

    [Theory]
    [InlineData("html-rtf", ".html", "<h1>Marker</h1><p>Visible <b>bold</b>.</p>")]
    [InlineData("markdown-rtf", ".md", "# Marker\n\nVisible **bold**.\n")]
    public async Task RtfTargetsRetainNativeOutputOptions(string route, string extension, string text) {
        var (bytes, _) = await ConvertFileAsync(route, extension, Encoding.UTF8.GetBytes(text), ".rtf",
            new OfficeWorkflowRouteSettings { RtfSave = new RtfSaveOptions { IncludeGenerator = false } });
        using var stream = new MemoryStream(bytes);
        using RtfDocument document = RtfDocument.Load(stream);
        Assert.Contains(document.Paragraphs, paragraph => paragraph.ToPlainText().Contains("Marker", StringComparison.Ordinal));
        Assert.DoesNotContain("\\generator", Encoding.ASCII.GetString(bytes), StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("docx-odt", "odt-docx", ".docx", ".odt")]
    [InlineData("xlsx-ods", "ods-xlsx", ".xlsx", ".ods")]
    [InlineData("pptx-odp", "odp-pptx", ".pptx", ".odp")]
    public async Task OpenDocumentPairsWriteRealNativePackagesInBothDirections(string forward, string reverse,
        string sourceExtension, string targetExtension) {
        byte[] original = CreateOfficeDocument(sourceExtension);
        var (odf, _) = await ConvertFileAsync(forward, sourceExtension, original, targetExtension,
            new OfficeWorkflowRouteSettings { OpenDocumentSave = new OdfSaveOptions { CompatibilityProfile = OdfCompatibilityProfile.Odf13 } });
        string content = PackageXml(odf);
        Assert.Contains("Marker", content, StringComparison.Ordinal);
        Assert.Contains("version=\"1.3\"", content, StringComparison.Ordinal);
        var (office, _) = await ConvertFileAsync(reverse, targetExtension, odf, sourceExtension);
        Assert.Contains("Marker", PackageXml(office), StringComparison.Ordinal);
    }

    [Fact]
    public async Task ExcelHtmlRoutesPreserveEncryptedWorkbookContentInBothDirections() {
        using var source = new MemoryStream();
        using (ExcelDocument workbook = ExcelDocument.Create()) {
            workbook.AddSheet("Secure").Cell(1, 1).SetValue("Encrypted marker");
            workbook.Save(source, new ExcelSaveOptions { Password = "test-only-password" });
        }
        source.Position = 0;
        using var html = new MemoryStream();
        OfficeConversionReport exported = await OfficeConverter.ConvertAsync(source, OfficeDocumentFormat.Excel,
            html, OfficeDocumentFormat.Html, new OfficeConversionOptions { SourcePassword = "test-only-password" });
        Assert.Contains("Encrypted marker", Encoding.UTF8.GetString(html.ToArray()), StringComparison.Ordinal);
        Assert.DoesNotContain(exported.Diagnostics, finding => finding.LossKind == OfficeConversionLossKind.Failure);
        html.Position = 0;
        using var restored = new MemoryStream();
        await OfficeConverter.ConvertAsync(html, OfficeDocumentFormat.Html, restored, OfficeDocumentFormat.Excel);
        restored.Position = 0;
        using ExcelDocument reopened = ExcelDocument.Load(restored);
        Assert.Equal("Secure", reopened.Sheets[0].Name);
        Assert.True(reopened.Sheets[0].TryGetCellText(1, 1, out string? value));
        Assert.Equal("Encrypted marker", value);
        Assert.True(source.CanRead);
        Assert.True(html.CanRead);
        Assert.True(restored.CanWrite);
    }

    [Fact]
    public async Task WordRtfPairUsesNativeWordSaveAndRtfReadContracts() {
        var (rtf, _) = await ConvertFileAsync("docx-rtf", ".docx", CreateOfficeDocument(".docx"), ".rtf");
        var (word, _) = await ConvertFileAsync("rtf-docx", ".rtf", rtf, ".docx");
        using var stream = new MemoryStream(word);
        using WordDocument document = WordDocument.Load(stream);
        Assert.Contains(document.Paragraphs, paragraph => paragraph.Text.Contains("Marker", StringComparison.Ordinal));
    }

    [Fact]
    public async Task OneNoteRoutesPreserveSectionOptionsAndSemanticProjection() {
        var (one, _) = await ConvertFileAsync("html-onenote", ".html", Encoding.UTF8.GetBytes("<h1>Marker</h1><p>Visible notes.</p>"), ".one",
            new OfficeWorkflowRouteSettings { HtmlOneNote = new HtmlOneNoteExportOptions { SectionName = "Imported section" } });
        using var input = new MemoryStream(one);
        using OneNoteSection section = OneNoteSection.Load(input);
        Assert.Equal("Imported section", section.Name);
        foreach ((string route, string extension) in new[] { ("onenote-html", ".html"), ("onenote-markdown", ".md"), ("onenote-pdf", ".pdf") }) {
            var (output, _) = await ConvertFileAsync(route, ".one", one, extension);
            if (extension == ".pdf") {
                using PdfDocument pdf = PdfDocument.Load(output);
                Assert.Contains("Marker", pdf.ExtractText(), StringComparison.Ordinal);
            } else Assert.Contains("Marker", Encoding.UTF8.GetString(output), StringComparison.Ordinal);
        }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task NotebookProjectionIncludesNestedSectionsFromLocalTocAndPortablePackage(bool package) {
        string directory = Directory.CreateTempSubdirectory("officeimo-notebook-route-").FullName;
        try {
            OneNoteNotebook notebook = CreateNotebook();
            string source;
            if (package) {
                source = Path.Combine(directory, "notebook.onepkg");
                OneNotePackageWriter.Write(notebook, source);
            } else {
                string folder = Path.Combine(directory, "Notebook");
                OneNoteNotebookWriter.Write(notebook, folder);
                source = Path.Combine(folder, "Open Notebook.onetoc2");
            }
            string output = Path.Combine(directory, "target.md");
            OfficeConversionReport report = await OfficeConverter.ConvertAsync(source, output,
                new OfficeConversionOptions { RouteId = "onenote-markdown" });
            string text = await File.ReadAllTextAsync(output);
            Assert.Contains("Root marker", text, StringComparison.Ordinal);
            Assert.Contains("Nested marker", text, StringComparison.Ordinal);
            Assert.DoesNotContain(report.Diagnostics, diagnostic => diagnostic.Code == "ONENOTE_TOC_STREAM_METADATA_ONLY");
        } finally { Directory.Delete(directory, recursive: true); }
    }

    [Fact]
    public async Task NotebookSectionChangeBeforePublicationKeepsTheExistingDestination() {
        string directory = Directory.CreateTempSubdirectory("officeimo-notebook-change-").FullName;
        try {
            string folder = Path.Combine(directory, "Notebook");
            OneNoteNotebookWriter.Write(CreateNotebook(), folder);
            string source = Path.Combine(folder, "Open Notebook.onetoc2");
            string section = Directory.EnumerateFiles(folder, "*.one").Single();
            string output = Path.Combine(directory, "target.md");
            await File.WriteAllTextAsync(output, "Keep existing destination");
            OfficeConversionException error = await Assert.ThrowsAsync<OfficeConversionException>(() => OfficeConverter.ConvertAsync(source, output,
                new OfficeConversionOptions { RouteId = "onenote-markdown", FileConflictPolicy = OfficePublicationPolicy.Replace,
                    PublicationGuard = new ChangingSectionGuard(section) }));
            Assert.True(error.Report.HasLoss);
            Assert.Equal("Keep existing destination", await File.ReadAllTextAsync(output));
        } finally { Directory.Delete(directory, recursive: true); }
    }

    [Fact]
    public async Task StrictPairLossRejectsBeforeWritingTheCallerStreamAndRetainsSourceDiagnostics() {
        using var source = new MemoryStream(Encoding.UTF8.GetBytes("---\n"));
        using var output = new MemoryStream();
        OfficeConversionException error = await Assert.ThrowsAsync<OfficeConversionException>(() =>
            OfficeConverter.ConvertAsync(source, OfficeDocumentFormat.Markdown, output, OfficeDocumentFormat.AsciiDoc,
                new OfficeConversionOptions { FailOnLoss = true, RegisteredSettings = new OfficeWorkflowRouteSettings {
                    MarkdownAsciiDoc = new MarkdownAsciiDocExportOptions { PreserveUnsupportedAsSource = false }
                } }));
        Assert.True(error.Report.HasLoss);
        Assert.Contains(error.Report.Diagnostics, diagnostic => diagnostic.LossKind == OfficeConversionLossKind.Omission);
        Assert.Equal(0, output.Length);
        Assert.True(source.CanRead);
        Assert.True(output.CanWrite);
    }

    [Theory]
    [InlineData("docx-rtf", ".docx", ".rtf")]
    [InlineData("docx-odt", ".docx", ".odt")]
    [InlineData("xlsx-ods", ".xlsx", ".ods")]
    [InlineData("xlsx-html", ".xlsx", ".html")]
    [InlineData("pptx-odp", ".pptx", ".odp")]
    public async Task OfficePairRoutesRejectXmlBeyondTheConfiguredLimitBeforeWriting(string route, string sourceExtension,
        string targetExtension) {
        byte[] original = CreateOfficeDocument(sourceExtension);
        using var source = new MemoryStream(original);
        using var output = new MemoryStream();
        OfficePackageSecurityException error = await Assert.ThrowsAsync<OfficePackageSecurityException>(() =>
            OfficeConverter.ConvertAsync(source, OfficeConverter.FormatForExtension(sourceExtension), output,
                OfficeConverter.FormatForExtension(targetExtension), new OfficeConversionOptions {
                    RouteId = route, SourceName = "source" + sourceExtension, TargetName = "output" + targetExtension,
                    MaximumXmlCharactersInPart = 1
                }));
        Assert.Equal(OfficePackageSecurityRule.MalformedPackage, error.Rule);
        Assert.Equal(0, output.Length);
        Assert.Equal(original, source.ToArray());
    }

    [Fact]
    public void RouteSnapshotsPreserveLossThresholdsWithoutSharingMutableLoadLimits() {
        var original = new OfficeWorkflowRouteSettings {
            WordOpenDocument = new WordOpenDocumentConversionOptions { LossPolicy = OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported },
            RtfLoad = new RtfLoadOptions(),
            AsciiDocLoad = new AsciiDocParseOptions { MaxTableCellCount = 11 }
        };
        OfficeWorkflowRouteSettings odf = original.ForRoute("docx-odt");
        Assert.Equal(OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported, odf.WordOpenDocument!.LossPolicy);
        original.WordOpenDocument.LossPolicy = OdfConversionLossPolicy.ReportOnly;
        Assert.Equal(OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported, odf.WordOpenDocument.LossPolicy);
        OfficeWorkflowRouteSettings rtf = original.ForRoute("rtf-markdown");
        original.RtfLoad.Limits.MaxDepth = 3;
        Assert.NotEqual(3, rtf.RtfLoad!.Limits.MaxDepth);
        Assert.Null(rtf.WordOpenDocument);
        OfficeWorkflowRouteSettings asciiDoc = original.ForRoute("asciidoc-markdown");
        original.AsciiDocLoad.MaxTableCellCount = 22;
        Assert.Equal(11, asciiDoc.AsciiDocLoad!.MaxTableCellCount);
    }

    private static async Task<(byte[] Bytes, OfficeConversionReport Report)> ConvertFileAsync(string route, string sourceExtension,
        byte[] source, string targetExtension, OfficeWorkflowRouteSettings? settings = null) {
        string directory = Directory.CreateTempSubdirectory("officeimo-pair-route-").FullName;
        try {
            string input = Path.Combine(directory, "source" + sourceExtension), output = Path.Combine(directory, "target" + targetExtension);
            await File.WriteAllBytesAsync(input, source);
            OfficeConversionReport report = await OfficeConverter.ConvertAsync(input, output, new OfficeConversionOptions {
                RouteId = route, RegisteredSettings = settings
            });
            Assert.Equal(source, await File.ReadAllBytesAsync(input));
            return (await File.ReadAllBytesAsync(output), report);
        } finally { Directory.Delete(directory, recursive: true); }
    }

    private static byte[] CreateOfficeDocument(string extension) {
        using var output = new MemoryStream();
        switch (extension) {
            case ".docx":
                using (WordDocument document = WordDocument.Create()) { document.AddParagraph("Marker"); document.Save(output); }
                break;
            case ".xlsx":
                using (ExcelDocument document = ExcelDocument.Create()) { document.AddSheet("Data").Cell(1, 1).SetValue("Marker"); document.Save(output); }
                break;
            case ".pptx":
                using (PowerPointDocument document = PowerPointDocument.Create()) { document.AddSlide().AddTextBox("Marker"); document.Save(output); }
                break;
        }
        return output.ToArray();
    }

    private static OneNoteNotebook CreateNotebook() {
        var notebook = new OneNoteNotebook { Name = "Notebook" };
        var section = OneNoteSection.Create();
        section.Name = "Root";
        section.Pages.Add(new OneNotePage { Title = "Root marker" });
        notebook.Sections.Add(section);
        var group = new OneNoteSectionGroup { Name = "Group" };
        var nested = OneNoteSection.Create();
        nested.Name = "Nested";
        nested.Pages.Add(new OneNotePage { Title = "Nested marker" });
        group.Sections.Add(nested);
        notebook.SectionGroups.Add(group);
        return notebook;
    }

    private sealed class ChangingSectionGuard(string path) : IOfficePublicationGuard {
        public async Task<bool> CanPublishAsync(string destination, bool isDirectory, CancellationToken token) {
            await File.AppendAllTextAsync(path, "changed", token);
            return true;
        }
    }

    private static string PackageXml(byte[] bytes) {
        using var stream = new MemoryStream(bytes);
        using var package = new ZipArchive(stream, ZipArchiveMode.Read);
        return string.Join("\n", package.Entries.Where(entry => entry.FullName.EndsWith(".xml", StringComparison.Ordinal))
            .Select(entry => { using var input = entry.Open(); using var reader = new StreamReader(input); return reader.ReadToEnd(); }));
    }
}

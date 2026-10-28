using OfficeIMO.AsciiDoc;
using OfficeIMO.AsciiDoc.Pdf;
using OfficeIMO.Excel;
using OfficeIMO.Excel.Pdf;
using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using OfficeIMO.Latex;
using OfficeIMO.Latex.Pdf;
using OfficeIMO.Markdown.Pdf;
using OfficeIMO.Mhtml;
using OfficeIMO.OneNote;
using OfficeIMO.OneNote.Pdf;
using OfficeIMO.OpenDocument;
using OfficeIMO.OpenDocument.Odp.Pdf;
using OfficeIMO.OpenDocument.Ods.Pdf;
using OfficeIMO.OpenDocument.Odt.Pdf;
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.Pdf;
using OfficeIMO.Rtf;
using OfficeIMO.Rtf.Pdf;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using OfficeIMO.Visio;
using OfficeIMO.Visio.Pdf;
using PdfCore = OfficeIMO.Pdf;
using OneNoteSection = global::OfficeIMO.OneNote.OneNoteSection;
using VisioDocument = global::OfficeIMO.Visio.VisioDocument;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfAdapterSecurityComplianceTests {
    public static IEnumerable<object[]> AdvertisedRoutes() =>
        new[] {
            "docx", "xlsx", "pptx", "html", "markdown", "rtf", "asciidoc", "latex",
            "mhtml", "onenote", "odt", "ods", "odp", "visio"
        }
            .Select(route => new object[] { route });

    public static IEnumerable<object[]> AutomaticFallbackRoutes() =>
        new[] { "docx", "xlsx", "pptx", "markdown" }
            .Select(route => new object[] { route });

    [Theory]
    [MemberData(nameof(AutomaticFallbackRoutes))]
    public void AutomaticFallbacksKeepWinAnsiArtifactsCompact(string route) {
        string marker = "COMPACT-" + route.ToUpperInvariant() + "-PROOF";

        byte[] pdf = ConvertWithAutomaticFallbacks(route, marker);

        Assert.InRange(pdf.Length, 500, 96 * 1024);
        Assert.Contains(marker, PdfCore.PdfReadDocument.Open(pdf).ExtractText(), StringComparison.Ordinal);
    }

    [Fact]
    public void FormatConversionDefaultsCompressContentButRespectAnExplicitOptOut() {
        const string Markdown = "# CONTENT-COMPRESSION-PROOF\n\nAdapter proof";

        byte[] compressed = (OfficeIMO.Markdown.MarkdownDocument.Parse(Markdown).ToPdf()).ToBytes();
        byte[] uncompressed = (OfficeIMO.Markdown.MarkdownDocument.Parse(Markdown).ToPdf(
            new MarkdownPdfExportOptions {
                PdfOptions = new PdfCore.PdfCreateOptions { CompressContentStreams = false },
                TextFallbacks = PdfCore.PdfTextFallbackFeatures.None
            })).ToBytes();

        Assert.True(new PdfCore.PdfCreateOptions().CompressContentStreams);
        AssertPageContentCompression(compressed, compressed: true);
        AssertPageContentCompression(uncompressed, compressed: false);
        Assert.Contains("CONTENT-COMPRESSION-PROOF", PdfCore.PdfReadDocument.Open(compressed).ExtractText(), StringComparison.Ordinal);
        Assert.Contains("CONTENT-COMPRESSION-PROOF", PdfCore.PdfReadDocument.Open(uncompressed).ExtractText(), StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("html")]
    [InlineData("mhtml")]
    public void HtmlRoutesKeepCompressionDefaultsWithCallerOptionsAndRespectOptOut(string route) {
        byte[] compressed = Convert(route, new PdfCore.PdfCreateOptions { Language = "en-US" });
        byte[] uncompressed = Convert(route, new PdfCore.PdfCreateOptions { CompressContentStreams = false });

        AssertPageContentCompression(compressed, compressed: true);
        AssertPageContentCompression(uncompressed, compressed: false);
        Assert.Contains(Marker(route), PdfCore.PdfReadDocument.Open(compressed).ExtractText(), StringComparison.Ordinal);
        Assert.Contains(Marker(route), PdfCore.PdfReadDocument.Open(uncompressed).ExtractText(), StringComparison.Ordinal);
    }

    [Theory]
    [MemberData(nameof(AutomaticFallbackRoutes))]
    public void AutomaticFallbacksRemainEnabledForUnicodeText(string route) {
        string marker = "UNICODE-" + route.ToUpperInvariant() + "-ŁÓDŹ";

        byte[] pdf = ConvertWithAutomaticFallbacks(route, marker);

        Assert.Contains(marker, PdfCore.PdfReadDocument.Open(pdf).ExtractText(), StringComparison.Ordinal);
    }

    [Theory]
    [MemberData(nameof(AdvertisedRoutes))]
    public void AdvertisedFormatAdaptersProducePasswordProtectedAes256Artifacts(string route) {
        PdfCore.PdfCreateOptions pdfOptions = new PdfCore.PdfCreateOptions()
            .SetEncryption("open", "owner");

        byte[] pdf = Convert(route, pdfOptions);
        PdfCore.PdfDocumentProbe probe = PdfCore.PdfInspector.Probe(pdf);

        Assert.True(probe.HasEncryption);
        Assert.Equal(6, probe.Security.EncryptionRevision);
        Assert.Equal(256, probe.Security.EncryptionLengthBits);
        Assert.Throws<PdfCore.PdfPasswordRequiredException>(() => PdfCore.PdfReadDocument.Open(pdf));
        Assert.Throws<PdfCore.PdfInvalidPasswordException>(() =>
            PdfCore.PdfReadDocument.Open(pdf, new PdfCore.PdfLoadOptions { Password = "wrong" }));
        Assert.Contains(
            Marker(route),
            PdfCore.PdfReadDocument.Open(pdf, new PdfCore.PdfLoadOptions { Password = "open" }).ExtractText(),
            StringComparison.Ordinal);
    }

    [Theory]
    [MemberData(nameof(AdvertisedRoutes))]
    public void AdvertisedFormatAdaptersPreservePdfA2BGenerationRequirements(string route) {
        byte[] pdf = Convert(route, CreatePdfA2BOptions());
        PdfCore.PdfDocumentInfo info = PdfCore.PdfInspector.Inspect(pdf);
        string raw = Encoding.ASCII.GetString(pdf);

        Assert.Equal("1.7", info.HeaderVersion);
        Assert.True(info.HasXmpMetadata);
        Assert.True(info.HasOutputIntents);
        Assert.False(PdfCore.PdfInspector.Probe(pdf).HasEncryption);
        Assert.Contains("pdfaid:part>2<", raw, StringComparison.Ordinal);
        Assert.Contains("pdfaid:conformance>B<", raw, StringComparison.Ordinal);
        Assert.Matches(@"/FontFile[23]\s", raw);
        Assert.Contains(Marker(route), PdfCore.PdfReadDocument.Open(pdf).ExtractText(), StringComparison.Ordinal);
    }

    [Fact]
    public void PdfA2BAndPasswordEncryptionRemainAnExplicitlyRejectedCombination() {
        PdfCore.PdfCreateOptions options = CreatePdfA2BOptions().SetEncryption("open", "owner");
        PdfCore.PdfDocument document = PdfCore.PdfDocument.Create(options)
            .Paragraph(paragraph => paragraph.Text(Marker("pdfa-encryption")));

        ArgumentException exception = Assert.Throws<ArgumentException>(() => document.ToBytes());

        Assert.Contains("PDF/A", exception.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("encryption", exception.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Theory]
    [InlineData(PdfCore.PdfStandardEncryptionAlgorithm.Aes128)]
    [InlineData(PdfCore.PdfStandardEncryptionAlgorithm.Aes256)]
    public void GeneratedPdfA2BArtifactCanBeProtectedAndUnlockedAfterCreation(PdfCore.PdfStandardEncryptionAlgorithm algorithm) {
        byte[] archivalPdf = Convert("html", CreatePdfA2BOptions());
        var encryption = new PdfCore.PdfStandardEncryptionOptions("open") {
            OwnerPassword = "owner",
            Algorithm = algorithm,
            AesCryptographyProvider = OfficeIMO.Security.OfficeManagedAesCryptographyProvider.Default
        };

        PdfCore.PdfSecurityMutationResult protectedPdf = PdfCore.PdfDocument.Load(new System.IO.MemoryStream(archivalPdf)).Security.Encrypt(encryption);
        PdfCore.PdfSecurityMutationResult unlockedPdf = protectedPdf.ToDocument().Security.Decrypt("owner");

        Assert.True(PdfCore.PdfInspector.Probe(protectedPdf.Pdf).HasEncryption);
        Assert.False(PdfCore.PdfInspector.Probe(unlockedPdf.Pdf).HasEncryption);
        Assert.Contains(
            Marker("html"),
            PdfCore.PdfReadDocument.Open(unlockedPdf.Pdf).ExtractText(),
            StringComparison.Ordinal);
    }

    private static void AssertPageContentCompression(byte[] bytes, bool compressed) {
        // Font streams retain their own compression policy when content-stream compression is disabled.
        var document = PdfCore.PdfReadDocument.Open(bytes);
        foreach (var page in document.Pages) {
            PdfCore.PdfObject contents = PdfCore.PdfObjectLookup.Resolve(document.Objects,
                page.PageDictionary.Items["Contents"])!;
            var streams = contents is PdfCore.PdfArray array
                ? array.Items.Select(item => Assert.IsType<PdfCore.PdfStream>(PdfCore.PdfObjectLookup.Resolve(document.Objects, item))).ToArray()
                : new[] { Assert.IsType<PdfCore.PdfStream>(contents) };
            Assert.NotEmpty(streams);
            Assert.All(streams, stream => Assert.Equal(compressed,
                stream.Dictionary.Get<PdfCore.PdfName>("Filter")?.Name == "FlateDecode"));
        }
    }

    private static PdfCore.PdfCreateOptions CreatePdfA2BOptions() {
        string? fontPath = PdfComplianceTestFonts.FindBundledOpenTypeCffFont();
        Assert.NotNull(fontPath);
        byte[] fontData = File.ReadAllBytes(fontPath!);
        return new PdfCore.PdfCreateOptions {
                FileVersion = PdfCore.PdfFileVersion.Pdf17,
                IncludeStandardFontToUnicodeMaps = true
            }
            .ConfigurePdfAGroundwork(PdfCore.PdfComplianceProfile.PdfA2B, "en-US")
            .RequireCompliance(PdfCore.PdfComplianceProfile.PdfA2B)
            .EmbedStandardFont(PdfCore.PdfStandardFont.Helvetica, fontData, "OfficeIMO Source Serif")
            .EmbedStandardFont(PdfCore.PdfStandardFont.HelveticaBold, fontData, "OfficeIMO Source Serif");
    }

    private static byte[] Convert(string route, PdfCore.PdfCreateOptions pdfOptions) => route switch {
        "docx" => ConvertDocx(pdfOptions),
        "xlsx" => ConvertXlsx(pdfOptions),
        "pptx" => ConvertPptx(pdfOptions),
        "html" => (HtmlConversionDocument.Parse("<h1>" + Marker(route) + "</h1><p>Adapter proof</p>").ToPdf(new HtmlToPdfOptions {
            PdfOptions = pdfOptions,
            TextFallbacks = PdfCore.PdfTextFallbackFeatures.None
        })).ToBytes(),
        "markdown" => (OfficeIMO.Markdown.MarkdownDocument.Parse("# " + Marker(route) + "\n\nAdapter proof").ToPdf(new MarkdownPdfExportOptions {
            PdfOptions = pdfOptions,
            TextFallbacks = PdfCore.PdfTextFallbackFeatures.None
        })).ToBytes(),
        "rtf" => (CreateRtf(route).ToPdf(new RtfToPdfOptions {
            PdfOptions = pdfOptions
        })).ToBytes(),
        "asciidoc" => AsciiDocDocument.ParseResult("= " + Marker(route) + "\n\nAdapter proof").Document.ToPdfBytes(
            new AsciiDocToPdfOptions {
                MarkdownOptions = CreateMarkdownOptions(pdfOptions)
            }),
        "latex" => LatexDocument.ParseResult(
                "\\documentclass{article}\\begin{document}\\section{" + Marker(route) + "}Adapter proof\\end{document}")
            .Document.ToPdfBytes(new LatexToPdfOptions {
                MarkdownOptions = CreateMarkdownOptions(pdfOptions)
            }),
        "mhtml" => new MhtmlDocument(
                "<h1>" + Marker(route) + "</h1><p>Adapter proof</p>",
                contentLocation: "https://proof.officeimo.test/report.html")
            .ToPdfBytes(CreateHtmlOptions(pdfOptions)),
        "onenote" => CreateOneNote(route).ToPdfBytes(new OneNoteToPdfOptions {
            MarkdownOptions = CreateMarkdownOptions(pdfOptions)
        }),
        "odt" => CreateOdt(route).ToPdfBytes(pdfOptions: new WordPdfExportOptions {
            PdfOptions = pdfOptions,
            TextFallbacks = PdfCore.PdfTextFallbackFeatures.None
        }),
        "ods" => CreateOds(route).ToPdfBytes(pdfOptions: new ExcelPdfExportOptions {
            PdfOptions = pdfOptions,
            TextFallbacks = PdfCore.PdfTextFallbackFeatures.None
        }),
        "odp" => CreateOdp(route).ToPdfBytes(pdfOptions: new PowerPointPdfExportOptions {
            PdfOptions = pdfOptions,
            TextFallbacks = PdfCore.PdfTextFallbackFeatures.None
        }),
        "visio" => CreateVisio(route).ToPdfBytes(new VisioToPdfOptions {
            ProjectionOptions = new PdfCore.PdfProjectionOptions { PdfOptions = pdfOptions }
        }),
        _ => throw new ArgumentOutOfRangeException(nameof(route))
    };

    private static HtmlToPdfOptions CreateHtmlOptions(PdfCore.PdfCreateOptions pdfOptions) => new HtmlToPdfOptions {
        PdfOptions = pdfOptions,
        TextFallbacks = PdfCore.PdfTextFallbackFeatures.None
    };

    private static MarkdownPdfExportOptions CreateMarkdownOptions(PdfCore.PdfCreateOptions pdfOptions) =>
        new MarkdownPdfExportOptions {
            PdfOptions = pdfOptions,
            TextFallbacks = PdfCore.PdfTextFallbackFeatures.None
        };

    private static byte[] ConvertWithAutomaticFallbacks(string route, string text) => route switch {
        "docx" => ConvertAutomaticDocx(text),
        "xlsx" => ConvertAutomaticXlsx(text),
        "pptx" => ConvertAutomaticPptx(text),
        "markdown" => (OfficeIMO.Markdown.MarkdownDocument.Parse("# " + text + "\n\nAdapter proof").ToPdf()).ToBytes(),
        _ => throw new ArgumentOutOfRangeException(nameof(route))
    };

    private static byte[] ConvertAutomaticDocx(string text) {
        using WordDocument document = WordDocument.Create();
        document.AddParagraph(text);
        document.AddParagraph("Adapter proof");
        return (document.ToPdf()).ToBytes();
    }

    private static byte[] ConvertAutomaticXlsx(string text) {
        using ExcelDocument document = ExcelDocument.Create();
        ExcelSheet sheet = document.AddWorksheet("Proof");
        sheet.Cell(1, 1, text);
        sheet.Cell(2, 1, "Adapter proof");
        return (document.ToPdf()).ToBytes(options:new global::OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true });
    }

    private static byte[] ConvertAutomaticPptx(string text) {
        using PowerPointDocument presentation = PowerPointDocument.Create();
        presentation.AddSlide().AddTextBox(
            text + "\nAdapter proof",left: global::OfficeIMO.OfficeLength.Points(36),top: global::OfficeIMO.OfficeLength.Points(36),width: global::OfficeIMO.OfficeLength.Points(420),height: global::OfficeIMO.OfficeLength.Points(100));
        return (presentation.ToPdf()).ToBytes(options:new global::OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true });
    }

    private static byte[] ConvertDocx(PdfCore.PdfCreateOptions pdfOptions) {
        using WordDocument document = WordDocument.Create();
        document.AddParagraph(Marker("docx"));
        document.AddParagraph("Adapter proof");
        return (document.ToPdf(new WordPdfExportOptions {
            PdfOptions = pdfOptions,
            TextFallbacks = PdfCore.PdfTextFallbackFeatures.None
        })).ToBytes();
    }

    private static byte[] ConvertXlsx(PdfCore.PdfCreateOptions pdfOptions) {
        using ExcelDocument document = ExcelDocument.Create();
        ExcelSheet sheet = document.AddWorksheet("Proof");
        sheet.Cell(1, 1, Marker("xlsx"));
        sheet.Cell(2, 1, "Adapter proof");
        return (document.ToPdf(new ExcelPdfExportOptions {
            PdfOptions = pdfOptions,
            TextFallbacks = PdfCore.PdfTextFallbackFeatures.None
        })).ToBytes(options:new global::OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true });
    }

    private static byte[] ConvertPptx(PdfCore.PdfCreateOptions pdfOptions) {
        using PowerPointDocument presentation = PowerPointDocument.Create();
        presentation.AddSlide().AddTextBox(
            Marker("pptx") + "\nAdapter proof",left: global::OfficeIMO.OfficeLength.Points(36),top: global::OfficeIMO.OfficeLength.Points(36),width: global::OfficeIMO.OfficeLength.Points(420),height: global::OfficeIMO.OfficeLength.Points(100));
        return (presentation.ToPdf(new PowerPointPdfExportOptions {
            PdfOptions = pdfOptions,
            TextFallbacks = PdfCore.PdfTextFallbackFeatures.None
        })).ToBytes(options:new global::OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true });
    }

    private static RtfDocument CreateRtf(string route) {
        RtfDocument document = RtfDocument.Create();
        document.AddParagraph(Marker(route));
        document.AddParagraph("Adapter proof");
        return document;
    }

    private static OneNoteSection CreateOneNote(string route) {
        var section = new OneNoteSection { Name = Marker(route) };
        var page = new OneNotePage { Title = Marker(route) };
        var paragraph = new OneNoteParagraph();
        paragraph.Runs.Add(new OneNoteTextRun { Text = "Adapter proof" });
        page.DirectContent.Add(paragraph);
        section.Pages.Add(page);
        return section;
    }

    private static OdtDocument CreateOdt(string route) {
        OdtDocument document = OdtDocument.Create();
        document.AddHeading(Marker(route), 1);
        document.AddParagraph("Adapter proof");
        return document;
    }

    private static OdsDocument CreateOds(string route) {
        OdsDocument document = OdsDocument.Create();
        OdsSheet sheet = document.AddSheet("Proof");
        sheet.Cell(0, 0).SetString(Marker(route));
        sheet.Cell(1, 0).SetString("Adapter proof");
        return document;
    }

    private static OdpPresentation CreateOdp(string route) {
        OdpPresentation presentation = OdpPresentation.Create();
        presentation.AddSlide("Proof").AddTextBox(
            OdfRect.FromCentimeters(1, 1, 20, 5),
            Marker(route) + "\nAdapter proof");
        return presentation;
    }

    private static VisioDocument CreateVisio(string route) {
        VisioDocument document = VisioDocument.Create();
        document.AddPage("Proof", 8.5, 11).Shapes.Add(
            new VisioShape("proof", 4.25, 5.5, 6, 1, Marker(route) + " Adapter proof"));
        return document;
    }

    private static string Marker(string route) => "ADAPTER-" + route.ToUpperInvariant() + "-PROOF";
}

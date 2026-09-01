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

        byte[] compressed = OfficeIMO.Markdown.MarkdownReader.Parse(Markdown).ToPdf();
        byte[] uncompressed = OfficeIMO.Markdown.MarkdownReader.Parse(Markdown).ToPdf(
            new MarkdownPdfSaveOptions {
                PdfOptions = new PdfCore.PdfOptions { CompressContentStreams = false },
                TextFallbacks = PdfCore.PdfTextFallbackFeatures.None
            });

        Assert.False(new PdfCore.PdfOptions().CompressContentStreams);
        Assert.Contains("/Filter /FlateDecode", Encoding.ASCII.GetString(compressed), StringComparison.Ordinal);
        Assert.DoesNotContain("/Filter /FlateDecode", Encoding.ASCII.GetString(uncompressed), StringComparison.Ordinal);
        Assert.Contains("CONTENT-COMPRESSION-PROOF", PdfCore.PdfReadDocument.Open(compressed).ExtractText(), StringComparison.Ordinal);
        Assert.Contains("CONTENT-COMPRESSION-PROOF", PdfCore.PdfReadDocument.Open(uncompressed).ExtractText(), StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("html")]
    [InlineData("mhtml")]
    public void HtmlRoutesKeepCompressionDefaultsWithCallerOptionsAndRespectOptOut(string route) {
        byte[] compressed = Convert(route, new PdfCore.PdfOptions { Language = "en-US" });
        byte[] uncompressed = Convert(route, new PdfCore.PdfOptions { CompressContentStreams = false });

        Assert.Contains("/Filter /FlateDecode", Encoding.ASCII.GetString(compressed), StringComparison.Ordinal);
        Assert.DoesNotContain("/Filter /FlateDecode", Encoding.ASCII.GetString(uncompressed), StringComparison.Ordinal);
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
        PdfCore.PdfOptions pdfOptions = new PdfCore.PdfOptions()
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
        Assert.Contains("/FontFile3 ", raw, StringComparison.Ordinal);
        Assert.Contains(Marker(route), PdfCore.PdfReadDocument.Open(pdf).ExtractText(), StringComparison.Ordinal);
    }

    [Fact]
    public void PdfA2BAndPasswordEncryptionRemainAnExplicitlyRejectedCombination() {
        PdfCore.PdfOptions options = CreatePdfA2BOptions().SetEncryption("open", "owner");
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

        PdfCore.PdfSecurityMutationResult protectedPdf = PdfCore.PdfDocument.Load(archivalPdf).Security.Encrypt(encryption);
        PdfCore.PdfSecurityMutationResult unlockedPdf = protectedPdf.ToDocument().Security.Decrypt("owner");

        Assert.True(PdfCore.PdfInspector.Probe(protectedPdf.Pdf).HasEncryption);
        Assert.False(PdfCore.PdfInspector.Probe(unlockedPdf.Pdf).HasEncryption);
        Assert.Contains(
            Marker("html"),
            PdfCore.PdfReadDocument.Open(unlockedPdf.Pdf).ExtractText(),
            StringComparison.Ordinal);
    }

    private static PdfCore.PdfOptions CreatePdfA2BOptions() {
        string? fontPath = PdfComplianceTestFonts.FindBundledOpenTypeCffFont();
        Assert.NotNull(fontPath);
        byte[] fontData = File.ReadAllBytes(fontPath!);
        return new PdfCore.PdfOptions {
                FileVersion = PdfCore.PdfFileVersion.Pdf17,
                IncludeStandardFontToUnicodeMaps = true
            }
            .ConfigurePdfAGroundwork(PdfCore.PdfComplianceProfile.PdfA2B, "en-US")
            .RequireCompliance(PdfCore.PdfComplianceProfile.PdfA2B)
            .EmbedStandardFont(PdfCore.PdfStandardFont.Helvetica, fontData, "OfficeIMO Source Serif")
            .EmbedStandardFont(PdfCore.PdfStandardFont.HelveticaBold, fontData, "OfficeIMO Source Serif");
    }

    private static byte[] Convert(string route, PdfCore.PdfOptions pdfOptions) => route switch {
        "docx" => ConvertDocx(pdfOptions),
        "xlsx" => ConvertXlsx(pdfOptions),
        "pptx" => ConvertPptx(pdfOptions),
        "html" => HtmlConversionDocument.Parse("<h1>" + Marker(route) + "</h1><p>Adapter proof</p>").ToPdf(new HtmlPdfSaveOptions {
            PdfOptions = pdfOptions,
            TextFallbacks = PdfCore.PdfTextFallbackFeatures.None
        }),
        "markdown" => OfficeIMO.Markdown.MarkdownReader.Parse("# " + Marker(route) + "\n\nAdapter proof").ToPdf(new MarkdownPdfSaveOptions {
            PdfOptions = pdfOptions,
            TextFallbacks = PdfCore.PdfTextFallbackFeatures.None
        }),
        "rtf" => CreateRtf(route).ToPdf(new RtfPdfSaveOptions {
            PdfOptions = pdfOptions
        }),
        "asciidoc" => AsciiDocDocument.Parse("= " + Marker(route) + "\n\nAdapter proof").Document.ToPdf(
            new AsciiDocPdfSaveOptions {
                MarkdownOptions = CreateMarkdownOptions(pdfOptions)
            }),
        "latex" => LatexDocument.Parse(
                "\\documentclass{article}\\begin{document}\\section{" + Marker(route) + "}Adapter proof\\end{document}")
            .Document.ToPdf(new LatexPdfSaveOptions {
                MarkdownOptions = CreateMarkdownOptions(pdfOptions)
            }),
        "mhtml" => new MhtmlDocument(
                "<h1>" + Marker(route) + "</h1><p>Adapter proof</p>",
                contentLocation: "https://proof.officeimo.test/report.html")
            .ToPdf(CreateHtmlOptions(pdfOptions)),
        "onenote" => CreateOneNote(route).ToPdf(new OneNotePdfSaveOptions {
            MarkdownOptions = CreateMarkdownOptions(pdfOptions)
        }),
        "odt" => CreateOdt(route).ToPdf(pdfOptions: new WordPdfSaveOptions {
            PdfOptions = pdfOptions,
            TextFallbacks = PdfCore.PdfTextFallbackFeatures.None
        }),
        "ods" => CreateOds(route).ToPdf(pdfOptions: new ExcelPdfSaveOptions {
            PdfOptions = pdfOptions,
            TextFallbacks = PdfCore.PdfTextFallbackFeatures.None
        }),
        "odp" => CreateOdp(route).ToPdf(pdfOptions: new PowerPointPdfSaveOptions {
            PdfOptions = pdfOptions,
            TextFallbacks = PdfCore.PdfTextFallbackFeatures.None
        }),
        "visio" => CreateVisio(route).ToPdf(new VisioPdfSaveOptions {
            ProjectionOptions = new PdfCore.PdfProjectionOptions { PdfOptions = pdfOptions }
        }),
        _ => throw new ArgumentOutOfRangeException(nameof(route))
    };

    private static HtmlPdfSaveOptions CreateHtmlOptions(PdfCore.PdfOptions pdfOptions) => new HtmlPdfSaveOptions {
        PdfOptions = pdfOptions,
        TextFallbacks = PdfCore.PdfTextFallbackFeatures.None
    };

    private static MarkdownPdfSaveOptions CreateMarkdownOptions(PdfCore.PdfOptions pdfOptions) =>
        new MarkdownPdfSaveOptions {
            PdfOptions = pdfOptions,
            TextFallbacks = PdfCore.PdfTextFallbackFeatures.None
        };

    private static byte[] ConvertWithAutomaticFallbacks(string route, string text) => route switch {
        "docx" => ConvertAutomaticDocx(text),
        "xlsx" => ConvertAutomaticXlsx(text),
        "pptx" => ConvertAutomaticPptx(text),
        "markdown" => OfficeIMO.Markdown.MarkdownReader.Parse("# " + text + "\n\nAdapter proof").ToPdf(),
        _ => throw new ArgumentOutOfRangeException(nameof(route))
    };

    private static byte[] ConvertAutomaticDocx(string text) {
        using WordDocument document = WordDocument.Create();
        document.AddParagraph(text);
        document.AddParagraph("Adapter proof");
        return document.ToPdf();
    }

    private static byte[] ConvertAutomaticXlsx(string text) {
        using ExcelDocument document = ExcelDocument.Create();
        ExcelSheet sheet = document.AddWorksheet("Proof");
        sheet.Cell(1, 1, text);
        sheet.Cell(2, 1, "Adapter proof");
        return document.ToPdf();
    }

    private static byte[] ConvertAutomaticPptx(string text) {
        using PowerPointPresentation presentation = PowerPointPresentation.Create();
        presentation.AddSlide().AddTextBoxPoints(
            text + "\nAdapter proof",
            leftPoints: 36,
            topPoints: 36,
            widthPoints: 420,
            heightPoints: 100);
        return presentation.ToPdf();
    }

    private static byte[] ConvertDocx(PdfCore.PdfOptions pdfOptions) {
        using WordDocument document = WordDocument.Create();
        document.AddParagraph(Marker("docx"));
        document.AddParagraph("Adapter proof");
        return document.ToPdf(new WordPdfSaveOptions {
            PdfOptions = pdfOptions,
            TextFallbacks = PdfCore.PdfTextFallbackFeatures.None
        });
    }

    private static byte[] ConvertXlsx(PdfCore.PdfOptions pdfOptions) {
        using ExcelDocument document = ExcelDocument.Create();
        ExcelSheet sheet = document.AddWorksheet("Proof");
        sheet.Cell(1, 1, Marker("xlsx"));
        sheet.Cell(2, 1, "Adapter proof");
        return document.ToPdf(new ExcelPdfSaveOptions {
            PdfOptions = pdfOptions,
            TextFallbacks = PdfCore.PdfTextFallbackFeatures.None
        });
    }

    private static byte[] ConvertPptx(PdfCore.PdfOptions pdfOptions) {
        using PowerPointPresentation presentation = PowerPointPresentation.Create();
        presentation.AddSlide().AddTextBoxPoints(
            Marker("pptx") + "\nAdapter proof",
            leftPoints: 36,
            topPoints: 36,
            widthPoints: 420,
            heightPoints: 100);
        return presentation.ToPdf(new PowerPointPdfSaveOptions {
            PdfOptions = pdfOptions,
            TextFallbacks = PdfCore.PdfTextFallbackFeatures.None
        });
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

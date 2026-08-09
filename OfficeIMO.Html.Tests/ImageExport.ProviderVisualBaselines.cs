using System.IO.Compression;
using OfficeIMO.Drawing;
using OfficeIMO.Email;
using OfficeIMO.Epub;
using OfficeIMO.Epub.Image;
using OfficeIMO.Html;
using OfficeIMO.OneNote;
using OfficeIMO.PowerPoint;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests;

[Trait("Category", "ImageExportVisualGate")]
public sealed class ImageExportProviderVisualBaselineTests {
    private const string UpdateVariable = "OFFICEIMO_UPDATE_IMAGE_EXPORT_BASELINES";
    private const string PortableFontFamily = "Carlito";
    private static readonly Lazy<byte[]> PortableRegularFont = new(() => LoadPortableFont("Carlito-Regular.ttf"));
    private static readonly Lazy<byte[]> PortableBoldFont = new(() => LoadPortableFont("Carlito-Bold.ttf"));

    [Fact]
    public void ManagedProvidersMatchApprovedPremiumRasterBaselines() {
        AssertBaseline("word-premium-page.png", RenderWord());
        AssertBaseline("powerpoint-premium-slide.png", RenderPowerPoint());
        AssertBaseline("html-premium-card.png", RenderHtml());
        AssertBaseline("onenote-premium-page.png", RenderOneNote());
        AssertBaseline("email-premium-message.png", RenderEmail());
        AssertBaseline("epub-premium-chapter.png", RenderEpub());
    }

    private static byte[] RenderWord() {
        using var stream = new MemoryStream();
        using WordDocument document = WordDocument.Create(stream);
        document.PageSettings.Width = 7200;
        document.PageSettings.Height = 4320;
        document.Margins.Type = WordMargin.Narrow;
        document.AddParagraph("Quarterly delivery brief")
            .SetFontSize(22)
            .SetBold()
            .SetFontFamily(PortableFontFamily)
            .SetColor(OfficeColor.FromRgb(30, 64, 175));
        document.AddParagraph("Polished previews with predictable output.")
            .SetFontFamily(PortableFontFamily);
        return document.ToImage()
            .WithFont(PortableFontFamily, PortableRegularFont.Value)
            .WithFont(PortableFontFamily, PortableBoldFont.Value, OfficeFontStyle.Bold)
            .FitWithin(360, 360)
            .AsPng()
            .Export()
            .Bytes;
    }

    private static byte[] RenderPowerPoint() {
        using var stream = new MemoryStream();
        using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
        presentation.SlideSize.SetSizePoints(480, 270);
        PowerPointSlide slide = presentation.AddSlide();
        slide.BackgroundColor = "F8FAFC";
        PowerPointTextBox title = slide.AddTextBoxPoints("Delivery dashboard", 28, 22, 300, 34);
        title.FontSize = 24;
        title.Bold = true;
        title.FontName = PortableFontFamily;
        title.Color = "1E3A8A";
        PowerPointTextBox status = slide.AddTextBoxPoints("On track", 30, 88, 170, 64);
        status.FontSize = 18;
        status.Bold = true;
        status.FontName = PortableFontFamily;
        status.Color = "166534";
        PowerPointTextBox evidence = slide.AddTextBoxPoints("Evidence ready", 260, 88, 170, 64);
        evidence.FontSize = 18;
        evidence.Bold = true;
        evidence.FontName = PortableFontFamily;
        evidence.Color = "155E75";
        return slide.ToImage()
            .WithFont(PortableFontFamily, PortableRegularFont.Value)
            .WithFont(PortableFontFamily, PortableBoldFont.Value, OfficeFontStyle.Bold)
            .AsPng()
            .Export()
            .Bytes;
    }

    private static byte[] RenderHtml() {
        const string html = """
            <style>
              body{font:16px Carlito,sans-serif;background:#eef2ff;color:#172554;padding:24px}
              article{background:white;border:1px solid #c7d2fe;border-radius:12px;padding:24px}
              h1{color:#1d4ed8;margin:0 0 12px}.badge{color:#166534;background:#dcfce7;padding:5px 9px}
            </style>
            <article><span class='badge'>Ready</span><h1>Premium export</h1><p>One rendering contract for polished document previews.</p></article>
            """;
        HtmlConversionDocument document = HtmlConversionDocument.Parse(html);
        var options = new HtmlRenderOptions {
            MaximumOutputWidth = 360,
            MaximumOutputHeight = 360
        };
        AddPortableFonts(options.Fonts);
        return document.ExportImage(OfficeImageExportFormat.Png, options).Bytes;
    }

    private static byte[] RenderOneNote() {
        var page = new OneNotePage {
            Title = "Research canvas",
            PageSize = OneNotePageSize.IndexCard
        };
        var outline = new OneNoteOutline {
            Layout = new OneNoteLayout { X = 0.4D, Y = 2.1D, Width = 4.8D }
        };
        var paragraph = new OneNoteParagraph();
        var introduction = new OneNoteTextRun { Text = "Key finding: " };
        introduction.Style.FontFamily = PortableFontFamily;
        paragraph.Runs.Add(introduction);
        var emphasized = new OneNoteTextRun { Text = "shared contracts keep previews consistent." };
        emphasized.Style.Bold = true;
        emphasized.Style.ColorArgb = 0xFF1D4ED8;
        emphasized.Style.FontFamily = PortableFontFamily;
        paragraph.Runs.Add(emphasized);
        outline.Children.Add(paragraph);
        page.Outlines.Add(outline);
        return page.ToImage()
            .ConfigureOptions(options => options.DefaultFont = new OfficeFontInfo(PortableFontFamily, 11D))
            .WithFont(PortableFontFamily, PortableRegularFont.Value)
            .WithFont(PortableFontFamily, PortableBoldFont.Value, OfficeFontStyle.Bold)
            .FitWithin(360, 360)
            .AsPng()
            .Export()
            .Bytes;
    }

    private static byte[] RenderEmail() {
        var email = new EmailDocument {
            Subject = "Design review approved",
            From = new EmailAddress("studio@example.com", "Studio"),
            Date = new DateTimeOffset(2026, 8, 9, 9, 30, 0, TimeSpan.Zero)
        };
        email.Recipients.Add(new EmailRecipient(
            EmailRecipientKind.To,
            new EmailAddress("reader@example.com", "Reader")));
        email.Body.Html = "<div style='font-family:Carlito,sans-serif'><h2 style='color:#1d4ed8'>Approved</h2><p>The image export review is ready for delivery.</p></div>";
        var options = new EmailImageExportOptions {
            MaximumOutputWidth = 360,
            MaximumOutputHeight = 360,
            DefaultFontFamily = PortableFontFamily
        };
        AddPortableFonts(options.Fonts);
        return email.ExportImage(OfficeImageExportFormat.Png, options).Bytes;
    }

    private static byte[] RenderEpub() {
        using var package = new MemoryStream(CreateEpub());
        EpubDocument book = EpubDocument.Load(
            package,
            new EpubReadOptions { IncludeRawHtml = true });
        var options = new EpubImageExportOptions {
            MaximumOutputWidth = 360,
            MaximumOutputHeight = 360
        };
        AddPortableFonts(options.Fonts);
        return Assert.Single(book.ExportImages(OfficeImageExportFormat.Png, options)).Bytes;
    }

    private static byte[] CreateEpub() {
        using var output = new MemoryStream();
        using (var archive = new ZipArchive(output, ZipArchiveMode.Create, leaveOpen: true)) {
            Write(archive, "mimetype", "application/epub+zip", CompressionLevel.NoCompression);
            Write(
                archive,
                "META-INF/container.xml",
                "<?xml version=\"1.0\"?><container xmlns=\"urn:oasis:names:tc:opendocument:xmlns:container\" version=\"1.0\"><rootfiles><rootfile full-path=\"OEBPS/content.opf\" media-type=\"application/oebps-package+xml\"/></rootfiles></container>");
            Write(
                archive,
                "OEBPS/content.opf",
                "<?xml version=\"1.0\"?><package xmlns=\"http://www.idpf.org/2007/opf\" version=\"3.0\" unique-identifier=\"id\"><metadata xmlns:dc=\"http://purl.org/dc/elements/1.1/\"><dc:identifier id=\"id\">premium</dc:identifier><dc:title>Premium chapter</dc:title></metadata><manifest><item id=\"chapter\" href=\"chapter.xhtml\" media-type=\"application/xhtml+xml\"/></manifest><spine><itemref idref=\"chapter\"/></spine></package>");
            Write(
                archive,
                "OEBPS/chapter.xhtml",
                "<?xml version=\"1.0\"?><html xmlns=\"http://www.w3.org/1999/xhtml\"><head><title>Visual systems</title><style>body{font-family:Carlito;color:#172554}h1{color:#1d4ed8}</style></head><body><h1>Visual systems</h1><p>Premium output, predictable contracts, and portable rendering.</p></body></html>");
        }
        return output.ToArray();
    }

    private static void Write(
        ZipArchive archive,
        string path,
        string value,
        CompressionLevel compression = CompressionLevel.Optimal) {
        ZipArchiveEntry entry = archive.CreateEntry(path, compression);
        using Stream stream = entry.Open();
        byte[] bytes = System.Text.Encoding.UTF8.GetBytes(value);
        stream.Write(bytes, 0, bytes.Length);
    }

    private static void AssertBaseline(string fileName, byte[] actualPng) {
        string baselineDirectory = Path.Combine(GetProjectRoot(), "VisualBaselines", "ImageExportProviders");
        string baselinePath = Path.Combine(baselineDirectory, fileName);
        if (string.Equals(Environment.GetEnvironmentVariable(UpdateVariable), "1", StringComparison.Ordinal)) {
            Directory.CreateDirectory(baselineDirectory);
            File.WriteAllBytes(baselinePath, actualPng);
            return;
        }

        Assert.True(File.Exists(baselinePath), "Missing image-export visual baseline: " + baselinePath);
        OfficeRasterImage expected = VisualBaselineTestSupport.DecodePng(
            File.ReadAllBytes(baselinePath),
            "Image-export visual baseline is not a supported PNG file.");
        VisualRasterComparison comparison = VisualBaselineTestSupport.CompareRasterImages(
            OfficePngWriter.Encode(expected),
            actualPng,
            channelTolerance: 6,
            allowedDifferentPixels: Math.Max(4, expected.Width * expected.Height / 100),
            maximumMeanAbsoluteError: 2D,
            maximumRootMeanSquareError: 14D,
            maximumMeanLuminanceError: 3D);
        Assert.True(
            comparison.Passed,
            fileName + " changed: " + comparison.DifferentPixels + " pixels differ; max channel delta " + comparison.MaxChannelDelta + ".");
    }

    private static string GetProjectRoot() {
        var directory = new DirectoryInfo(AppContext.BaseDirectory);
        while (directory != null) {
            if (File.Exists(Path.Combine(directory.FullName, "OfficeIMO.Html.Tests.csproj"))) {
                return directory.FullName;
            }
            directory = directory.Parent;
        }
        throw new DirectoryNotFoundException("Could not locate the OfficeIMO.Html.Tests project root.");
    }

    private static byte[] LoadPortableFont(string fileName) {
        DirectoryInfo repositoryRoot = Directory.GetParent(GetProjectRoot())
            ?? throw new DirectoryNotFoundException("Could not locate the OfficeIMO repository root.");
        string fontPath = Path.Combine(
            repositoryRoot.FullName,
            "Website",
            "Apps",
            "OfficeIMO.Web.Converter",
            "Assets",
            "Fonts",
            fileName);
        return File.ReadAllBytes(fontPath);
    }

    private static void AddPortableFonts(OfficeFontFaceCollection fonts) {
        fonts.Add(PortableFontFamily, PortableRegularFont.Value, OfficeFontStyle.Regular);
        fonts.Add(PortableFontFamily, PortableBoldFont.Value, OfficeFontStyle.Bold);
    }
}

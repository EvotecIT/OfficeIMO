using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Pdf;
using OfficeIMO.TestAssets;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class WordFieldLayoutPaginationFontsTests {
    private const string SourceFamily = "OfficeIMO Field Layout Source";
    private const string WideFamily = "OfficeIMO Field Layout Wide";

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void TableOfContentsPagesMatchThePdfExportedWithTheSameOptions(bool portable) {
        WordPdfExportOptions options = CreateWideFontOptions(portable);
        using WordDocument document = CreateChapters(SourceFamily);

        document.UpdateFields(new WordFieldUpdateOptions { UpdateTablesOfContents = true }.WithPdfLayoutPagination(options));
        Dictionary<string, int> tocPages = ReadTocPagesByBookmark(document);
        Dictionary<string, int> pdfPages = ReadHeadingPages(document, options);

        Assert.Equal(12, tocPages.Count);
        Assert.Equal(tocPages.OrderBy(_ => _.Key), pdfPages.Where(_ => tocPages.ContainsKey(_.Key)).OrderBy(_ => _.Key));
        // The wide registered font changes the layout, so the pages above come from the caller's fonts.
        using WordDocument withoutFonts = CreateChapters(SourceFamily);
        withoutFonts.UpdateFields(new WordFieldUpdateOptions { UpdateTablesOfContents = true }.WithPdfLayoutPagination());
        Assert.NotEqual(tocPages.Values.Max(), ReadTocPagesByBookmark(withoutFonts).Values.Max());
    }

    [Fact]
    public void PortableResourcePolicyKeepsInstalledFontsOutOfTableOfContentsPages() {
        var options = new WordPdfExportOptions { ResourcePolicy = PdfResourcePolicy.CreatePortableDeterministic() };
        using WordDocument document = CreateChapters("Calibri");

        document.UpdateFields(new WordFieldUpdateOptions { UpdateTablesOfContents = true }.WithPdfLayoutPagination(options));
        Dictionary<string, int> tocPages = ReadTocPagesByBookmark(document);

        using var stream = new MemoryStream();
        document.SaveAsPdf(stream, options);
        using (var pdf = PdfPigDocument.Open(stream.ToArray())) {
            // Only the standard PDF fonts lay the text out, so the page numbers do not depend on the host.
            Assert.All(pdf.GetPages().SelectMany(page => page.Letters), letter =>
                Assert.DoesNotContain("Calibri", letter.FontName, StringComparison.OrdinalIgnoreCase));
        }
        Dictionary<string, int> pdfPages = ReadHeadingPages(document, options);
        Assert.Equal(tocPages.OrderBy(_ => _.Key), pdfPages.Where(_ => tocPages.ContainsKey(_.Key)).OrderBy(_ => _.Key));
    }

    [Fact]
    public void LayoutOptionsAreCopiedWhenSelected() {
        WordPdfExportOptions options = CreateWideFontOptions(portable: true);
        WordFieldUpdateOptions selected = new WordFieldUpdateOptions { UpdateTablesOfContents = true }.WithPdfLayoutPagination(options);
        options.PdfOptions = null;
        options.ResourcePolicy = PdfResourcePolicy.CreateDefault();

        using WordDocument document = CreateChapters(SourceFamily);
        document.UpdateFields(selected);
        Assert.Equal(ReadTocPagesByBookmark(document).OrderBy(_ => _.Key),
            ReadHeadingPages(document, CreateWideFontOptions(portable: true)).Where(_ => _.Key.StartsWith("_OfficeIMO_Toc_", StringComparison.Ordinal)).OrderBy(_ => _.Key));
    }

    [Fact]
    public void EstimatedPageCountUsesTheSameLayoutAsThePdf() {
        WordPdfExportOptions options = CreateWideFontOptions(portable: true);
        using WordDocument document = CreateChapters(SourceFamily);
        using var stream = new MemoryStream();
        document.SaveAsPdf(stream, options);

        int pages = document.GetEstimatedPageCount(options);
        Assert.Equal(PdfReadDocument.Open(stream.ToArray()).Pages.Count, pages);
        Assert.Equal(pages, document.GetEstimatedImagePageCount(new WordImageExportOptions { LayoutOptions = options }));
        Assert.NotEqual(document.GetEstimatedPageCount(), pages);
    }

    private static WordPdfExportOptions CreateWideFontOptions(bool portable) {
        var pdf = new PdfCreateOptions();
        pdf.RegisterNamedFontFamily(new PdfEmbeddedFontFamily(WideFamily,
            ManagedTextShapingTestAssets.CreateFontWithUniformAdvance(1000, 900, Enumerable.Range(32, 95).ToArray())));
        pdf.RegisterFontFamilySubstitution(SourceFamily, WideFamily, PdfFontFamilySubstitutionImpact.Compatible);
        var options = new WordPdfExportOptions { PdfOptions = pdf };
        if (portable) options.ResourcePolicy = PdfResourcePolicy.CreatePortableDeterministic();
        return options;
    }

    /// <summary>Chapters whose headings move to later pages only through text wrapping: there are no page breaks.</summary>
    private static WordDocument CreateChapters(string fontFamily) {
        WordDocument document = WordDocument.Create();
        document.Settings.FontFamily = fontFamily;
        document.AddTableOfContent();
        for (int chapter = 1; chapter <= 12; chapter++) {
            document.AddParagraph("Chapter " + chapter).Style = WordParagraphStyles.Heading1;
            for (int line = 0; line < 2 + chapter % 3; line++) {
                document.AddParagraph("Chapter " + chapter + " paragraph " + line + " " + string.Join(" ", Enumerable.Repeat("measured text", 30)));
            }
        }
        return document;
    }

    private static Dictionary<string, int> ReadTocPagesByBookmark(WordDocument document) {
        var pages = new Dictionary<string, int>(StringComparer.Ordinal);
        foreach (Hyperlink hyperlink in document.TableOfContent!.SdtBlock.Descendants<Hyperlink>().Where(_ => !string.IsNullOrEmpty(_.Anchor?.Value))) {
            pages[hyperlink.Anchor!.Value!] = int.Parse(hyperlink.Descendants<Text>().Last().Text, System.Globalization.CultureInfo.InvariantCulture);
        }
        return pages;
    }

    private static Dictionary<string, int> ReadHeadingPages(WordDocument document, WordPdfExportOptions options) {
        using var stream = new MemoryStream();
        document.SaveAsPdf(stream, options);
        var pages = new Dictionary<string, int>(StringComparer.Ordinal);
        foreach (PdfNamedDestination destination in PdfReadDocument.Open(stream.ToArray()).NamedDestinations) {
            if (destination.PageNumber is int page) pages[destination.Name] = page;
        }
        return pages;
    }
}

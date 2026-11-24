using OfficeIMO.Drawing;
using OfficeIMO.Epub;
using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using OfficeIMO.Pdf;

namespace OfficeIMO.Tests;

public sealed class OfficeConverterEpubTests {
    [Fact]
    public void StyledEpubConversionUsesSpineOrderLinkedStylesImagesAndNavigation() {
        var book = EpubPublication.Create("A publication", "en", "urn:officeimo:styled-epub-pdf-test");
        book.AddResource("figure", "EPUB/art/dot.png", "image/png",
            OfficeRasterImageEncoder.Encode(new OfficeRasterImage(8, 8, OfficeColor.FromRgb(38, 95, 112)), OfficeImageExportFormat.Png));
        book.AddStylesheet("print", "EPUB/styles/book.css", "h1 { font-size: 22pt } img { width: 24px; height: 24px } @page { margin: 36pt; @bottom-center { content: 'Book page ' counter(page) } }");
        book.AddChapter("first", "EPUB/text/z.xhtml", "Opening",
            "<h1 id='heading'>Opening text</h1><p><a href='a.xhtml#heading'>Continue reading</a></p>" +
            "<p><a href='a.xhtml'>Start of closing chapter</a></p><p><a href='a.xhtml#top'>Authored top anchor</a></p>" +
            "<img src='../art/dot.png' alt='Opening figure' />",
            ["print"]);
        book.AddChapter("second", "EPUB/text/a.xhtml", "Closing",
            "<h1 id='heading'>Closing text</h1><p id='top'><a href='#heading'>Back to heading</a></p><img src='../art/dot.png' alt='Closing figure' />",
            ["print"]);
        book.SetNavigation([new EpubNavigationEntry("Beginning", "EPUB/text/z.xhtml",
            [new EpubNavigationEntry("End of the story", "EPUB/text/a.xhtml#heading")])]);
        var navigation = book.GetContentXml("navigation");
        var parent = navigation.Descendants().First(element => element.Name.LocalName == "a");
        parent.Name = parent.Name.Namespace + "span";
        parent.RemoveAttributes();
        book.SetContentXml("navigation", navigation);

        using var source = new MemoryStream(book.Write().Bytes);
        using var output = new MemoryStream();
        OfficeConversionReport report = OfficeConverter.Convert(source, OfficeDocumentFormat.Epub, output, OfficeDocumentFormat.Pdf,
            new OfficeConversionOptions { Html = new HtmlPdfExportOptions { FidelityPolicy = HtmlRenderFidelityPolicy.FailOnLoss } });
        using var pdf = PdfDocument.Load(output.ToArray(), new PdfLoadOptions { IncludeArtifactText = true });

        Assert.Equal(2, report.ItemsWritten);
        string text = pdf.ExtractText();
        Assert.Contains("Opening text", text, StringComparison.Ordinal);
        Assert.Contains("Closing text", text, StringComparison.Ordinal);
        Assert.Contains("Book page 1", text, StringComparison.Ordinal);
        Assert.Contains("Book page 2", text, StringComparison.Ordinal);
        PdfLinkAnnotation link = Assert.Single(pdf.Inspect().LinkAnnotations, item => item.Contents == "Continue reading");
        Assert.Equal(1, link.PageNumber);
        Assert.Equal(2, Assert.Single(pdf.NamedDestinations, target => target.Name == link.DestinationName).PageNumber);
        PdfNamedDestination start = FindTarget("Start of closing chapter");
        PdfNamedDestination authored = FindTarget("Authored top anchor");
        Assert.Equal(2, start.PageNumber);
        Assert.Equal(2, authored.PageNumber);
        Assert.True(start.DestinationTop > authored.DestinationTop);
        Assert.True(text.IndexOf("Opening text", StringComparison.Ordinal) < text.IndexOf("Closing text", StringComparison.Ordinal));
        Assert.Equal(22D, Assert.Single(pdf.Read().TextBlocks, block => block.Text == "Opening text").FontSize, precision: 2);
        Assert.Equal(new[] { 1, 2 }, pdf.Images.Placements().Select(placement => placement.PageNumber));
        PdfOutlineItem root = Assert.Single(pdf.Outlines);
        Assert.Equal("Beginning", root.Title);
        Assert.Equal(2, root.PageNumber);
        PdfOutlineItem child = Assert.Single(root.Children);
        Assert.Equal("End of the story", child.Title);
        Assert.Equal(2, child.PageNumber);

        PdfTaggedContentInfo tagged = Assert.IsType<PdfTaggedContentInfo>(pdf.TaggedContent);
        var structures = tagged.StructureElements.ToDictionary(element => element.ObjectNumber);
        PdfStructureElementInfo document = Assert.Single(tagged.StructureElements, element => element.StructureType == "Document");
        PdfStructureElementInfo[] sections = document.ChildElementObjectNumbers.Select(id => structures[id]).ToArray();
        PdfStructureElementInfo[] headings = tagged.StructureElements.Where(element => element.StructureType == "H1").ToArray();
        Assert.Equal(2, headings.Length);
        Assert.Equal(new[] { "Sect", "Sect" }, sections.Select(section => section.StructureType));
        Assert.Equal(sections.Select(section => (int?)section.ObjectNumber), headings.Select(heading => heading.ParentObjectNumber));

        PdfNamedDestination FindTarget(string label) {
            PdfLinkAnnotation annotation = Assert.Single(pdf.Inspect().LinkAnnotations, item => item.Contents == label);
            return Assert.Single(pdf.NamedDestinations, target => target.Name == annotation.DestinationName);
        }
    }

    [Fact]
    public void StyledEpubConversionKeepsMixedFixedAndReflowableSpinePageSizes() {
        var book = EpubPublication.Create("Mixed layout", "en", "urn:officeimo:epub-pdf-fixed-layout");
        book.AddChapter("reflow", "EPUB/text/reflow.xhtml", "Reflow", "<h1>Reflowable chapter</h1>");
        book.AddChapter("fixed", "EPUB/text/fixed.xhtml", "Fixed", "<h1>Fixed-layout chapter</h1>");
        book.SetFixedLayoutPage("fixed", new EpubFixedLayoutPage(400, 240));

        using var source = new MemoryStream(book.Write().Bytes);
        using var output = new MemoryStream();
        OfficeConversionReport report = OfficeConverter.Convert(source, OfficeDocumentFormat.Epub, output, OfficeDocumentFormat.Pdf);
        using var pdf = PdfDocument.Load(output.ToArray());
        Assert.Equal(2, report.ItemsWritten);
        IReadOnlyList<PdfPageInfo> pages = pdf.Inspect().Pages;
        Assert.NotEqual(300D, pages[0].Width);
        Assert.Equal(300D, pages[1].Width, 2);
        Assert.Equal(180D, pages[1].Height, 2);
        Assert.Contains("Fixed-layout chapter", pdf.ExtractText(), StringComparison.Ordinal);
    }

    [Fact]
    public void StyledEpubOutlineTargetsTaggedContentForPdfUa2() {
        var book = EpubPublication.Create("Tagged publication", "en", "urn:officeimo:epub-pdfua-outline");
        book.AddChapter("one", "EPUB/text/one.xhtml", "Opening", "<h1>First heading</h1><p>First chapter.</p>");
        book.AddChapter("two", "EPUB/text/two.xhtml", "Closing", "<h1 id='closing'>Second heading</h1><p>Second chapter.</p>");
        book.SetNavigation([new EpubNavigationEntry("Opening", "EPUB/text/one.xhtml"),
            new EpubNavigationEntry("Closing", "EPUB/text/two.xhtml#closing")]);
        string font = Path.Combine(AppContext.BaseDirectory, "Fixtures", "OfficeIMOBaselineSans-Regular.ttf");
        var html = new HtmlPdfExportOptions { PdfOptions = new PdfCreateOptions()
            .UseFontFamily("OfficeIMO Baseline Sans", font).UsePdfUa(PdfComplianceProfile.PdfUa2, "en") };
        using var source = new MemoryStream(book.Write().Bytes);
        using var output = new MemoryStream();
        OfficeConverter.Convert(source, OfficeDocumentFormat.Epub, output, OfficeDocumentFormat.Pdf, new OfficeConversionOptions { Html = html });
        using var pdf = PdfDocument.Load(output.ToArray());
        Assert.Equal(new int?[] { 1, 2 }, pdf.Outlines.Select(outline => outline.PageNumber));
        Assert.True(pdf.HasTaggedContent);
    }
}

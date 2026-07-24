using System.Text;
using System.Text.RegularExpressions;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfOutlineTests {
    [Fact]
    public void CanvasOutline_WritesNestedBookmarksAtAbsolutePageCoordinates() {
        byte[] bytes = PdfDocument.Create()
            .Page(page => page
                .Size(400D, 600D)
                .Margin(0D)
                .Canvas(canvas => canvas
                    .Outline("Canvas chapter", 1, 72D)
                    .Outline("Canvas detail", 2, 144D)
                    .Text("Visible content", 36D, 72D, 200D, 24D)))
            .ToBytes();

        PdfDocumentInfo info = PdfInspector.Inspect(bytes);
        PdfOutlineItem chapter = Assert.Single(info.Outlines);
        Assert.Equal("Canvas chapter", chapter.Title);
        Assert.Equal(1, chapter.Level);
        Assert.Equal(1, chapter.PageNumber);
        Assert.Equal(528D, chapter.DestinationTop);

        PdfOutlineItem detail = Assert.Single(chapter.Children);
        Assert.Equal("Canvas detail", detail.Title);
        Assert.Equal(2, detail.Level);
        Assert.Equal(456D, detail.DestinationTop);
    }

    [Fact]
    public void CanvasOutline_RejectsInvalidArguments() {
        var canvas = new PdfPageCanvas();

        Assert.Throws<ArgumentNullException>(() => canvas.Outline(null!, 1, 0D));
        Assert.Throws<ArgumentException>(() => canvas.Outline(" ", 1, 0D));
        Assert.Throws<ArgumentOutOfRangeException>(() => canvas.Outline("Title", 0, 0D));
        Assert.Throws<ArgumentOutOfRangeException>(() =>
            canvas.Outline("Too deep", PdfPageCanvas.MaximumOutlineLevel + 1, 0D));
        Assert.Throws<ArgumentOutOfRangeException>(() => canvas.Outline("Title", 1, double.NaN));
    }

    [Fact]
    public void CreateOutlineFromHeadings_WritesNestedBookmarksAndInspectorReadsThem() {
        byte[] bytes = PdfDocument.Create(new PdfOptions {
                CreateOutlineFromHeadings = true,
                DefaultFont = PdfStandardFont.Helvetica
            })
            .H1("Executive summary")
            .Paragraph(p => p.Text("Summary body."))
            .H2("Risk posture")
            .Paragraph(p => p.Text("Risk body."))
            .PageBreak()
            .H1("Appendix")
            .Paragraph(p => p.Text("Appendix body."))
            .ToBytes();

        string pdfText = Encoding.ASCII.GetString(bytes);
        Assert.Contains("/Outlines", pdfText);
        Assert.Contains("/PageMode /UseOutlines", pdfText);
        Assert.Contains("/Type /Outlines", pdfText);
        Assert.Contains("/Title (Executive summary)", pdfText);
        Assert.Contains("/Title (Risk posture)", pdfText);
        Assert.Contains("/Title (Appendix)", pdfText);
        Assert.Contains("/First", pdfText);
        Assert.Contains("/Last", pdfText);
        Assert.Contains("/Parent", pdfText);
        Assert.Contains("/Count 3", pdfText);
        Assert.Equal(3, CountOccurrences(pdfText, "/Dest ["));
        Assert.Equal(3, Regex.Matches(pdfText, @"/Dest \[\d+ 0 R /XYZ 0 \d+(?:\.\d+)? 0\]").Count);

        PdfDocumentInfo info = PdfInspector.Inspect(bytes);

        Assert.Equal(2, info.Outlines.Count);
        Assert.Equal("Executive summary", info.Outlines[0].Title);
        Assert.Equal(1, info.Outlines[0].Level);
        Assert.Equal(1, info.Outlines[0].PageNumber);
        Assert.True(info.Outlines[0].IsExpanded);
        Assert.True(info.Outlines[0].DestinationTop > 0);

        PdfOutlineItem child = Assert.Single(info.Outlines[0].Children);
        Assert.Equal("Risk posture", child.Title);
        Assert.Equal(2, child.Level);
        Assert.Equal(1, child.PageNumber);
        Assert.True(child.IsExpanded);

        Assert.Equal("Appendix", info.Outlines[1].Title);
        Assert.Equal(1, info.Outlines[1].Level);
        Assert.Equal(2, info.Outlines[1].PageNumber);
        Assert.True(info.Outlines[1].IsExpanded);
        Assert.Empty(info.Outlines[1].Children);
    }

    [Fact]
    public void OutlineExpansionLevel_CollapsesGeneratedOutlineBranchesAndInspectorReadsState() {
        var options = new PdfOptions {
            CreateOutlineFromHeadings = true,
            OutlineExpansionLevel = 0,
            DefaultFont = PdfStandardFont.Helvetica
        };

        PdfOptions clone = options.Clone();
        Assert.Equal(0, clone.OutlineExpansionLevel);

        byte[] bytes = PdfDocument.Create(options)
            .H1("Executive summary")
            .H2("Risk posture")
            .H3("Operational detail")
            .PageBreak()
            .H1("Appendix")
            .ToBytes();

        string pdfText = Encoding.ASCII.GetString(bytes);
        Assert.Matches(@"/Type /Outlines /First \d+ 0 R /Last \d+ 0 R /Count 2", pdfText);
        Assert.Contains("/Count -2", pdfText);
        Assert.Contains("/Count -1", pdfText);

        PdfDocumentInfo info = PdfInspector.Inspect(bytes);

        Assert.Equal(2, info.Outlines.Count);
        Assert.False(info.Outlines[0].IsExpanded);
        PdfOutlineItem level2 = Assert.Single(info.Outlines[0].Children);
        Assert.False(level2.IsExpanded);
        PdfOutlineItem level3 = Assert.Single(level2.Children);
        Assert.True(level3.IsExpanded);
        Assert.True(info.Outlines[1].IsExpanded);
    }

    [Fact]
    public void OutlineExpansionLevel_RejectsNegativeValues() {
        Assert.Throws<ArgumentOutOfRangeException>(() => new PdfOptions {
            OutlineExpansionLevel = -1
        });
    }

    [Fact]
    public void CreateOutlineFromHeadings_DefaultsToOffForMinimalOutput() {
        byte[] bytes = PdfDocument.Create(new PdfOptions {
                DefaultFont = PdfStandardFont.Helvetica
            })
            .H1("Heading")
            .Paragraph(p => p.Text("Body."))
            .ToBytes();

        string pdfText = Encoding.ASCII.GetString(bytes);
        Assert.DoesNotContain("/Outlines", pdfText);

        PdfDocumentInfo info = PdfInspector.Inspect(bytes);
        Assert.Empty(info.Outlines);
    }

    [Fact]
    public void Headings_RejectEmptyOrWhitespaceTitlesBeforeOutlineGeneration() {
        Assert.Throws<ArgumentNullException>(() =>
            PdfDocument.Create().H1(null!));

        Assert.Throws<ArgumentException>(() =>
            PdfDocument.Create().H1(string.Empty));

        Assert.Throws<ArgumentException>(() =>
            PdfDocument.Create().H2("   "));

        Assert.Throws<ArgumentException>(() =>
            PdfDocument.Create().H3("\t"));

        Assert.Throws<ArgumentException>(() =>
            PdfDocument.Create()
                .Compose(compose =>
                    compose.Page(page =>
                        page.Content(content =>
                            content.Row(row =>
                                row.Column(1, column => column.H1(" ")))))));
    }

    private static int CountOccurrences(string text, string value) {
        int count = 0;
        int startIndex = 0;
        while (true) {
            int index = text.IndexOf(value, startIndex, StringComparison.Ordinal);
            if (index < 0) {
                return count;
            }

            count++;
            startIndex = index + value.Length;
        }
    }
}

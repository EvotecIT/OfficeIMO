using System.Text;
using System.Text.RegularExpressions;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfOutlineTests {
    [Fact]
    public void CreateOutlineFromHeadings_WritesNestedBookmarksAndInspectorReadsThem() {
        byte[] bytes = PdfDoc.Create(new PdfOptions {
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
        Assert.True(info.Outlines[0].DestinationTop > 0);

        PdfOutlineItem child = Assert.Single(info.Outlines[0].Children);
        Assert.Equal("Risk posture", child.Title);
        Assert.Equal(2, child.Level);
        Assert.Equal(1, child.PageNumber);

        Assert.Equal("Appendix", info.Outlines[1].Title);
        Assert.Equal(1, info.Outlines[1].Level);
        Assert.Equal(2, info.Outlines[1].PageNumber);
        Assert.Empty(info.Outlines[1].Children);
    }

    [Fact]
    public void CreateOutlineFromHeadings_DefaultsToOffForMinimalOutput() {
        byte[] bytes = PdfDoc.Create(new PdfOptions {
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
            PdfDoc.Create().H1(null!));

        Assert.Throws<ArgumentException>(() =>
            PdfDoc.Create().H1(string.Empty));

        Assert.Throws<ArgumentException>(() =>
            PdfDoc.Create().H2("   "));

        Assert.Throws<ArgumentException>(() =>
            PdfDoc.Create().H3("\t"));

        Assert.Throws<ArgumentException>(() =>
            PdfDoc.Create()
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

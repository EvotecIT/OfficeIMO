using OfficeIMO.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfHeadingTableAdmissionTests {
    [Theory]
    [InlineData(false, 0, false, false)]
    [InlineData(true, 0, false, false)]
    [InlineData(false, 1, false, false)]
    [InlineData(true, 1, false, false)]
    [InlineData(false, 2, false, false)]
    [InlineData(true, 2, false, false)]
    [InlineData(true, 1, true, false)]
    [InlineData(false, 0, false, true)]
    [InlineData(true, 0, false, true)]
    public void HeadingKeepsWithTableStartingRowAdmission(bool orderedCells, int headerRows, bool keepTogether, bool oversizedGroup) {
        var options = new PdfCreateOptions {
            PageSetup = {
                Size = new PdfPageSize(global::OfficeIMO.OfficeLength.Points(260), global::OfficeIMO.OfficeLength.Points(240)),
                Margins = {
                    Left = global::OfficeIMO.OfficeLength.Points(30), Right = global::OfficeIMO.OfficeLength.Points(30),
                    Top = global::OfficeIMO.OfficeLength.Points(30), Bottom = global::OfficeIMO.OfficeLength.Points(30)
                }
            },
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = global::OfficeIMO.OfficeLength.Points(10)
        };
        var style = new PdfTableStyle {
            FontSize = global::OfficeIMO.OfficeLength.Points(10), HeaderFontSize = global::OfficeIMO.OfficeLength.Points(10),
            LineHeight = 1, CellPaddingX = global::OfficeIMO.OfficeLength.Points(2), CellPaddingY = global::OfficeIMO.OfficeLength.Points(2),
            HeaderRowCount = headerRows, FooterRowCount = oversizedGroup ? 0 : 1, MinimumBodyRowsOnFirstPage = 2,
            SpacingBefore = global::OfficeIMO.OfficeLength.Points(0), SpacingAfter = global::OfficeIMO.OfficeLength.Points(0),
            KeepTogether = keepTogether
        };
        if (oversizedGroup) {
            style.AllowRowBreakAcrossPages = false;
            style.RowMinHeights = new List<global::OfficeIMO.OfficeLength?> {
                global::OfficeIMO.OfficeLength.Points(80), global::OfficeIMO.OfficeLength.Points(110)
            };
        }
        var rows = new List<PdfTableCell[]>();
        for (int row = 0; row < headerRows + (oversizedGroup ? 2 : 5); row++) {
            string text = row < headerRows ? "Header" + row : row == headerRows + 4 ? "Footer" : "Body" + (row - headerRows);
            PdfTableCell cell = orderedCells
                ? new PdfTableCell(Array.Empty<PdfTextRun>(), null, blocks: new IPdfBlock[] {
                    new RichParagraphBlock(new[] { PdfTextRun.Normal(text) }, OfficeAlignment.Left, null,
                        new PdfParagraphStyle { FontSize = global::OfficeIMO.OfficeLength.Points(10), LineHeight = 1,
                            SpacingBefore = global::OfficeIMO.OfficeLength.Points(0), SpacingAfter = global::OfficeIMO.OfficeLength.Points(0) })
                })
                : new PdfTableCell(text);
            rows.Add(new[] { cell });
        }
        byte[] bytes = PdfDocument.Create(options)
            .Paragraph(paragraph => paragraph.AddText("Introduction"), style: new PdfParagraphStyle {
                LineHeight = 1, SpacingAfter = global::OfficeIMO.OfficeLength.Points(132)
            })
            .H3("ControlsHeading", style: new PdfHeadingStyle {
                FontSize = global::OfficeIMO.OfficeLength.Points(12), LineHeight = 1,
                SpacingBefore = global::OfficeIMO.OfficeLength.Points(0), SpacingAfter = global::OfficeIMO.OfficeLength.Points(4)
            })
            .Table(rows, style: style)
            .ToBytes();

        using var pdf = PdfPigDocument.Open(bytes);
        Assert.Equal(oversizedGroup ? 3 : 2, pdf.NumberOfPages);
        Assert.Contains("Introduction", pdf.GetPage(1).Text);
        Assert.DoesNotContain("ControlsHeading", pdf.GetPage(1).Text);
        Assert.DoesNotContain("Body0", pdf.GetPage(1).Text);
        Assert.Contains("ControlsHeading", pdf.GetPage(2).Text);
        Assert.Contains("Body0", pdf.GetPage(2).Text);
        Assert.Contains("Body1", pdf.GetPage(oversizedGroup ? 3 : 2).Text);
        if (!oversizedGroup) Assert.Contains("Footer", pdf.GetPage(2).Text);
    }
}

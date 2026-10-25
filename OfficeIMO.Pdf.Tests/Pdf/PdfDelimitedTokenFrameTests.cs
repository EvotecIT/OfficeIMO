using System;
using System.Collections.Generic;
using System.Linq;
using OfficeIMO.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfDelimitedTokenFrameTests {
    [Theory]
    [InlineData(0, false)]
    [InlineData(0, true)]
    [InlineData(1, false)]
    [InlineData(1, true)]
    [InlineData(2, false)]
    [InlineData(2, true)]
    public void DelimitedSegmentsRespectMeasuredFrameAcrossParagraphAndCellOwners(int route, bool boundedGroup) {
        string prefix = boundedGroup ? "/ALPHABET/" : "ALPHABET";
        string token = boundedGroup ? prefix + "01" : prefix + "_01_02";
        using var metric = PdfPigDocument.Open(Create(prefix, 300D, route: 0));
        var metricLetters = metric.GetPage(1).Letters;
        double prefixAdvance = metricLetters[metricLetters.Count - 1].EndBaseLine.X - metricLetters[0].StartBaseLine.X;
        var last = metricLetters[metricLetters.Count - 1];
        double halfLastAdvance = (last.EndBaseLine.X - last.StartBaseLine.X) / 2D;
        Assert.True(halfLastAdvance > 0D);

        foreach (double innerWidth in new[] { prefixAdvance - halfLastAdvance, prefixAdvance, prefixAdvance + halfLastAdvance }) {
            using var pdf = PdfPigDocument.Open(Create(token, innerWidth, route));
            var letters = pdf.GetPages().SelectMany(page => page.Letters).ToArray();
            Assert.Equal(token, string.Concat(letters.Select(letter => letter.Value)));
            double textLeft = route == 0 ? 20D : 22D;
            Assert.All(letters, letter => {
                Assert.True(letter.StartBaseLine.X >= textLeft - .001D);
                Assert.True(letter.EndBaseLine.X <= textLeft + innerWidth + .001D,
                    "A delimited segment must not gain an extra glyph of overflow permission.");
            });
        }
    }

    private static byte[] Create(string text, double innerWidth, int route) {
        double frameWidth = innerWidth + (route == 0 ? 0D : 4D);
        var options = new PdfOptions {
            PageWidth = frameWidth + 40D, PageHeight = 400D,
            MarginLeft = 20D, MarginRight = 20D, MarginTop = 20D, MarginBottom = 20D,
            DefaultFont = PdfStandardFont.Helvetica, DefaultFontSize = 10D
        };
        var document = PdfDocument.Create(options);
        if (route == 0) document.Content.Text(text, style: new PdfParagraphStyle {
            FontSize = 10D, LineHeight = 1D, SpacingAfter = 0D
        });
        else {
            var style = new PdfTableStyle {
                FontSize = 10D, LineHeight = 1D, HeaderRowCount = 0,
                CellPaddingX = 2D, CellPaddingY = 2D,
                ColumnWidthPoints = new List<double?> { frameWidth },
                MaxWidth = frameWidth, PreferredWidth = frameWidth, PreserveWidth = true
            };
            PdfTableCell cell = new(text);
            if (route == 2) {
                var inner = new TableBlock(new[] { new[] { cell } }, PdfAlign.Left, style);
                cell = new PdfTableCell(Array.Empty<PdfTextRun>(), null, blocks: new IPdfBlock[] { inner });
                style = style.Clone(); style.CellPaddingX = style.CellPaddingY = 0D;
            }
            document.Content.Table(new[] { new[] { cell } }, style: style);
        }
        return document.ToBytes();
    }
}

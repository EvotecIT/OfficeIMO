using System;
using System.Collections.Generic;
using System.IO;
using OfficeIMO.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfDocumentPaginationFidelityTests {
    [Fact]
    public void Paragraph_CustomOrphanAndWidowCountsProduceBalancedSplit() {
        var options = new PdfOptions {
            PageWidth = 260,
            PageHeight = 170,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 10
        };

        byte[] bytes = PdfDocument.Create(options)
            .Paragraph(paragraph => paragraph.Text("IntroMarker"), style: new PdfParagraphStyle {
                SpacingAfter = 45
            })
            .Paragraph(paragraph => paragraph
                .Text("BalancedLine1").LineBreak()
                .Text("BalancedLine2").LineBreak()
                .Text("BalancedLine3").LineBreak()
                .Text("BalancedLine4").LineBreak()
                .Text("BalancedLine5").LineBreak()
                .Text("BalancedLine6"), style: new PdfParagraphStyle {
                    MinimumOrphanLines = 3,
                    MinimumWidowLines = 3,
                    SpacingAfter = 0
                })
            .ToBytes();

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        Assert.Equal(2, pdf.NumberOfPages);
        Assert.Contains("BalancedLine3", pdf.GetPage(1).Text);
        Assert.DoesNotContain("BalancedLine4", pdf.GetPage(1).Text);
        Assert.Contains("BalancedLine4", pdf.GetPage(2).Text);
        Assert.Contains("BalancedLine6", pdf.GetPage(2).Text);
    }

    [Fact]
    public void Table_MinimumBodyRowsOnLastPageKeepsRowsWithFooter() {
        var options = new PdfOptions {
            PageWidth = 280,
            PageHeight = 160,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 9
        };
        var style = TableStyles.Minimal();
        style.HeaderRowCount = 1;
        style.FooterRowCount = 1;
        style.MinimumBodyRowsOnLastPage = 3;

        var rows = new List<string[]> {
            new[] { "Metric", "Value" }
        };
        for (int rowIndex = 1; rowIndex <= 7; rowIndex++) {
            rows.Add(new[] { "FinalGroupRow" + rowIndex, rowIndex.ToString() });
        }
        rows.Add(new[] { "FinalGroupTotal", "7" });

        byte[] bytes = PdfDocument.Create(options)
            .Table(rows, style: style)
            .ToBytes();

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        Assert.Equal(2, pdf.NumberOfPages);
        Assert.DoesNotContain("FinalGroupRow5", pdf.GetPage(1).Text);
        Assert.Contains("FinalGroupRow5", pdf.GetPage(2).Text);
        Assert.Contains("FinalGroupRow6", pdf.GetPage(2).Text);
        Assert.Contains("FinalGroupRow7", pdf.GetPage(2).Text);
        Assert.Contains("FinalGroupTotal", pdf.GetPage(2).Text);
    }

    [Fact]
    public void Table_MinimumBodyRowsOnFirstPagePreventsOrphanedHeader() {
        var options = new PdfOptions {
            PageWidth = 280,
            PageHeight = 180,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 10
        };
        var style = TableStyles.Minimal();
        style.HeaderRowCount = 1;
        style.RepeatHeaderRowCount = 1;
        style.FixedRowHeights = new List<double?> { 20, 20, 20 };

        byte[] bytes = PdfDocument.Create(options)
            .Paragraph(paragraph => paragraph.Text("IntroMarker"), style: new PdfParagraphStyle {
                SpacingAfter = 50
            })
            .Table(new[] {
                new[] { "OrphanHeader", "Value" },
                new[] { "FirstBodyRow", "1" },
                new[] { "SecondBodyRow", "2" }
            }, style: style)
            .ToBytes();

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        Assert.Equal(2, pdf.NumberOfPages);
        Assert.DoesNotContain("OrphanHeader", pdf.GetPage(1).Text);
        Assert.Contains("OrphanHeader", pdf.GetPage(2).Text);
        Assert.Contains("FirstBodyRow", pdf.GetPage(2).Text);
    }

    [Fact]
    public void PaginationControlsCloneAndRejectNegativeValues() {
        var paragraphStyle = new PdfParagraphStyle {
            MinimumOrphanLines = 3,
            MinimumWidowLines = 4
        };
        PdfParagraphStyle paragraphClone = paragraphStyle.Clone();
        Assert.Equal(3, paragraphClone.MinimumOrphanLines);
        Assert.Equal(4, paragraphClone.MinimumWidowLines);

        var tableStyle = new PdfTableStyle {
            MinimumBodyRowsOnFirstPage = 2,
            MinimumBodyRowsOnLastPage = 4
        };
        Assert.Equal(2, new PdfTableStyle().MinimumBodyRowsOnFirstPage);
        Assert.Equal(2, tableStyle.Clone().MinimumBodyRowsOnFirstPage);
        Assert.Equal(4, tableStyle.Clone().MinimumBodyRowsOnLastPage);

        Assert.Throws<ArgumentOutOfRangeException>(() => paragraphStyle.MinimumOrphanLines = -1);
        Assert.Throws<ArgumentOutOfRangeException>(() => paragraphStyle.MinimumWidowLines = -1);
        Assert.Throws<ArgumentException>(() => tableStyle.MinimumBodyRowsOnFirstPage = -1);
        Assert.Throws<ArgumentException>(() => tableStyle.MinimumBodyRowsOnLastPage = -1);
    }
}

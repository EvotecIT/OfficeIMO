using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using OfficeIMO.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfDocumentDeferredTableTests {
    [Fact]
    public void TableDeferred_DefersReplayableFactoryAndKeepsLogicalEdgesSingle() {
        int factoryCalls = 0;
        var style = TableStyles.Minimal();
        style.HeaderRowCount = 1;
        style.FooterRowCount = 1;
        style.Caption = "Deferred caption";
        style.SpacingAfter = 8;

        PdfDocument document = PdfDocument.Create(new PdfOptions {
                PageWidth = 300,
                PageHeight = 500,
                MarginLeft = 24,
                MarginRight = 24,
                MarginTop = 24,
                MarginBottom = 24,
                DefaultFontSize = 9
            })
            .TableDeferred(CreateRows, batchSize: 2, style: style)
            .Paragraph(paragraph => paragraph.Text("After table"));

        Assert.Equal(0, factoryCalls);

        byte[] bytes = document.ToBytes();

        Assert.Equal(2, factoryCalls);
        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        string text = string.Join("\n", pdf.GetPages().Select(page => page.Text));
        Assert.Equal(1, CountOccurrences(text, "Deferred caption"));
        Assert.Equal(1, CountOccurrences(text, "Column A"));
        Assert.Equal(1, CountOccurrences(text, "Grand total"));
        Assert.Equal(1, CountOccurrences(text, "After table"));
        for (int rowIndex = 1; rowIndex <= 7; rowIndex++) {
            Assert.Equal(1, CountOccurrences(text, "Body " + rowIndex));
        }

        IEnumerable<string[]> CreateRows() {
            factoryCalls++;
            yield return new[] { "Column A", "Column B" };
            for (int rowIndex = 1; rowIndex <= 7; rowIndex++) {
                yield return new[] { "Body " + rowIndex, "Value " + rowIndex };
            }

            yield return new[] { "Grand total", "7" };
        }
    }

    [Fact]
    public void TableDeferred_RepeatsHeaderOnNewPagesWithoutRepeatingAtBatchBoundary() {
        var style = TableStyles.Minimal();
        style.HeaderRowCount = 1;
        style.RepeatHeaderRowCount = 1;

        byte[] bytes = PdfDocument.Create(new PdfOptions {
                PageWidth = 260,
                PageHeight = 150,
                MarginLeft = 20,
                MarginRight = 20,
                MarginTop = 20,
                MarginBottom = 20,
                DefaultFontSize = 10
            })
            .TableDeferred(
                () => CreateManyRows(18),
                batchSize: 3,
                style: style)
            .ToBytes();

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        Assert.True(pdf.NumberOfPages > 1);
        foreach (var page in pdf.GetPages()) {
            Assert.True(CountOccurrences(page.Text, "Stable header") == 1, "Expected one header on page " + page.Number + ", but found: " + page.Text);
        }

        string allText = string.Join("\n", pdf.GetPages().Select(page => page.Text));
        for (int rowIndex = 1; rowIndex <= 18; rowIndex++) {
            Assert.Contains("Record " + rowIndex, allText, StringComparison.Ordinal);
        }
    }

    [Fact]
    public void TableDeferred_RejectsGlobalAutoFitAndInvalidBatchSize() {
        var autoFit = TableStyles.Minimal();
        autoFit.AutoFitColumns = true;

        Assert.Throws<ArgumentOutOfRangeException>(() =>
            PdfDocument.Create().TableDeferred(() => CreateManyRows(2), batchSize: 0));

        PdfDocument document = PdfDocument.Create()
            .TableDeferred(() => CreateManyRows(2), style: autoFit);

        ArgumentException exception = Assert.Throws<ArgumentException>(() => document.ToBytes());
        Assert.Contains("automatic column fitting", exception.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void TableDeferred_RejectsRoleCountsBeyondReplayableRows() {
        var style = TableStyles.Minimal();
        style.HeaderRowCount = 1;
        style.FooterRowCount = 2;
        PdfDocument document = PdfDocument.Create()
            .TableDeferred(() => new[] {
                new[] { "Header" },
                new[] { "Only body" }
            }, style: style);

        ArgumentException exception = Assert.Throws<ArgumentException>(() => document.ToBytes());
        Assert.Contains("header and footer row counts", exception.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void TableDeferred_RejectsColumnGrowthInLaterBatch() {
        var style = TableStyles.Minimal();
        style.HeaderRowCount = 0;
        PdfDocument document = PdfDocument.Create()
            .TableDeferred(() => new[] {
                new[] { "First" },
                new[] { "Second" },
                new[] { "Third", "New column" }
            }, batchSize: 2, style: style);

        ArgumentException exception = Assert.Throws<ArgumentException>(() => document.ToBytes());
        Assert.Contains("consistent column count", exception.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void TableDeferred_RichCellsRemainSupportedAcrossBatches() {
        byte[] bytes = PdfDocument.Create()
            .TableDeferred(
                () => new[] {
                    new[] { new PdfTableCell("Header") },
                    new[] { new PdfTableCell(new[] { TextRun.Bolded("Rich body") }) },
                    new[] { new PdfTableCell("Last body") }
                },
                batchSize: 1)
            .ToBytes();

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        string text = string.Join("\n", pdf.GetPages().Select(page => page.Text));
        Assert.Equal(1, CountOccurrences(text, "Rich body"));
        Assert.Equal(1, CountOccurrences(text, "Last body"));
    }

    [Fact]
    public void TableDeferred_ParticipatesInPrecedingKeepWithNextMeasurement() {
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
                SpacingAfter = 70
            })
            .Paragraph(paragraph => paragraph.Text("KeepWithDeferredTable"), style: new PdfParagraphStyle {
                KeepWithNext = true
            })
            .TableDeferred(() => new[] {
                new[] { "DeferredHeader", "Value" },
                new[] { "DeferredBody", "1" }
            }, batchSize: 1)
            .ToBytes();

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        Assert.Equal(2, pdf.NumberOfPages);
        Assert.DoesNotContain("KeepWithDeferredTable", pdf.GetPage(1).Text);
        Assert.Contains("KeepWithDeferredTable", pdf.GetPage(2).Text);
        Assert.Contains("DeferredHeader", pdf.GetPage(2).Text);
    }

    private static IEnumerable<string[]> CreateManyRows(int bodyRowCount) {
        yield return new[] { "Stable header", "Value" };
        for (int rowIndex = 1; rowIndex <= bodyRowCount; rowIndex++) {
            yield return new[] { "Record " + rowIndex + " ", "Value " + rowIndex };
        }
    }

    private static int CountOccurrences(string value, string text) {
        int count = 0;
        int index = 0;
        while ((index = value.IndexOf(text, index, StringComparison.Ordinal)) >= 0) {
            count++;
            index += text.Length;
        }

        return count;
    }
}

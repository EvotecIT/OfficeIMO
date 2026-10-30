using System;
using System.Linq;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfBookmarkTableOfContentsTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void TableOfContents_UnequalColumnContinuationUsesEachActualRightEdge(bool moveBeforeFirstEntry) {
        var entries = Enumerable.Range(0, 6)
            .Select(index => new PdfTableOfContentsEntry("Item " + (char)('A' + index), "details")).ToArray();
        byte[] bytes = PdfDocument.Create(new PdfCreateOptions {
            PageWidth = global::OfficeIMO.OfficeLength.Points(460), PageHeight = global::OfficeIMO.OfficeLength.Points(200), MarginLeft = global::OfficeIMO.OfficeLength.Points(20), MarginRight = global::OfficeIMO.OfficeLength.Points(20), MarginTop = global::OfficeIMO.OfficeLength.Points(20), MarginBottom = global::OfficeIMO.OfficeLength.Points(20)
        }).Columns(column => {
            column.Spacer(global::OfficeIMO.OfficeLength.Points(moveBeforeFirstEntry ? 145 : 60));
            column.TableOfContents(new PdfTableOfContentsOptions { Entries = entries });
        }, new PdfMultiColumnOptions {
            BalanceLastPage = false,
            ColumnDefinitions = new[] {
                new PdfFlowColumn(PdfColumnWidth.Fixed(100), 20),
                new PdfFlowColumn(PdfColumnWidth.Fixed(300), 0)
            }
        }).PageBreak().Bookmark("details").H1("Destination body").ToBytes();
        using var pdf = UglyToad.PdfPig.PdfDocument.Open(bytes);
        Assert.Equal(2, pdf.NumberOfPages);
        double[] rightEdges = pdf.GetPage(1).Letters.Where(letter => letter.Value == "2")
            .Select(letter => letter.EndBaseLine.X).ToArray();
        Assert.Equal(entries.Length, rightEdges.Length);
        Assert.Contains(rightEdges, edge => Math.Abs(edge - 439D) < 0.05D);
        Assert.All(rightEdges, edge => Assert.InRange(Math.Min(Math.Abs(edge - 119D), Math.Abs(edge - 439D)), 0D, 0.05D));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void TableOfContents_SequentialColumnsKeepFinalPageLinksAndBalancedEntries(bool balance) {
        var entries = Enumerable.Range(0, 16)
            .Select(index => new PdfTableOfContentsEntry("Column entry " + index, "details")).ToArray();
        byte[] bytes = PdfDocument.Create(new PdfCreateOptions { PageHeight = global::OfficeIMO.OfficeLength.Points(400), MarginTop = global::OfficeIMO.OfficeLength.Points(30), MarginBottom = global::OfficeIMO.OfficeLength.Points(30) })
            .Columns(column => column.TableOfContents(new PdfTableOfContentsOptions {
                Entries = entries, PageNumberFormatter = page => "page-" + page
            }), new PdfMultiColumnOptions { BalanceLastPage = balance })
            .PageBreak().Bookmark("details").H1("Destination body").ToBytes();
        var loaded = PdfDocument.Load(bytes);
        var pages = loaded.Reader.TextByPage();
        int destinationPage = Enumerable.Range(0, pages.Count).Single(index => pages[index].Contains("Destination body")) + 1;
        foreach (var entry in entries) { Assert.Contains(entry.Title, pages[0]); }
        Assert.Contains("page-" + destinationPage, pages[0]);
        Assert.NotEmpty(loaded.Read().GetLinksByDestinationName("details"));
        using var pdf = UglyToad.PdfPig.PdfDocument.Open(bytes);
        if (balance) {
            Assert.Contains(pdf.GetPage(1).Letters, letter => letter.StartBaseLine.X > 300);
        }
    }

    [Fact]
    public void TableOfContents_RowColumnsRetainTheirPageDependentContentRestriction() {
        Assert.Throws<NotSupportedException>(() => PdfDocument.Create()
            .Row(row => row.Column(PdfColumnWidth.Relative(), column => column.TableOfContents())));
    }

    [Fact]
    public void TableOfContents_AuthoredBookmarks_UsesFinalPagesAndSnapshotsEntries() {
        var entries = new[] {
            new PdfTableOfContentsEntry("Details entry", "details", 2),
            new PdfTableOfContentsEntry("Introduction entry", "intro"),
            new PdfTableOfContentsEntry("Excluded entry", "excluded", 3)
        };
        PdfDocument document = PdfDocument.Create()
            .TableOfContents(new PdfTableOfContentsOptions {
                Entries = entries, MaximumLevel = 2, PageNumberFormatter = page => "page-" + page
            });
        entries[0] = new PdfTableOfContentsEntry("Changed after composition", "changed");
        document.PageBreak().Bookmark("intro").H1("Introduction body");
        for (int index = 0; index < 100; index++) document.Paragraph(paragraph => paragraph.Text("Content overflowing the initial page: " + index));
        document.PageBreak().Bookmark("details").H1("Details body");
        byte[] bytes = document.ToBytes();
        var pages = PdfDocument.Load(bytes).Reader.TextByPage();
        int detailPage = Enumerable.Range(0, pages.Count).Single(index => pages[index].Contains("Details body")) + 1;
        Assert.True(detailPage > 3);
        Assert.Contains("Details entry", pages[0]);
        Assert.Contains("page-" + detailPage, pages[0]);
        Assert.Contains("page-2", pages[0]);
        Assert.DoesNotContain("Excluded entry", pages[0]);
        Assert.DoesNotContain("Changed after composition", pages[0]);
        Assert.True(pages[0].IndexOf("Details entry", StringComparison.Ordinal) < pages[0].IndexOf("Introduction entry", StringComparison.Ordinal));
        Assert.NotEmpty(PdfDocument.Load(bytes).Read().GetLinksByDestinationName("details"));
    }
}

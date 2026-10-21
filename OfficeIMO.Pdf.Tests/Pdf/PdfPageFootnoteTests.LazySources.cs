using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using OfficeIMO.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfPageFootnoteTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void LazyFootnoteDeferredTableRejectsAssociationsInLaterBatchesUntilQualified(bool foreignOwner) {
        PdfDocument document = PdfDocument.Create(Options());
        PdfDocument other = PdfDocument.Create(Options());
        PdfFootnoteDefinition note = (foreignOwner ? other : document).Content.DefineFootnote(
            body => body.Text("LazyAssociatedNote", style: Style()));
        bool materialized = false;
        document.Content.TableDeferred(() => {
            materialized = true;
            return new[] {
                new[] { new PdfTableCell("PlainFirstBatch") },
                new[] { new PdfTableCell(new[] { PdfTextRun.Normal("LazyReference").WithFootnoteReference(note) }) }
            };
        }, batchSize: 1, style: LazyFootnoteTableStyle());
        Assert.False(materialized);

        Assert.Throws<NotSupportedException>(() => document.ToBytes());

        Assert.True(materialized);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void LazyFootnoteReferencesRemainQualifiedOnlyInStaticFlow(bool replayable) {
        PdfDocument document = PdfDocument.Create(Options());
        PdfFootnoteDefinition note = document.Content.DefineFootnote(body => body.Text("StaticQualifiedNote", style: Style()));
        bool materialized = false;
        Action<PdfContentBuilder> build = body => {
            materialized = true;
            body.Paragraph(paragraph => paragraph.Runs(new[] {
                PdfTextRun.Normal("StaticQualifiedReference").WithFootnoteReference(note)
            }), style: Style());
        };
        if (replayable) {
            document.Content.Deferred(_ => build);
            Assert.False(materialized);
            Assert.Throws<NotSupportedException>(() => document.ToBytes());
            Assert.True(materialized);
        } else {
            document.Content.Flow(build);
            using var pdf = PdfPigDocument.Open(document.ToBytes());
            string text = string.Concat(pdf.GetPages().Select(page => page.Text));
            Assert.Single(Regex.Matches(text, "StaticQualifiedReference"));
            Assert.Single(Regex.Matches(text, "StaticQualifiedNote"));
        }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void LazyFootnoteOrdinaryDeferredRowsRetainEagerPaginationAndGlyphPositions(bool noteOutsideTable) {
        using var expected = PdfPigDocument.Open(RenderLazyFootnoteOrdinaryTable(deferred: false, noteOutsideTable));
        using var actual = PdfPigDocument.Open(RenderLazyFootnoteOrdinaryTable(deferred: true, noteOutsideTable));
        Assert.Equal(expected.NumberOfPages, actual.NumberOfPages);
        string[] tokens = Enumerable.Range(0, 12).Select(index => "LazyPlain" + index.ToString("00"))
            .Concat(new[] { "OrdinaryTableAfter", noteOutsideTable ? "OutsideReference" : "OrdinaryTableBefore" })
            .Concat(noteOutsideTable ? new[] { "OutsideNote" } : Array.Empty<string>()).ToArray();
        string actualText = string.Concat(actual.GetPages().Select(page => page.Text));
        foreach (string token in tokens) {
            Assert.Single(Regex.Matches(actualText, Regex.Escape(token)));
            var expectedPage = Assert.Single(expected.GetPages(), page => page.GetWords().Any(word => word.Text == token));
            var actualPage = Assert.Single(actual.GetPages(), page => page.GetWords().Any(word => word.Text == token));
            Assert.Equal(expectedPage.Number, actualPage.Number);
            var expectedPosition = Assert.Single(expectedPage.GetWords(), word => word.Text == token).Letters[0].StartBaseLine;
            var actualPosition = Assert.Single(actualPage.GetWords(), word => word.Text == token).Letters[0].StartBaseLine;
            Assert.Equal(expectedPosition.X, actualPosition.X, 3);
            Assert.Equal(expectedPosition.Y, actualPosition.Y, 3);
        }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void LazyFootnoteDeferredTableCannotBypassTheStaticBodyProfile(bool floating) {
        PdfDocument document = PdfDocument.Create(Options());
        AddLazyFootnoteStaticReference(document);
        PdfTableStyle style = LazyFootnoteTableStyle();
        if (floating) style.Position = new PdfTablePosition();
        else style.ConsumesVerticalFlow = false;
        document.Content.TableDeferred(() => new[] { new[] { new PdfTableCell("UnsupportedDeferredBody") } },
            batchSize: 1, style: style);

        Assert.Throws<NotSupportedException>(() => document.ToBytes());
    }

    [Fact]
    public void LazyFootnoteReplayableBodyCannotIntroduceColumnsBesideAStaticNote() {
        PdfDocument document = PdfDocument.Create(Options());
        AddLazyFootnoteStaticReference(document);
        bool materialized = false;
        document.Content.Deferred(_ => body => {
            materialized = true;
            body.Columns(column => column.Text("UnsupportedDynamicBody", style: Style()), new PdfMultiColumnOptions {
                ColumnCount = 2, BalanceLastPage = false
            });
        });
        Assert.False(materialized);

        Assert.Throws<NotSupportedException>(() => document.ToBytes());

        Assert.True(materialized);
    }

    private static byte[] RenderLazyFootnoteOrdinaryTable(bool deferred, bool noteOutsideTable) {
        PdfDocument document = PdfDocument.Create(Options());
        if (noteOutsideTable) AddLazyFootnoteStaticReference(document);
        else document.Content.Text("OrdinaryTableBefore", style: Style());
        PdfTableCell[][] rows = Enumerable.Range(0, 12).Select(index =>
            new[] { new PdfTableCell("LazyPlain" + index.ToString("00")) }).ToArray();
        if (deferred) document.Content.TableDeferred(() => rows, batchSize: 1, style: LazyFootnoteTableStyle());
        else document.Content.Table(rows, style: LazyFootnoteTableStyle());
        document.Content.Text("OrdinaryTableAfter", style: Style());
        return document.ToBytes();
    }

    private static void AddLazyFootnoteStaticReference(PdfDocument document) {
        PdfFootnoteDefinition note = document.Content.DefineFootnote(body => body.Text("OutsideNote", style: Style()));
        document.Content.Paragraph(paragraph => paragraph.Runs(new[] {
            PdfTextRun.Normal("OutsideReference").WithFootnoteReference(note)
        }), style: Style());
    }

    private static PdfTableStyle LazyFootnoteTableStyle() {
        PdfTableStyle style = FootnoteTableStyle();
        style.ColumnWidthPoints = new List<double?> { 220D };
        style.MinimumBodyRowsOnFirstPage = 0;
        style.MinimumBodyRowsOnLastPage = 0;
        style.AutoFitColumns = false;
        return style;
    }
}

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeIMO.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfPageFootnoteTests {
    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    public void NoteTablePreservesCellGeometryAndExplicitHeaderRepetitionAcrossPages(int repeatHeaders) {
        var document = PdfDocument.Create(Options());
        PdfTableStyle tableStyle = NoteGridStyle();
        tableStyle.HeaderRowCount = 1;
        tableStyle.RepeatHeaderRowCount = repeatHeaders;
        PdfTableCell[][] rows = new[] { new[] { new PdfTableCell("NoteHeader"), new PdfTableCell("RightHeader") } }
            .Concat(Enumerable.Range(0, 18).Select(index => new[] {
                new PdfTableCell("Left" + index.ToString("00")), new PdfTableCell("Right" + index.ToString("00"))
            })).ToArray();
        PdfFootnoteDefinition note = document.Content.DefineFootnote(body => {
            body.Text("BeforeTable", style: Style());
            body.Table(rows, style: tableStyle);
            body.Text("AfterTable", style: Style());
        }, new PdfFootnoteOptions {
            ContinuationNotice = body => body.Text("TableContinues", style: Style())
        });
        document.Content.Paragraph(paragraph => paragraph.Runs(new[] {
            PdfTextRun.Normal("Reference").WithFootnoteReference(note)
        }), style: Style());
        document.Content.Text("FollowingBody", style: Style());

        foreach (byte[] bytes in new[] { document.ToBytes(), document.ToBytes() }) {
            using var pdf = PdfPigDocument.Open(bytes);
            Assert.True(pdf.NumberOfPages > 1);
            var pages = pdf.GetPages().ToArray();
            Assert.Equal(repeatHeaders == 0 ? 1 : pages.Length,
                pages.SelectMany(page => page.GetWords()).Count(word => word.Text == "NoteHeader"));
            Assert.Contains("BeforeTable", pages[0].Text, StringComparison.Ordinal);
            Assert.Contains("AfterTable", pages[pages.Length - 1].Text, StringComparison.Ordinal);
            var followingBody = Assert.Single(pages.SelectMany(page => page.GetWords()), word => word.Text == "FollowingBody");
            var bodyPage = Assert.Single(pages, page => page.GetWords().Any(word => word.Text == "FollowingBody"));
            var noteWordsOnBodyPage = bodyPage.GetWords().Where(word =>
                word.Text.StartsWith("Left", StringComparison.Ordinal) || word.Text.StartsWith("Right", StringComparison.Ordinal) ||
                word.Text is "BeforeTable" or "AfterTable" or "NoteHeader" or "TableContinues").ToArray();
            Assert.NotEmpty(noteWordsOnBodyPage);
            Assert.All(noteWordsOnBodyPage, word => Assert.True(
                followingBody.BoundingBox.Bottom >= word.BoundingBox.Top - .001D,
                "Body text must remain above the note's painted content on their shared page."));
            Assert.DoesNotContain("TableContinues", pages[pages.Length - 1].Text, StringComparison.Ordinal);
            Assert.All(pages.Take(pages.Length - 1), page => Assert.Contains("TableContinues", page.Text, StringComparison.Ordinal));
            for (int index = 0; index < 18; index++) {
                string leftToken = "Left" + index.ToString("00"), rightToken = "Right" + index.ToString("00");
                var page = Assert.Single(pages, candidate => candidate.GetWords().Any(word => word.Text == leftToken));
                var left = Assert.Single(page.GetWords(), word => word.Text == leftToken);
                var right = Assert.Single(page.GetWords(), word => word.Text == rightToken);
                Assert.Equal(left.BoundingBox.Bottom, right.BoundingBox.Bottom, 3);
                Assert.True(right.BoundingBox.Left - left.BoundingBox.Left > 80D);
            }
            Assert.All(pages.SelectMany(page => page.Letters), letter => Assert.True(letter.StartBaseLine.Y >= 19D));
        }
        Assert.Equal(repeatHeaders, tableStyle.RepeatHeaderRowCount);
    }

    [Fact]
    public void SplitNoteTableRowKeepsInlineImageLinkAndDestinationExactlyOnceDuringAdmissionAndRepeatedSerialization() {
        PdfCreateOptions options = Options();
        options.CompressContentStreams = false;
        var document = PdfDocument.Create(options).TaggedPdfCatalogMarkers();
        var runs = new List<PdfTextRun>();
        for (int line = 0; line < 24; line++) {
            if (line == 11) {
                runs.Add(PdfTextRun.Link("LinkedCell", "https://example.com/note"));
                runs.Add(PdfTextRun.Inline(new PdfInlineImage(PdfPngTestImages.CreateRgbPng(0, 180, 80), 8, 8)));
            } else runs.Add(PdfTextRun.Normal("CellLine" + line.ToString("00")));
            if (line < 23) runs.Add(PdfTextRun.Normal("\n"));
        }
        PdfFootnoteDefinition note = document.Content.DefineFootnote(body => body.Table(new[] { new[] {
            new PdfTableCell(runs, namedDestinationName: "NoteCell"), new PdfTableCell("ShortNeighbor")
        } }, style: NoteGridStyle()));
        document.Content.Text("PreviousBody", style: Style()).Spacer(global::OfficeIMO.OfficeLength.Points(60));
        document.Content.Paragraph(paragraph => paragraph.Runs(new[] {
            PdfTextRun.Normal("Reference").WithFootnoteReference(note)
        }), style: Style());

        foreach (byte[] bytes in new[] { document.ToBytes(), document.ToBytes() }) {
            using var pdf = PdfPigDocument.Open(bytes);
            Assert.True(pdf.NumberOfPages > 2);
            var pages = pdf.GetPages().ToArray();
            for (int line = 0; line < 24; line++) {
                if (line == 11) continue;
                string token = "CellLine" + line.ToString("00");
                Assert.Single(pages.SelectMany(page => page.GetWords()), word => word.Text == token);
            }
            Assert.Single(pages.SelectMany(page => page.GetWords()), word => word.Text == "LinkedCell");
            Assert.Single(pages.SelectMany(page => page.GetWords()), word => word.Text == "ShortNeighbor");
            Assert.Single(pages.SelectMany(page => page.GetHyperlinks()));
            Assert.Single(PdfImageExtractor.ExtractImagePlacements(bytes));
            Assert.Single(PdfDocumentReadResult.Load(bytes).NamedDestinations, destination => destination.Name == "NoteCell");
            string syntax = Encoding.ASCII.GetString(bytes);
            Assert.Contains("/Type /StructElem /S /Note", syntax, StringComparison.Ordinal);
            Assert.Contains("/Type /StructElem /S /Table", syntax, StringComparison.Ordinal);
            Assert.Contains("/Type /StructElem /S /TR", syntax, StringComparison.Ordinal);
            Assert.Contains("/Type /StructElem /S /TD", syntax, StringComparison.Ordinal);
        }
    }

    [Fact]
    public void NoteTableRowSpanResumesMergedTextAndPreservesColumnSpanGeometry() {
        var document = PdfDocument.Create(Options());
        PdfTableCell[][] rows = {
            new[] {
                PdfTableCell.Merge(string.Join("\n", Enumerable.Range(0, 26).Select(index => "Merged" + index.ToString("00"))), rowSpan: 3),
                PdfTableCell.Span("WideTop", 2)
            },
            new[] { new PdfTableCell("MiddleLeft"), new PdfTableCell("MiddleRight") },
            new[] { new PdfTableCell("BottomLeft"), new PdfTableCell("BottomRight") },
            new[] { PdfTableCell.Span("FinalWide", 3) }
        };
        PdfFootnoteDefinition note = document.Content.DefineFootnote(body => {
            body.Table(rows, style: NoteGridStyle());
            body.Text("NoteAfterMergedTable", style: Style());
        });
        document.Content.Paragraph(paragraph => paragraph.Runs(new[] {
            PdfTextRun.Normal("Reference").WithFootnoteReference(note)
        }), style: Style());

        foreach (byte[] bytes in new[] { document.ToBytes(), document.ToBytes() }) {
            using var pdf = PdfPigDocument.Open(bytes);
            Assert.True(pdf.NumberOfPages > 2);
            var words = pdf.GetPages().SelectMany(page => page.GetWords()).ToArray();
            for (int index = 0; index < 26; index++)
                Assert.Single(words, word => word.Text == "Merged" + index.ToString("00"));
            foreach (string token in new[] { "WideTop", "MiddleLeft", "MiddleRight", "BottomLeft", "BottomRight", "FinalWide", "NoteAfterMergedTable" })
                Assert.Single(words, word => word.Text == token);
            Assert.Contains("NoteAfterMergedTable", pdf.GetPage(pdf.NumberOfPages).Text, StringComparison.Ordinal);
            Assert.All(pdf.GetPages().SelectMany(page => page.Letters), letter => Assert.True(letter.StartBaseLine.Y >= 19D));
        }
    }

    [Fact]
    public void NoteTableRejectsRecursiveReferencesAndImpossibleUnbreakableFirstRow() {
        var document = PdfDocument.Create(Options());
        var otherNote = document.Content.DefineFootnote(body => body.Text("OtherNote", style: Style()));
        Assert.Throws<NotSupportedException>(() => document.Content.DefineFootnote(body => body.Table(new[] { new[] {
            new PdfTableCell(new[] { PdfTextRun.Normal("Recursive").WithFootnoteReference(otherNote) })
        } }, style: NoteGridStyle())));

        PdfTableStyle style = NoteGridStyle();
        style.FixedRowHeights = new List<global::OfficeIMO.OfficeLength?> { global::OfficeIMO.OfficeLength.Points(100D) };
        style.RowAllowBreakAcrossPages = new List<bool?> { false };
        PdfFootnoteDefinition note = document.Content.DefineFootnote(body => body.Table(new[] {
            new[] { "CannotFit" }
        }, style: style));
        document.Content.Paragraph(paragraph => paragraph.Runs(new[] {
            PdfTextRun.Normal("Reference").WithFootnoteReference(note)
        }), style: Style());
        Assert.Throws<ArgumentException>(() => document.ToBytes());
    }

    [Fact]
    public void ContinuedNoteTableRetainsPanelInsetsAndKeptHeadingWithItsFirstLegalRowFragment() {
        PdfCreateOptions options = Options();
        const double horizontalPadding = 8D, verticalPadding = 3D;
        var document = PdfDocument.Create(options);
        var heading = Style();
        heading.KeepWithNext = true;
        PdfFootnoteDefinition note = document.Content.DefineFootnote(body => body.Panel(panel => {
            panel.Text("TableHeading", style: heading);
            panel.Table(new[] { new[] {
                new PdfTableCell(string.Join("\n", Enumerable.Range(0, 22).Select(index => "PanelCell" + index.ToString("00"))))
            } }, style: NoteGridStyle());
            panel.Text("PanelAfterTable", style: Style());
        }, new PdfPanelStyle {
            PaddingX = horizontalPadding, PaddingY = verticalPadding, SpacingBefore = 0, SpacingAfter = 0,
            BorderColor = global::OfficeIMO.OfficeColor.FromNormalizedRgb(.2D, .2D, .2D), BorderWidth = .5D
        }));
        document.Content.Text("PreviousBody", style: Style()).Spacer(global::OfficeIMO.OfficeLength.Points(60));
        document.Content.Paragraph(paragraph => paragraph.Runs(new[] {
            PdfTextRun.Normal("Reference").WithFootnoteReference(note)
        }), style: Style());

        using var pdf = PdfPigDocument.Open(document.ToBytes());
        Assert.DoesNotContain("Reference", pdf.GetPage(1).Text, StringComparison.Ordinal);
        var first = Assert.Single(pdf.GetPages(), page => page.Text.Contains("TableHeading"));
        Assert.Contains("PanelCell00", first.Text, StringComparison.Ordinal);
        for (int index = 0; index < 22; index++)
            Assert.Single(pdf.GetPages().SelectMany(page => page.GetWords()), word => word.Text == "PanelCell" + index.ToString("00"));
        Assert.Contains("PanelAfterTable", pdf.GetPage(pdf.NumberOfPages).Text, StringComparison.Ordinal);
        Assert.All(pdf.GetPages().SelectMany(page => page.Letters), letter => Assert.True(letter.StartBaseLine.Y >= options.MarginBottom.ToPoints()));
        Assert.All(pdf.GetPages().SelectMany(page => page.GetWords()).Where(word => word.Text.StartsWith("PanelCell", StringComparison.Ordinal)),
            word => {
                Assert.True(word.BoundingBox.Left >= options.MarginLeft.ToPoints() + horizontalPadding);
                Assert.True(word.BoundingBox.Bottom >= options.MarginBottom.ToPoints() + verticalPadding - .001D);
            });
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void MergedNoteTableCellsRejectInlineImagesInRunsAndResolvedParagraphs(bool inResolvedParagraph) {
        var document = PdfDocument.Create(Options());
        PdfTextRun[] imageRuns = { PdfTextRun.Inline(new PdfInlineImage(PdfPngTestImages.CreateRgbPng(0, 180, 80), 8, 8)) };
        PdfTextRun[] cellRuns = inResolvedParagraph ? new[] { PdfTextRun.Normal("CellLabel") } : imageRuns;
        PdfTableCellParagraph[]? paragraphs = inResolvedParagraph ? new[] { new PdfTableCellParagraph(imageRuns) } : null;
        var merged = new PdfTableCell(cellRuns, paragraphs, rowSpan: 2);

        Assert.Throws<NotSupportedException>(() => document.Content.DefineFootnote(body => body.Table(new[] {
            new[] { merged, new PdfTableCell("FirstNeighbor") },
            new[] { new PdfTableCell("SecondNeighbor") }
        }, style: NoteGridStyle())));
    }

    [Theory]
    [InlineData(false, null, false)]
    [InlineData(false, true, true)]
    [InlineData(true, false, false)]
    public void MergedNoteTableAdmissionUsesEffectiveRowSplitOverrides(bool allowRowBreaks, bool? rowOverride, bool supported) {
        var document = PdfDocument.Create(Options());
        PdfTableStyle style = NoteGridStyle();
        style.AllowRowBreakAcrossPages = allowRowBreaks;
        if (rowOverride.HasValue) style.RowAllowBreakAcrossPages = new List<bool?> { rowOverride, rowOverride };
        PdfTableCell[][] rows = {
            new[] {
                PdfTableCell.Merge(string.Join("\n", Enumerable.Range(0, 24).Select(index => "OverrideCell" + index.ToString("00"))), rowSpan: 2),
                new PdfTableCell("FirstNeighbor")
            },
            new[] { new PdfTableCell("SecondNeighbor") }
        };
        if (!supported) {
            Assert.Throws<NotSupportedException>(() => document.Content.DefineFootnote(body => body.Table(rows, style: style)));
            return;
        }
        PdfFootnoteDefinition note = document.Content.DefineFootnote(body => body.Table(rows, style: style));
        document.Content.Paragraph(paragraph => paragraph.Runs(new[] {
            PdfTextRun.Normal("Reference").WithFootnoteReference(note)
        }), style: Style());

        using var pdf = PdfPigDocument.Open(document.ToBytes());
        Assert.True(pdf.NumberOfPages > 1);
        var words = pdf.GetPages().SelectMany(page => page.GetWords()).ToArray();
        for (int index = 0; index < 24; index++)
            Assert.Single(words, word => word.Text == "OverrideCell" + index.ToString("00"));
        Assert.Single(words, word => word.Text == "FirstNeighbor");
        Assert.Single(words, word => word.Text == "SecondNeighbor");
    }

    private static PdfTableStyle NoteGridStyle() => new() {
        HeaderRowCount = 0, RepeatHeaderRowCount = 0, MinimumBodyRowsOnFirstPage = 0, MinimumBodyRowsOnLastPage = 0,
        FontSize = global::OfficeIMO.OfficeLength.Points(8), LineHeight = 1.25D, CellPaddingX = global::OfficeIMO.OfficeLength.Points(2), CellPaddingY = global::OfficeIMO.OfficeLength.Points(1),
        BorderColor = global::OfficeIMO.OfficeColor.FromNormalizedRgb(.2D, .2D, .2D), BorderWidth = global::OfficeIMO.OfficeLength.Points(.5D), SpacingBefore = global::OfficeIMO.OfficeLength.Points(0), SpacingAfter = global::OfficeIMO.OfficeLength.Points(0)
    };
}

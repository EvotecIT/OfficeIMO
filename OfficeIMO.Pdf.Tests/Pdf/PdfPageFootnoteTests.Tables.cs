using System;
using System.Collections.Generic;
using System.Linq;
using OfficeIMO.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfPageFootnoteTests {
    [Fact]
    public void SplitTableRowReservesTheNoteOnlyOnItsActualReferencePage() {
        var options = Options();
        options.PageHeight = 260;
        var document = PdfDocument.Create(options);
        var note = document.Content.DefineFootnote(body => {
            body.Text("TableNoteFirst", style: Style());
            body.Text("TableNoteLast", style: Style());
        });
        document.Content.Table(new[] { new[] { new PdfTableCell(TableReferenceLines(42, 26, "LateReference", note)) } },
            style: FootnoteTableStyle());

        using var pdf = PdfPigDocument.Open(document.ToBytes());
        Assert.Equal(2, pdf.NumberOfPages);
        Assert.DoesNotContain("LateReference", pdf.GetPage(1).Text);
        Assert.DoesNotContain("TableNoteFirst", pdf.GetPage(1).Text);
        Assert.Contains("LateReference", pdf.GetPage(2).Text);
        Assert.Contains("TableNoteFirst", pdf.GetPage(2).Text);
        Assert.Contains("CellLine42", pdf.GetPage(2).Text);
        double firstBaseline = TableMarkerBaseline(pdf, 2, "TableNoteFirst");
        double lastBaseline = TableMarkerBaseline(pdf, 2, "TableNoteLast");
        Assert.InRange(lastBaseline, 20D, 30D);
        Assert.InRange(firstBaseline - lastBaseline, 9.99D, 10.01D);
        Assert.True(TableMarkerBaseline(pdf, 2, "CellLine42") > firstBaseline,
            "The last admitted table line must remain above the reserved note region.");
    }

    [Fact]
    public void SameLabelTableReferencesKeepDistinctDefinitionsAndSerializationState() {
        var document = PdfDocument.Create(Options());
        var first = document.Content.DefineFootnote(body => body.Text("FirstDefinition", style: Style()));
        var second = document.Content.DefineFootnote(body => body.Text("SecondDefinition", style: Style()));
        document.Content.Table(new[] {
            new[] { new PdfTableCell(new[] { PdfTextRun.Normal("Reference").WithFootnoteReference(first) }) },
            new[] { new PdfTableCell(new[] { PdfTextRun.Normal("Reference").WithFootnoteReference(second) }) }
        }, style: FootnoteTableStyle());

        foreach (byte[] bytes in new[] { document.ToBytes(), document.ToBytes() }) {
            using var pdf = PdfPigDocument.Open(bytes);
            Assert.Equal(1, pdf.NumberOfPages);
            Assert.Single(pdf.GetPage(1).GetWords(), word => word.Text == "FirstDefinition");
            Assert.Single(pdf.GetPage(1).GetWords(), word => word.Text == "SecondDefinition");
            Assert.True(TableMarkerBaseline(pdf, 1, "FirstDefinition") > TableMarkerBaseline(pdf, 1, "SecondDefinition"),
                "Footnotes must retain their table source order even when the reference labels are identical.");
        }
    }

    [Fact]
    public void KeptTableAdmitsBothReferencesBeforeGreedilyExpandingItsFirstNote() {
        var options = Options();
        options.PageHeight = 260;
        var document = PdfDocument.Create(options);
        var first = document.Content.DefineFootnote(body => {
            for (int index = 0; index < 20; index++) body.Text("LongTableNote" + index.ToString("00"), style: Style());
        });
        var second = document.Content.DefineFootnote(body => body.Text("SecondTableNote", style: Style()));
        var style = FootnoteTableStyle();
        style.KeepTogether = true;
        document.Content.Table(new[] {
            new[] { new PdfTableCell(new[] { PdfTextRun.Normal("KeptFirst\n"), PdfTextRun.Normal("FirstReference").WithFootnoteReference(first) }) },
            new[] { new PdfTableCell(new[] { PdfTextRun.Normal("KeptSecond\n"), PdfTextRun.Normal("SecondReference").WithFootnoteReference(second) }) }
        }, style: style);

        using var pdf = PdfPigDocument.Open(document.ToBytes());
        Assert.Equal(2, pdf.NumberOfPages);
        Assert.Contains("FirstReference", pdf.GetPage(1).Text);
        Assert.Contains("SecondReference", pdf.GetPage(1).Text);
        Assert.Contains("LongTableNote00", pdf.GetPage(1).Text);
        Assert.Contains("SecondTableNote", pdf.GetPage(1).Text);
        Assert.DoesNotContain("KeptSecond", pdf.GetPage(2).Text);
        Assert.Contains("LongTableNote19", pdf.GetPage(2).Text);
        Assert.True(TableMarkerBaseline(pdf, 1, "SecondReference") > TableMarkerBaseline(pdf, 1, "LongTableNote00"),
            "Both kept rows must finish before the first note fragment begins.");
    }

    [Fact]
    public void MergedTableCellSchedulesItsLateReferenceFromTheAdmittedSpanFragment() {
        var options = Options();
        options.PageHeight = 260;
        var document = PdfDocument.Create(options);
        var note = document.Content.DefineFootnote(body => {
            body.Text("MergedNoteFirst", style: Style());
            body.Text("MergedNoteLast", style: Style());
        });
        document.Content.Table(new[] {
            new[] { new PdfTableCell(TableReferenceLines(42, 26, "MergedReference", note), rowSpan: 2), new PdfTableCell("NeighbourFirst") },
            new[] { new PdfTableCell("NeighbourLast") }
        }, style: FootnoteTableStyle());

        using var pdf = PdfPigDocument.Open(document.ToBytes());
        var referencePage = Assert.Single(pdf.GetPages(), page => page.Text.Contains("MergedReference"));
        Assert.True(referencePage.Number > 1);
        Assert.Contains("MergedNoteFirst", referencePage.Text);
        Assert.DoesNotContain("MergedNoteFirst", pdf.GetPage(1).Text);
        foreach (int line in Enumerable.Range(1, 42))
            Assert.Single(pdf.GetPages().SelectMany(page => page.GetWords()), word => word.Text == "CellLine" + line.ToString("00"));
        Assert.Single(pdf.GetPages().SelectMany(page => page.GetWords()), word => word.Text == "MergedNoteLast");
        Assert.All(pdf.GetPages().SelectMany(page => page.Letters), letter => Assert.True(letter.StartBaseLine.Y >= 19D));
    }

    [Fact]
    public void ImpossibleTableReferenceAndFirstNoteFragmentThrowBeforeThePageLimit() {
        var options = Options();
        options.MaxGeneratedPages = 4;
        var document = PdfDocument.Create(options);
        var note = document.Content.DefineFootnote(body => body.Text("CannotFit", style: new PdfParagraphStyle {
            FontSize = 8, LineSpacing = PdfLineSpacing.Exactly(110), SpacingBefore = 0, SpacingAfter = 0
        }));
        document.Content.Table(new[] { new[] {
            new PdfTableCell(new[] { PdfTextRun.Normal("Reference").WithFootnoteReference(note) })
        } }, style: FootnoteTableStyle());

        Assert.Throws<ArgumentException>(() => document.ToBytes());
    }

    private static PdfTableStyle FootnoteTableStyle() => new() {
        HeaderRowCount = 0, FontSize = 8, LineHeight = 1.25D,
        CellPaddingX = 0, CellPaddingY = 0, MinRowHeight = 0,
        SpacingBefore = 0, SpacingAfter = 0, BorderWidth = 0, RowSeparatorWidth = 0
    };

    private static IReadOnlyList<PdfTextRun> TableReferenceLines(int count, int referenceLine,
        string referenceLabel, PdfFootnoteDefinition note) {
        var runs = new List<PdfTextRun>();
        for (int line = 1; line <= count; line++) {
            runs.Add(PdfTextRun.Normal("CellLine" + line.ToString("00")));
            if (line == referenceLine) runs.Add(PdfTextRun.Normal(" " + referenceLabel).WithFootnoteReference(note));
            if (line < count) runs.Add(PdfTextRun.Normal("\n"));
        }
        return runs;
    }

    private static double TableMarkerBaseline(PdfPigDocument pdf, int pageNumber, string marker) {
        var letters = pdf.GetPage(pageNumber).Letters;
        int index = string.Concat(letters.Select(letter => letter.Value)).IndexOf(marker, StringComparison.Ordinal);
        Assert.True(index >= 0, "Missing table-footnote marker " + marker);
        return letters[index].StartBaseLine.Y;
    }
}

using System;
using System.IO;
using System.Linq;
using OfficeIMO.Pdf;
using UglyToad.PdfPig.Content;
using Xunit;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;

namespace OfficeIMO.Tests.Pdf;

public class PdfThemeCompositionDefaultsTests {
    private static PdfCreateOptions SmallPage() => new PdfCreateOptions {
        PageWidth = OfficeLength.Points(420),
        PageHeight = OfficeLength.Points(600),
        MarginLeft = OfficeLength.Points(40),
        MarginRight = OfficeLength.Points(40),
        MarginTop = OfficeLength.Points(40),
        MarginBottom = OfficeLength.Points(40)
    };

    private static Letter[] Letters(byte[] bytes, int page = 1) {
        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        return pdf.GetPage(page).Letters.Where(letter => !string.IsNullOrWhiteSpace(letter.Value)).ToArray();
    }

    private static Letter[] Word(Letter[] letters, string word) {
        string text = string.Concat(letters.Select(letter => letter.Value));
        int index = text.IndexOf(word, StringComparison.Ordinal);
        Assert.True(index >= 0, "Expected to find '" + word + "' in '" + text + "'.");
        return letters.Skip(index).Take(word.Length).ToArray();
    }

    [Fact]
    public void Headings_UseThemeTypeScaleAndTextColourByDefault() {
        byte[] bytes = PdfDocument.Create(SmallPage())
            .H1("Alpha").H2("Bravo")
            .ToBytes();

        Letter[] letters = Letters(bytes);
        Assert.InRange(Word(letters, "Alpha")[0].PointSize, OfficeTheme.Default.TypeScale.Heading1.ToPoints() - 0.2, OfficeTheme.Default.TypeScale.Heading1.ToPoints() + 0.2);
        Assert.InRange(Word(letters, "Bravo")[0].PointSize, OfficeTheme.Default.TypeScale.Heading2.ToPoints() - 0.2, OfficeTheme.Default.TypeScale.Heading2.ToPoints() + 0.2);
    }

    [Fact]
    public void Heading_KeepsThemeSpaceBeforeAfterAnExplicitlyStyledTable() {
        byte[] bytes = PdfDocument.Create(SmallPage())
            .Table(new[] { new[] { "Head" }, new[] { "Lastrow" } }, style: new PdfTableStyle { CellPaddingY = OfficeLength.Points(2) })
            .H2("Following")
            .ToBytes();

        Letter[] letters = Letters(bytes);
        double rowBottom = Word(letters, "Lastrow").Min(letter => letter.GlyphRectangle.Bottom);
        double headingTop = Word(letters, "Following").Max(letter => letter.GlyphRectangle.Top);
        Assert.True(rowBottom - headingTop >= 12, "Gap was " + (rowBottom - headingTop) + " points.");
    }

    [Fact]
    public void DefaultTable_FitsColumnsToContentAndRightAlignsNumbers() {
        byte[] bytes = PdfDocument.Create(SmallPage())
            .Table(new[] {
                new[] { "Item", "Qty", "Amount" },
                new[] { "USB-C docking station with charger", "8", "5,147.50 EUR" },
                new[] { "Monitor", "12", "312.00 EUR" }
            })
            .ToBytes();

        Letter[] letters = Letters(bytes);
        double itemX = Word(letters, "Item")[0].StartBaseLine.X;
        double qtyX = Word(letters, "Qty")[0].StartBaseLine.X;
        // Equal columns would give the item column a third of the 340 point frame.
        Assert.True(qtyX - itemX > 150, "Item column is " + (qtyX - itemX) + " points wide.");
        double bigAmountEnd = Word(letters, "5,147.50 EUR".Replace(" ", string.Empty)).Last().EndBaseLine.X;
        double smallAmountEnd = Word(letters, "312.00EUR").Last().EndBaseLine.X;
        Assert.InRange(Math.Abs(bigAmountEnd - smallAmountEnd), 0, 0.5);
    }

    [Fact]
    public void PageNumbers_DefaultToPageNOfM() {
        byte[] bytes = PdfDocument.Create(new PdfCreateOptions { ShowPageNumbers = true })
            .Paragraph(p => p.Text("Body"))
            .PageBreak()
            .Paragraph(p => p.Text("Second"))
            .ToBytes();

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        Assert.Contains("Page 1 of 2", pdf.GetPage(1).Text, StringComparison.Ordinal);
        Assert.Contains("Page 2 of 2", pdf.GetPage(2).Text, StringComparison.Ordinal);
    }

    [Fact]
    public void Title_SitsAboveHeadingsAndStaysOutOfTheOutline() {
        var document = PdfDocument.Create(new PdfCreateOptions { CreateOutlineFromHeadings = true });
        document.Content.Title("Handbook").Section("Before", section => section.Paragraph(p => p.Text("Body")));
        byte[] bytes = document.ToBytes();

        Letter[] letters = Letters(bytes);
        Assert.True(Word(letters, "Handbook")[0].PointSize > Word(letters, "Before")[0].PointSize + 4);
        PdfOutlineItem outline = Assert.Single(PdfInspector.Inspect(bytes).Outlines);
        Assert.Equal("Before", outline.Title);
    }

    [Fact]
    public void HyphenatedIdentifier_MovesToTheNextLineWhole() {
        byte[] bytes = PdfDocument.Create(new PdfCreateOptions {
                PageWidth = OfficeLength.Points(300), PageHeight = OfficeLength.Points(300), MarginLeft = OfficeLength.Points(20), MarginRight = OfficeLength.Points(20), MarginTop = OfficeLength.Points(20), MarginBottom = OfficeLength.Points(20)
            })
            .Paragraph(p => p.FontSize(OfficeLength.Points(18)).Bold("Invoice EVO-DEMO-2026-0917"))
            .ToBytes();

        Letter[] letters = Letters(bytes);
        Letter[] identifier = Word(letters, "EVO-DEMO-2026-0917");
        Assert.Single(identifier.Select(letter => Math.Round(letter.StartBaseLine.Y, 1)).Distinct());
    }

    [Fact]
    public void KeepWithNextTable_CarriesItsFinalRowsWithTheFollowingBlock() {
        var options = SmallPage();
        var rows = Enumerable.Range(0, 28).Select(index => new[] { "Row" + index.ToString("00", System.Globalization.CultureInfo.InvariantCulture), "Value" }).ToList();
        rows.Insert(0, new[] { "Name", "Value" });
        PdfDocument document = PdfDocument.Create(options);
        document.Content
            .Spacer(OfficeLength.Points(150))
            .Table(rows, style: new PdfTableStyle { KeepWithNext = true, CellPaddingY = OfficeLength.Points(2) })
            .PanelParagraph(p => p.Text("Summary line one\nSummary line two\nSummary line three\nSummary line four\nSummary line five\nSummary line six\nSummary line seven"));
        byte[] bytes = document.ToBytes();

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        Assert.Equal(2, pdf.NumberOfPages);
        string first = pdf.GetPage(1).Text;
        string second = pdf.GetPage(2).Text;
        Assert.Contains("Row00", first, StringComparison.Ordinal);
        Assert.Contains("Row27", second, StringComparison.Ordinal);
        Assert.Contains("Summary line one", second, StringComparison.Ordinal);
    }

    [Fact]
    public void ComboBoxAppearance_DrawsADropDownMarker() {
        byte[] bytes = PdfDocument.Create()
            .ChoiceField("choice", new[] { "One", "Two" }, value: "One")
            .TextField("text", value: "Plain")
            .ToBytes();

        Assert.Equal(1, CountOccurrences(PdfTestContentText.ReadDecoded(bytes), " l h f\n"));
    }

    [Fact]
    public void PanelParagraph_DropsTrailingParagraphSpacingInsideThePanel() {
        var style = new PdfPanelStyle { PaddingX = OfficeLength.Points(10), PaddingY = OfficeLength.Points(10), SpacingAfter = OfficeLength.Points(0) };
        byte[] collapsed = PdfDocument.Create(SmallPage())
            .PanelParagraph(p => p.Text("Inside"), style)
            .Paragraph(p => p.Text("After"))
            .ToBytes();
        byte[] explicitZero = PdfDocument.Create(SmallPage())
            .Panel(panel => panel.Paragraph(p => p.Text("Inside"), style: new PdfParagraphStyle { SpacingAfter = OfficeLength.Points(0) }), style)
            .Paragraph(p => p.Text("After"))
            .ToBytes();

        double collapsedAfter = Word(Letters(collapsed), "After")[0].StartBaseLine.Y;
        double explicitAfter = Word(Letters(explicitZero), "After")[0].StartBaseLine.Y;
        Assert.InRange(Math.Abs(collapsedAfter - explicitAfter), 0, 0.01);
    }

    private static int CountOccurrences(string text, string value) {
        int count = 0;
        for (int index = text.IndexOf(value, StringComparison.Ordinal); index >= 0; index = text.IndexOf(value, index + value.Length, StringComparison.Ordinal)) count++;
        return count;
    }
}

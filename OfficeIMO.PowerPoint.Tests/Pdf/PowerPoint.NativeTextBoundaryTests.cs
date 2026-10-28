using System;
using System.IO;
using System.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.Pdf;
using Xunit;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Tests.Pdf;

public sealed class PowerPointNativeTextBoundaryTests {
    [Theory]
    [InlineData(false, 18)]
    [InlineData(true, 18)]
    [InlineData(false, 24)]
    [InlineData(true, 24)]
    public void CondensedTableTextRemainsVisibleWhenTheCellGrowsToFit(bool rich, double height) {
        using var presentation = PowerPointDocument.Create(new MemoryStream());
        var slide = presentation.AddSlide();
        var table = slide.AddTable(1,1,global::OfficeIMO.OfficeLength.Points(20),global::OfficeIMO.OfficeLength.Points(20),global::OfficeIMO.OfficeLength.Points(200),global::OfficeIMO.OfficeLength.Points(height));
        table.SetRowHeights(global::OfficeIMO.OfficeLength.Points(height));
        var cell = table.GetCell(0, 0);
        cell.Text = "gypsy"; cell.FontSize = 24;
        cell.PaddingTopPoints = cell.PaddingBottomPoints = 0;
        cell.Paragraphs[0].LineSpacingMultiplier = .75D;
        cell.Paragraphs[0].Runs[0].Bold = rich;
        var result = presentation.ConvertToPdfWithEvidence();
        var span = Assert.Single(PdfCore.PdfDocument.Load(new System.IO.MemoryStream(result.ToBytes())).Pages[0].GetTextSpans(),
            span => span.Text.Contains("gypsy", StringComparison.Ordinal));
        Assert.True(span.IsVisible);
        Assert.Equal(24D, span.FontSize);
        Assert.DoesNotContain(result.Warnings, warning => warning.Code == "table-cell-overflow");
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public void CondensedTextReportsPaintBeyondItsFrame(bool rich, bool precedingParagraph) {
        using var presentation = PowerPointDocument.Create(new MemoryStream());
        var slide = presentation.AddSlide();
        var box = slide.AddTextBoxPoints(precedingParagraph ? "FIRST" : "gypsy", 20, 20, 200, precedingParagraph ? 36 : 18);
        box.FontSize = 24; box.FontName = "Helvetica";
        box.TextMarginLeftPoints = box.TextMarginRightPoints = box.TextMarginTopPoints = box.TextMarginBottomPoints = 0;
        if (precedingParagraph) box.AddParagraph("gypsy");
        foreach (var paragraph in box.Paragraphs) {
            paragraph.LineSpacingMultiplier = .75D;
            paragraph.Runs[0].Bold = rich;
        }
        var result = presentation.ConvertToPdfWithEvidence();
        var warning = Assert.Single(result.Warnings, warning => warning.Code == "text-box-overflow");
        Assert.Equal(PdfCore.PdfLayoutDiagnosticKind.Overflow, warning.LayoutDiagnostic!.Kind);
        Assert.Equal(OfficeConversionLossKind.None, warning.LossKind);
        Assert.Contains(PdfCore.PdfReadDocument.Open(result.ToBytes()).Pages[0].GetTextSpans(),
            span => span.Text.Contains("gypsy", StringComparison.Ordinal) && span.IsVisible);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void FirstLineTallerThanItsTextFrameRemainsVisibleAndReportsBoundedOverflow(bool rich) {
        using var presentation = PowerPointDocument.Create(new MemoryStream());
        PowerPointTextBox box = presentation.AddSlide().AddTextBoxPoints("TALL", 20, 20, 150, 10);
        box.FontSize = 32;
        box.Paragraphs[0].Runs[0].Underline = rich;
        var result = presentation.ConvertToPdfWithEvidence(new PowerPointPdfExportOptions {
            ResourcePolicy = PdfCore.PdfResourcePolicy.CreatePortableDeterministic()
        });
        var warning = Assert.Single(result.Warnings, warning => warning.Code == "text-box-overflow");
        Assert.Equal(PdfCore.PdfLayoutDiagnosticKind.Overflow, warning.LayoutDiagnostic!.Kind);
        Assert.Equal(OfficeConversionLossKind.None, warning.LossKind);
        Assert.True(warning.LayoutDiagnostic.HasBounds);
        var span = Assert.Single(PdfCore.PdfDocument.Load(new System.IO.MemoryStream(result.ToBytes())).Pages[0].GetTextSpans(),
            span => span.Text.Contains("TALL", StringComparison.Ordinal));
        Assert.True(span.IsVisible);
        Assert.Equal(32D, span.FontSize);
    }

    [Fact]
    public void TableCellGrowthRetainsTheCompleteParagraph() {
        using var presentation = PowerPointDocument.Create(new MemoryStream());
        PowerPointSlide slide = presentation.AddSlide();
        PowerPointTable table = slide.AddTable(1,1,global::OfficeIMO.OfficeLength.Points(20),global::OfficeIMO.OfficeLength.Points(20),global::OfficeIMO.OfficeLength.Points(140),global::OfficeIMO.OfficeLength.Points(40));
        table.SetRowHeights(global::OfficeIMO.OfficeLength.Points(40));
        PowerPointTableCell cell = table.GetCell(0, 0);
        cell.FontSize = 12;
        cell.Text = "VISIBLE alpha beta gamma delta epsilon zeta eta theta iota kappa lambda mu TAIL";
        cell.Paragraphs[0].SetBullet('*');
        var result = presentation.ConvertToPdfWithEvidence();
        var spans = PdfCore.PdfReadDocument.Open(result.ToBytes()).Pages[0].GetTextSpans();
        string text = string.Concat(spans.Select(span => span.Text));
        Assert.Contains("VISIBLE", text);
        Assert.Contains("TAIL", text);
        Assert.DoesNotContain("...", text);
        Assert.All(spans, span => Assert.True(span.IsVisible));
        Assert.DoesNotContain(result.Warnings, warning => warning.Code == "table-cell-overflow");
    }

    [Theory]
    [InlineData("bullet")]
    [InlineData("alignment")]
    [InlineData("spacing")]
    public void OverflowingParagraphsRetainTheirTailAndFollowingParagraphs(string flow) {
        var fonts = new OfficeFontFaceCollection().Add("BoundaryFont",
            OfficeIMO.TestAssets.PdfTestFontAssets.LoadBundledOpenTypeCffFont());
        var profile = new OfficeRenderingProfile("native-overflow", fonts);
        using var presentation = PowerPointDocument.Create(new MemoryStream());
        presentation.SlideSize.SetSize(OfficeIMO.OfficeLength.Points(400),OfficeIMO.OfficeLength.Points(360));
        PowerPointSlide slide = presentation.AddSlide();
        PowerPointTextBox box = slide.AddTextBoxPoints(
            "VISIBLE alpha beta gamma delta epsilon zeta eta theta iota kappa lambda mu nu xi omicron TAIL", 20, 20, 120, 40);
        box.FontName = "BoundaryFont";
        box.FontSize = 12;
        if (flow == "bullet") box.Paragraphs[0].SetBullet('*');
        else {
            PowerPointParagraph next = box.AddParagraph("LATER");
            if (flow == "alignment") next.Alignment = PowerPointTextAlignment.Right;
            else box.Paragraphs[0].SetSpaceAfterPoints(4);
        }
        var result = presentation.ConvertToPdfWithEvidence(new PowerPointPdfExportOptions().UseRenderingProfile(profile));
        using var pdf = UglyToad.PdfPig.PdfDocument.Open(result.ToBytes());
        string pdfText = string.Concat(pdf.GetPage(1).Letters.Select(letter => letter.Value));
        Assert.Contains("VISIBLE", pdfText);
        Assert.Contains("TAIL", pdfText);
        if (flow != "bullet") Assert.Contains("LATER", pdfText);
        Assert.DoesNotContain("...", pdfText);
        var warning = Assert.Single(result.Warnings, warning => warning.Code == "text-box-overflow");
        Assert.Equal(OfficeConversionLossKind.None, warning.LossKind);
        var spans = PdfCore.PdfReadDocument.Open(result.ToBytes()).Pages[0].GetTextSpans();
        var tail = Assert.Single(spans, span => span.Text.Contains("TAIL", StringComparison.Ordinal));
        Assert.True(tail.IsVisible);
        Assert.True(360D - tail.Y > 60D);
        if (flow != "bullet") {
            var later = Assert.Single(spans, span => span.Text.Contains("LATER", StringComparison.Ordinal));
            Assert.True(later.IsVisible);
            Assert.True(later.Y < tail.Y);
        }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void PlainAndRichTextOverflowRetainsTheCompleteText(bool rich) {
        using var presentation = PowerPointDocument.Create(new MemoryStream());
        PowerPointTextBox box = presentation.AddSlide().AddTextBoxPoints("VISIBLE alpha beta gamma delta epsilon TAIL", 20, 20, 90, 23);
        box.FontSize = 12;
        box.Paragraphs[0].Runs[0].Underline = rich;
        var result = presentation.ConvertToPdfWithEvidence(new PowerPointPdfExportOptions { ResourcePolicy = PdfCore.PdfResourcePolicy.CreatePortableDeterministic() });
        using var pdf = UglyToad.PdfPig.PdfDocument.Open(result.ToBytes());
        string text = string.Concat(pdf.GetPage(1).Letters.Select(letter => letter.Value));
        Assert.Contains("VISIBLE", text);
        Assert.Contains("TAIL", text);
        Assert.DoesNotContain("...", text);
        var warning = Assert.Single(result.Warnings, warning => warning.Code == "text-box-overflow");
        Assert.Equal(PdfCore.PdfLayoutDiagnosticKind.Overflow, warning.LayoutDiagnostic!.Kind);
        Assert.Equal(OfficeConversionLossKind.None, warning.LossKind);
        Assert.True(warning.LayoutDiagnostic.HasBounds);
    }

    [Theory]
    [InlineData("plain")]
    [InlineData("rich")]
    [InlineData("paragraph")]
    public void ConfiguredFallbackFontsApplyToUnstyledGlyphsButPreserveAuthoredFonts(string flow) {
        using var presentation = PowerPointDocument.Create(new MemoryStream());
        PowerPointSlide slide = presentation.AddSlide();
        PowerPointTextBox box = slide.AddTextBoxPoints("DEFAULT", 20, 20, 220, 80);
        box.FontSize = 12;
        if (flow == "rich") box.Paragraphs[0].Runs[0].Underline = true;
        if (flow == "paragraph") box.Paragraphs[0].SetBullet('*');
        PowerPointTextBox authored = slide.AddTextBoxPoints("AUTHORED", 20, 120, 220, 50);
        authored.FontName = "Courier";
        authored.FontSize = 12;
        byte[] bytes = presentation.ToPdf(new PowerPointPdfExportOptions {
            FontFamily = "serif", ResourcePolicy = PdfCore.PdfResourcePolicy.CreatePortableDeterministic()
        }).ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true });
        using var pdf = UglyToad.PdfPig.PdfDocument.Open(bytes);
        var letters = pdf.GetPage(1).Letters;
        string text = string.Concat(letters.Select(letter => letter.Value));
        int defaultIndex = text.IndexOf("DEFAULT", StringComparison.Ordinal);
        int authoredIndex = text.IndexOf("AUTHORED", StringComparison.Ordinal);
        Assert.True(defaultIndex >= 0 && authoredIndex >= 0, text);
        Assert.Contains("Times", letters[defaultIndex].FontName, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("Courier", letters[authoredIndex].FontName, StringComparison.OrdinalIgnoreCase);
    }
}

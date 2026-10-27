using System;
using System.IO;
using System.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.Pdf;
using PdfCore = OfficeIMO.Pdf;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace OfficeIMO.Tests.Pdf;

public sealed class PowerPointParagraphFlowContractTests {
    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public void EmptyParagraphOutsideTheFrameRetainsFollowingTextAndDiagnosesTextBoxOverflow(bool table, bool followingText) {
        using var presentation = PowerPointDocument.Create(new MemoryStream());
        presentation.SlideSize.SetSize(OfficeIMO.OfficeLength.Points(300),OfficeIMO.OfficeLength.Points(220));
        var slide = presentation.AddSlide();
        if (table) {
            var cell = slide.AddTablePoints(1, 1, 20, 20, 240, 60).GetCell(0, 0);
            cell.Text = "VISIBLE"; cell.FontName = "Helvetica"; cell.FontSize = 32;
            cell.AddParagraph("");
            if (followingText) cell.AddParagraph("OVERFLOW");
        } else {
            var box = slide.AddTextBoxPoints("VISIBLE", 20, 20, 240, 60);
            box.FontName = "Helvetica"; box.FontSize = 32;
            box.TextMarginLeftPoints = box.TextMarginRightPoints = box.TextMarginTopPoints = box.TextMarginBottomPoints = 0;
            box.AddParagraph("");
            if (followingText) box.AddParagraph("OVERFLOW");
        }
        var result = presentation.ConvertToPdfWithEvidence();
        var page = PdfCore.PdfReadDocument.Open(result.ToBytes()).Pages[0];
        Assert.True(Span(page, "VISIBLE").IsVisible);
        if (followingText) {
            Assert.True(Span(page, "OVERFLOW").IsVisible);
            Assert.True(Span(page, "VISIBLE").Y - Span(page, "OVERFLOW").Y > 60D);
        }
        if (!table) {
            var overflows = result.Warnings.Where(warning => warning.Code == "text-box-overflow").ToArray();
            Assert.Equal(followingText ? 1 : 0, overflows.Length);
            Assert.All(overflows, warning => Assert.Equal(OfficeConversionLossKind.None, warning.LossKind));
        }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void CondensedDescendersRemainVisibleBeyondTheTextBoxBoundary(bool rich) {
        using var presentation = PowerPointDocument.Create(new MemoryStream());
        presentation.SlideSize.SetSize(OfficeIMO.OfficeLength.Points(300),OfficeIMO.OfficeLength.Points(220));
        var box = presentation.AddSlide().AddTextBoxPoints("gypsy\nplaying\nquickly", 20, 20, 240, 40);
        box.FontSize = 24; box.FontName = "Helvetica";
        box.TextMarginLeftPoints = box.TextMarginRightPoints = box.TextMarginTopPoints = box.TextMarginBottomPoints = 0;
        box.Paragraphs[0].LineSpacingMultiplier = 0.75D;
        box.Paragraphs[0].Runs[0].Bold = rich;
        byte[] bytes = presentation.ToPdf(new PowerPointPdfExportOptions {
            ResourcePolicy = OfficeIMO.Pdf.PdfResourcePolicy.CreatePortableDeterministic()
        }).ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true });
        var page = PdfCore.PdfReadDocument.Open(bytes).Pages[0];
        Assert.All(new[] { "gypsy", "playing", "quickly" }, text => Assert.True(Span(page, text).IsVisible));
        Assert.InRange(Span(page, "gypsy").Y - Span(page, "playing").Y, 17.9D, 18.1D);
        Assert.InRange(Span(page, "playing").Y - Span(page, "quickly").Y, 17.9D, 18.1D);
        Assert.True(page.GetPageSize().Height - Span(page, "quickly").Y > 60D);
    }

    [Theory]
    [InlineData(false, false, false)]
    [InlineData(false, false, true)]
    [InlineData(false, true, false)]
    [InlineData(false, true, true)]
    [InlineData(true, false, false)]
    [InlineData(true, false, true)]
    [InlineData(true, true, false)]
    [InlineData(true, true, true)]
    public void CondensedParagraphLeadingSurvivesNativeRendering(bool table, bool rich, bool trailingParagraph) {
        using var presentation = PowerPointDocument.Create(new MemoryStream());
        presentation.SlideSize.SetSize(OfficeIMO.OfficeLength.Points(300),OfficeIMO.OfficeLength.Points(220));
        var slide = presentation.AddSlide();
        PowerPointParagraph paragraph;
        if (table) {
            var cell = slide.AddTablePoints(1, 1, 20, 20, 240, 160).GetCell(0, 0);
            cell.Text = "FIRST\nSECOND\nTHIRD"; cell.FontName = "Helvetica"; cell.FontSize = 12;
            paragraph = cell.Paragraphs[0];
            if (trailingParagraph) cell.AddParagraph("END");
        } else {
            var box = slide.AddTextBoxPoints("FIRST\nSECOND\nTHIRD", 20, 20, 240, 160);
            box.FontName = "Helvetica"; box.FontSize = 12;
            paragraph = box.Paragraphs[0];
            if (trailingParagraph) box.AddParagraph("END").Alignment = PowerPointTextAlignment.Right;
        }
        paragraph.LineSpacingMultiplier = 0.75D;
        if (rich) paragraph.Runs[0].Bold = true;
        var page = Read(presentation).Pages[0];
        Assert.InRange(Span(page, "FIRST").Y - Span(page, "SECOND").Y, 8.9D, 9.1D);
        Assert.InRange(Span(page, "SECOND").Y - Span(page, "THIRD").Y, 8.9D, 9.1D);
        if (trailingParagraph) Assert.True(Span(page, "THIRD").Y > Span(page, "END").Y);
    }

    [Theory]
    [InlineData("paragraph", false)]
    [InlineData("paragraph", true)]
    [InlineData("list", false)]
    [InlineData("list", true)]
    [InlineData("master", false)]
    [InlineData("master", true)]
    [InlineData("end", false)]
    [InlineData("end", true)]
    public void BlankParagraphUsesEffectiveFontSize(string source, bool differentAlignment) {
        using var presentation = PowerPointDocument.Create(new MemoryStream());
        presentation.SlideSize.SetSize(OfficeIMO.OfficeLength.Points(300),OfficeIMO.OfficeLength.Points(220));
        foreach (bool spacer in new[] { false, true }) {
            var box = presentation.AddSlide().AddTextBoxPoints("FIRST", 20, 20, 240, 160);
            box.FontName = "Helvetica"; box.FontSize = 12;
            if (spacer) {
                var empty = box.AddParagraph("");
                empty.Paragraph.Elements<A.Run>().Single().RunProperties = new A.RunProperties();
                empty.Paragraph.RemoveAllChildren<A.EndParagraphRunProperties>();
                if (source == "end") empty.Paragraph.Append(new A.EndParagraphRunProperties { FontSize = 3000 });
                else SetDefaults(box, empty.Paragraph, source, new A.DefaultRunProperties { FontSize = 3000 });
            }
            var second = box.AddParagraph("SECOND");
            if (differentAlignment) second.Alignment = PowerPointTextAlignment.Right;
        }
        using var pdf = UglyToad.PdfPig.PdfDocument.Open(presentation.ToPdf().ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true }));
        Assert.InRange(BaselineFromTop(pdf, 2, "SECOND") - BaselineFromTop(pdf, 1, "SECOND"), 35D, 37D);
    }

    [Fact]
    public void BlankTableCellParagraphRetainsItsLineAndSpacing() {
        using var presentation = PowerPointDocument.Create(new MemoryStream());
        presentation.SlideSize.SetSize(OfficeIMO.OfficeLength.Points(300),OfficeIMO.OfficeLength.Points(220));
        foreach (bool spacer in new[] { false, true }) {
            var cell = presentation.AddSlide().AddTablePoints(1, 1, 20, 20, 240, 160).GetCell(0, 0);
            cell.Text = "FIRST"; cell.FontName = "Helvetica"; cell.FontSize = 12;
            if (spacer) cell.AddParagraph("").SetSpaceBeforePoints(7).SetSpaceAfterPoints(11);
            cell.AddParagraph("SECOND");
        }
        var pdf = Read(presentation);
        Assert.InRange(Span(pdf.Pages[0], "SECOND").Y - Span(pdf.Pages[1], "SECOND").Y, 32D, 34D);
    }

    [Theory]
    [InlineData("paragraph")]
    [InlineData("list")]
    [InlineData("master")]
    [InlineData("end")]
    public void BlankTableCellParagraphUsesEffectiveFontSize(string source) {
        using var presentation = PowerPointDocument.Create(new MemoryStream());
        presentation.SlideSize.SetSize(OfficeIMO.OfficeLength.Points(300),OfficeIMO.OfficeLength.Points(220));
        foreach (bool spacer in new[] { false, true }) {
            var cell = presentation.AddSlide().AddTablePoints(1, 1, 20, 20, 240, 160).GetCell(0, 0);
            cell.Text = "FIRST"; cell.FontName = "Helvetica"; cell.FontSize = 12;
            if (spacer) {
                var empty = cell.AddParagraph("");
                empty.Paragraph.RemoveAllChildren<A.Run>();
                empty.Paragraph.RemoveAllChildren<A.EndParagraphRunProperties>();
                var defaults = new A.DefaultRunProperties { FontSize = 3000 };
                if (source == "end") empty.Paragraph.Append(new A.EndParagraphRunProperties { FontSize = 3000 });
                else if (source == "paragraph") empty.Paragraph.ParagraphProperties = new A.ParagraphProperties(defaults);
                else if (source == "list") cell.Cell.TextBody!.ListStyle!.Append(new A.Level1ParagraphProperties(defaults));
                else {
                    var master = cell.SlidePart!.SlideLayoutPart!.SlideMasterPart!.SlideMaster!.TextStyles!.OtherStyle!;
                    var level = master.GetFirstChild<A.Level1ParagraphProperties>() ?? master.AppendChild(new A.Level1ParagraphProperties());
                    level.RemoveAllChildren<A.DefaultRunProperties>(); level.Append(defaults);
                }
            }
            cell.AddParagraph("SECOND").Runs[0].FontSize = 12;
        }
        var pdf = Read(presentation);
        Assert.InRange(Span(pdf.Pages[0], "SECOND").Y - Span(pdf.Pages[1], "SECOND").Y, 35D, 37D);
    }

    [Theory]
    [InlineData("list", false)]
    [InlineData("list", true)]
    [InlineData("master", false)]
    [InlineData("master", true)]
    public void SingleEffectiveRunRetainsInheritedCapitalization(string source, bool flow) {
        using var presentation = PowerPointDocument.Create(new MemoryStream());
        var box = presentation.AddSlide().AddTextBoxPoints("mixed case", 20, 20, 240, 120);
        box.FontName = "Helvetica"; box.FontSize = 12;
        if (flow) box.AddParagraph("SECOND").Alignment = PowerPointTextAlignment.Right;
        SetDefaults(box, box.Paragraphs[0].Paragraph, source, new A.DefaultRunProperties { Capital = A.TextCapsValues.All });
        using var pdf = UglyToad.PdfPig.PdfDocument.Open(presentation.ToPdf().ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true }));
        Assert.Contains("MIXED CASE", pdf.GetPage(1).Text, StringComparison.Ordinal);
        Assert.Contains("MIXED CASE", string.Concat(Read(presentation).Pages[0].GetTextSpans().Select(span => span.Text)), StringComparison.Ordinal);
    }

    private static double BaselineFromTop(UglyToad.PdfPig.PdfDocument pdf, int pageNumber, string value) {
        var page = pdf.GetPage(pageNumber);
        string text = string.Concat(page.Letters.Select(letter => letter.Value));
        return page.Height - page.Letters[text.IndexOf(value, StringComparison.Ordinal)].StartBaseLine.Y;
    }

    private static void SetDefaults(PowerPointTextBox box, A.Paragraph paragraph, string source, A.DefaultRunProperties defaults) {
        if (source == "paragraph") paragraph.ParagraphProperties = new A.ParagraphProperties(defaults);
        else if (source == "list") box.TextBody!.ListStyle!.Append(new A.Level1ParagraphProperties(defaults));
        else {
            var style = Assert.IsType<P.OtherStyle>(box.MasterTextStyle);
            var level = style.GetFirstChild<A.Level1ParagraphProperties>() ?? style.AppendChild(new A.Level1ParagraphProperties());
            level.RemoveAllChildren<A.DefaultRunProperties>(); level.Append(defaults);
        }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void EmptyParagraphRetainsItsLineAndSpacingInPdfSpans(bool bullet) {
        using var presentation = PowerPointDocument.Create(new MemoryStream());
        presentation.SlideSize.SetSize(OfficeIMO.OfficeLength.Points(300),OfficeIMO.OfficeLength.Points(220));
        foreach (bool spacer in new[] { false, true }) {
            var box = presentation.AddSlide().AddTextBoxPoints("FIRST", 20, 20, 240, 160);
            box.FontName = "Helvetica";
            box.FontSize = 12;
            if (spacer) {
                var empty = box.AddParagraph("");
                empty.SetSpaceBeforePoints(7);
                empty.SetSpaceAfterPoints(11);
            }
            box.AddParagraph("SECOND").Alignment = PowerPointTextAlignment.Right;
            if (bullet) box.Paragraphs[0].SetBullet('*');
        }
        var pdf = Read(presentation);
        Assert.InRange(Span(pdf.Pages[0], "SECOND").Y - Span(pdf.Pages[1], "SECOND").Y, 32D, 34D);
    }

    [Theory]
    [InlineData("paragraph")]
    [InlineData("list")]
    [InlineData("master")]
    public void PlainParagraphUsesItsEffectiveInheritedColor(string source) {
        using var presentation = PowerPointDocument.Create(new MemoryStream());
        var slide = presentation.AddSlide();
        var box = slide.AddTextBoxPoints("INHERITED", 20, 20, 240, 100);
        box.FontName = "Helvetica";
        box.FontSize = 12;
        box.AddParagraph("SECOND").Alignment = PowerPointTextAlignment.Right;
        var paragraph = box.TextBody!.Elements<A.Paragraph>().First();
        paragraph.Elements<A.Run>().Single().RunProperties = new A.RunProperties();
        var defaults = new A.DefaultRunProperties(new A.SolidFill(new A.RgbColorModelHex { Val = "C02060" }));
        if (source == "paragraph") paragraph.ParagraphProperties = new A.ParagraphProperties(defaults);
        else if (source == "list") box.TextBody.ListStyle!.Append(new A.Level1ParagraphProperties(defaults));
        else {
            var style = Assert.IsType<P.OtherStyle>(box.MasterTextStyle);
            var level = style.GetFirstChild<A.Level1ParagraphProperties>() ?? style.AppendChild(new A.Level1ParagraphProperties());
            level.RemoveAllChildren<A.DefaultRunProperties>();
            level.Append(defaults);
        }
        var text = Span(Read(presentation).Pages[0], "INHERITED");
        Assert.Equal(OfficeColor.FromRgb(192, 32, 96), text.Color);
    }
    private static PdfCore.PdfReadDocument Read(PowerPointDocument presentation) =>
        PdfCore.PdfReadDocument.Open(presentation.ToPdf().ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true }));

    private static PdfCore.PdfTextSpan Span(PdfCore.PdfReadPage page, string text) =>
        Assert.Single(page.GetTextSpans(), span => span.Text.Contains(text, StringComparison.Ordinal));
}

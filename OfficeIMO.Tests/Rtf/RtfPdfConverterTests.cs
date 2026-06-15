using OfficeIMO.Rtf;
using OfficeIMO.Rtf.Pdf;
using PdfCore = OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Rtf;

public class RtfPdfConverterTests {
    [Fact]
    public void RtfDocument_ToPdfDocument_Renders_Paragraphs_Runs_And_PageSetup() {
        RtfDocument document = RtfDocument.Create();
        document.Info.Title = "RTF PDF";
        document.Info.Author = "OfficeIMO";
        document.PageSetup.SetPaperSize(11900, 16840);
        document.PageSetup.SetMargins(leftTwips: 1440, rightTwips: 1440, topTwips: 720, bottomTwips: 720);
        int red = document.AddColor(200, 20, 30);
        int mono = document.AddFont("Courier New");

        RtfParagraph paragraph = document.AddParagraph();
        paragraph.SetAlignment(RtfTextAlignment.Center);
        paragraph.AddText("Hello ");
        paragraph.AddText("bold").SetBold().SetForegroundColor(red).SetFontSize(16);
        paragraph.AddLineBreak();
        RtfRun monoRun = paragraph.AddText("mono");
        monoRun.FontId = mono;

        byte[] pdf = document.SaveAsPdf();
        string text = PdfCore.PdfReadDocument.Load(pdf).ExtractText();

        Assert.StartsWith("%PDF-", Encoding.ASCII.GetString(pdf, 0, 5), StringComparison.Ordinal);
        Assert.Contains("Hello", text, StringComparison.Ordinal);
        Assert.Contains("bold", text, StringComparison.Ordinal);
        Assert.Contains("mono", text, StringComparison.Ordinal);
    }

    [Fact]
    public void RtfString_ToPdfDocument_Renders_Field_Result_Text() {
        const string rtf = @"{\rtf1\ansi Parsed {\field{\*\fldinst HYPERLINK ""https://evotec.xyz""}{\fldrslt link}} text\par}";

        byte[] pdf = rtf.SaveAsPdf();
        string text = PdfCore.PdfReadDocument.Load(pdf).ExtractText();

        Assert.Contains("Parsed", text, StringComparison.Ordinal);
        Assert.Contains("link", text, StringComparison.Ordinal);
        Assert.Contains("text", text, StringComparison.Ordinal);
    }

    [Fact]
    public void RtfDocument_ToPdfDocument_Skips_Hidden_Text_Unless_Requested() {
        RtfDocument document = RtfDocument.Create();
        RtfParagraph paragraph = document.AddParagraph();
        paragraph.AddText("Visible ");
        paragraph.AddText("Hidden").SetHidden();

        string defaultText = PdfCore.PdfReadDocument.Load(document.SaveAsPdf()).ExtractText();
        string includedText = PdfCore.PdfReadDocument.Load(document.SaveAsPdf(new RtfPdfSaveOptions {
            IncludeHiddenText = true
        })).ExtractText();

        Assert.Contains("Visible", defaultText, StringComparison.Ordinal);
        Assert.DoesNotContain("Hidden", defaultText, StringComparison.Ordinal);
        Assert.Contains("Hidden", includedText, StringComparison.Ordinal);
    }

    [Fact]
    public void RtfDocument_ToPdfDocument_Renders_Tables() {
        RtfDocument document = RtfDocument.Create();
        RtfTable table = document.AddTable(2, 2);
        table.Rows[0].Cells[0].AddParagraph("A1");
        table.Rows[0].Cells[1].AddParagraph("B1");
        table.Rows[1].Cells[0].AddParagraph("A2");
        table.Rows[1].Cells[1].AddParagraph("B2");

        byte[] pdf = document.SaveAsPdf();
        string text = PdfCore.PdfReadDocument.Load(pdf).ExtractText();

        Assert.Contains("A1", text, StringComparison.Ordinal);
        Assert.Contains("B1", text, StringComparison.Ordinal);
        Assert.Contains("A2", text, StringComparison.Ordinal);
        Assert.Contains("B2", text, StringComparison.Ordinal);
    }

    [Fact]
    public void RtfDocument_ToPdfDocument_Renders_Explicit_ListText_Markers() {
        RtfDocument document = RtfDocument.Create();
        document.AddParagraph("Item").SetList(listId: 3, level: 0, kind: RtfListKind.Decimal).SetListText("7.\t");
        document.AddParagraph("Next").SetList(listId: 3, level: 0, kind: RtfListKind.Decimal);

        byte[] pdf = document.SaveAsPdf();
        string text = PdfCore.PdfReadDocument.Load(pdf).ExtractText();

        Assert.Contains("7.", text, StringComparison.Ordinal);
        Assert.Contains("8.", text, StringComparison.Ordinal);
        Assert.Contains("Item", text, StringComparison.Ordinal);
        Assert.Contains("Next", text, StringComparison.Ordinal);
    }

    [Fact]
    public void RtfDocument_ToPdfDocument_Generates_Semantic_List_Markers() {
        RtfDocument document = RtfDocument.Create();
        document.AddParagraph("First").SetList(listId: 9, level: 0, kind: RtfListKind.Decimal);
        document.AddParagraph("Second").SetList(listId: 9, level: 0, kind: RtfListKind.Decimal);
        document.AddParagraph("Bullet").SetList(listId: 10, level: 0, kind: RtfListKind.Bullet);

        byte[] pdf = document.SaveAsPdf();
        string text = PdfCore.PdfReadDocument.Load(pdf).ExtractText();

        Assert.Contains("1.", text, StringComparison.Ordinal);
        Assert.Contains("2.", text, StringComparison.Ordinal);
        Assert.Contains("\u2022", text, StringComparison.Ordinal);
        Assert.Contains("First", text, StringComparison.Ordinal);
        Assert.Contains("Second", text, StringComparison.Ordinal);
        Assert.Contains("Bullet", text, StringComparison.Ordinal);
    }

    [Fact]
    public void RtfDocument_ToPdfDocument_Renders_Default_Header_And_Footer_Text() {
        RtfDocument document = RtfDocument.Create();
        document.AddHeader().AddParagraph("Default header");
        document.AddFooter().AddParagraph("Default footer");
        document.AddParagraph("Body");

        byte[] pdf = document.SaveAsPdf();
        string text = PdfCore.PdfReadDocument.Load(pdf).ExtractText();

        Assert.Contains("Default header", text, StringComparison.Ordinal);
        Assert.Contains("Default footer", text, StringComparison.Ordinal);
        Assert.Contains("Body", text, StringComparison.Ordinal);
    }

    [Fact]
    public void RtfDocument_ToPdfDocument_Renders_First_And_Even_HeaderFooter_Variants() {
        RtfDocument document = RtfDocument.Create();
        document.PageSetup.SetDifferentFirstPageHeaderFooter();
        document.AddHeader(RtfHeaderFooterKind.RightHeader).AddParagraph("Odd header");
        document.AddHeader(RtfHeaderFooterKind.LeftHeader).AddParagraph("Even header");
        document.AddHeader(RtfHeaderFooterKind.FirstHeader).AddParagraph("First header");
        document.AddFooter(RtfHeaderFooterKind.RightFooter).AddParagraph("Odd footer");
        document.AddFooter(RtfHeaderFooterKind.LeftFooter).AddParagraph("Even footer");
        document.AddFooter(RtfHeaderFooterKind.FirstFooter).AddParagraph("First footer");

        RtfParagraph first = document.AddParagraph("First page");
        first.AddPageBreak();
        RtfParagraph second = document.AddParagraph("Second page");
        second.AddPageBreak();
        document.AddParagraph("Third page");

        byte[] pdf = document.SaveAsPdf();
        PdfCore.PdfReadDocument read = PdfCore.PdfReadDocument.Load(pdf);

        Assert.Equal(3, read.Pages.Count);
        Assert.Contains("First header", read.Pages[0].ExtractText(), StringComparison.Ordinal);
        Assert.Contains("First footer", read.Pages[0].ExtractText(), StringComparison.Ordinal);
        Assert.Contains("Even header", read.Pages[1].ExtractText(), StringComparison.Ordinal);
        Assert.Contains("Even footer", read.Pages[1].ExtractText(), StringComparison.Ordinal);
        Assert.Contains("Odd header", read.Pages[2].ExtractText(), StringComparison.Ordinal);
        Assert.Contains("Odd footer", read.Pages[2].ExtractText(), StringComparison.Ordinal);
    }

    [Fact]
    public void RtfDocument_ToPdfDocument_Can_Skip_HeaderFooter_Text() {
        RtfDocument document = RtfDocument.Create();
        document.AddHeader().AddParagraph("Hidden header");
        document.AddFooter().AddParagraph("Hidden footer");
        document.AddParagraph("Visible body");

        byte[] pdf = document.SaveAsPdf(new RtfPdfSaveOptions {
            IncludeHeaderFooters = false
        });
        string text = PdfCore.PdfReadDocument.Load(pdf).ExtractText();

        Assert.Contains("Visible body", text, StringComparison.Ordinal);
        Assert.DoesNotContain("Hidden header", text, StringComparison.Ordinal);
        Assert.DoesNotContain("Hidden footer", text, StringComparison.Ordinal);
    }
}

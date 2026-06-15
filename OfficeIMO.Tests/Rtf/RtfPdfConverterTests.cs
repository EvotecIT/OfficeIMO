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
}

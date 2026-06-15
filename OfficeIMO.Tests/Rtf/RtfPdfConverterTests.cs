using OfficeIMO.Rtf;
using OfficeIMO.Rtf.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
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
    public void RtfDocument_ToPdfDocument_Preserves_Table_Merge_Spans() {
        RtfDocument document = RtfDocument.Create();
        RtfTable table = document.AddTable(2, 3);
        table.Rows[0].Cells[0].HorizontalMerge = RtfTableCellMerge.First;
        table.Rows[0].Cells[0].VerticalMerge = RtfTableCellMerge.First;
        table.Rows[0].Cells[0].AddParagraph("Merged");
        table.Rows[0].Cells[1].HorizontalMerge = RtfTableCellMerge.Continue;
        table.Rows[0].Cells[2].AddParagraph("TopRight");
        table.Rows[1].Cells[0].VerticalMerge = RtfTableCellMerge.Continue;
        table.Rows[1].Cells[1].HorizontalMerge = RtfTableCellMerge.Continue;
        table.Rows[1].Cells[1].VerticalMerge = RtfTableCellMerge.Continue;
        table.Rows[1].Cells[2].AddParagraph("Body");

        PdfCore.PdfDocument pdfDocument = document.ToPdfDocument();
        PdfCore.TableBlock pdfTable = Assert.IsType<PdfCore.TableBlock>(Assert.Single(pdfDocument.Blocks));

        Assert.Equal(3, pdfTable.ColumnCount);
        Assert.Equal(2, pdfTable.Cells.Count);
        Assert.Equal(2, pdfTable.Cells[0].Count);
        Assert.Single(pdfTable.Cells[1]);
        Assert.NotNull(pdfTable.Style);
        Assert.Equal(0, pdfTable.Style!.HeaderRowCount);
        Assert.Equal(2, pdfTable.Cells[0][0].ColumnSpan);
        Assert.Equal(2, pdfTable.Cells[0][0].RowSpan);
        Assert.Equal("Merged", pdfTable.Cells[0][0].Text);
        Assert.Equal("TopRight", pdfTable.Cells[0][1].Text);
        Assert.Equal("Body", pdfTable.Cells[1][0].Text);

        string text = PdfCore.PdfReadDocument.Load(pdfDocument.ToBytes()).ExtractText();
        Assert.Contains("Merged", text, StringComparison.Ordinal);
        Assert.Contains("TopRight", text, StringComparison.Ordinal);
        Assert.Contains("Body", text, StringComparison.Ordinal);
    }

    [Fact]
    public void RtfDocument_ToPdfDocument_Renders_Section_Blocks_And_Breaks() {
        RtfDocument document = RtfDocument.Create();
        RtfSection first = document.AddSection();
        first.AddParagraph("First section");
        RtfSection second = document.AddSection(RtfSectionBreakKind.NextPage);
        second.AddParagraph("Second section");
        RtfSection continuous = document.AddSection(RtfSectionBreakKind.Continuous);
        continuous.AddParagraph("Continuous section");

        byte[] pdf = document.SaveAsPdf();
        PdfCore.PdfReadDocument read = PdfCore.PdfReadDocument.Load(pdf);

        Assert.Equal(2, read.Pages.Count);
        Assert.Contains("First section", read.Pages[0].ExtractText(), StringComparison.Ordinal);
        Assert.DoesNotContain("Second section", read.Pages[0].ExtractText(), StringComparison.Ordinal);
        Assert.Contains("Second section", read.Pages[1].ExtractText(), StringComparison.Ordinal);
        Assert.Contains("Continuous section", read.Pages[1].ExtractText(), StringComparison.Ordinal);
    }

    [Fact]
    public void RtfString_ToPdfDocument_Renders_Parsed_Section_Breaks() {
        const string rtf = @"{\rtf1\ansi\sectd\sbkpage\pard Parsed first\par\sect\sectd\sbkpage\pard Parsed second\par}";

        byte[] pdf = rtf.SaveAsPdf();
        PdfCore.PdfReadDocument read = PdfCore.PdfReadDocument.Load(pdf);

        Assert.Equal(2, read.Pages.Count);
        Assert.Contains("Parsed first", read.Pages[0].ExtractText(), StringComparison.Ordinal);
        Assert.Contains("Parsed second", read.Pages[1].ExtractText(), StringComparison.Ordinal);
    }

    [Fact]
    public void RtfDocument_ToPdfDocument_Applies_Section_PageSetup_To_Pdf_Pages() {
        RtfDocument document = RtfDocument.Create();
        document.PageSetup.SetPaperSize(12240, 15840);
        document.PageSetup.SetMargins(leftTwips: 1440, rightTwips: 1440, topTwips: 1440, bottomTwips: 1440);

        RtfSection first = document.AddSection();
        first.PageSetup.SetPaperSize(4800, 6400);
        first.PageSetup.SetMargins(leftTwips: 720, topTwips: 720);
        first.AddParagraph("Small first section");

        RtfSection second = document.AddSection(RtfSectionBreakKind.NextPage);
        second.PageSetup.SetPaperSize(4800, 8400);
        second.PageSetup.SetLandscape();
        second.PageSetup.SetMargins(leftTwips: 2880, rightTwips: 720, topTwips: 720, bottomTwips: 720);
        second.AddParagraph("Landscape second section");

        RtfSection continuous = document.AddSection(RtfSectionBreakKind.Continuous);
        continuous.AddParagraph("Continuous after landscape");

        byte[] pdf = document.SaveAsPdf();
        PdfCore.PdfDocumentInfo info = PdfCore.PdfInspector.Inspect(pdf);
        PdfCore.PdfReadDocument read = PdfCore.PdfReadDocument.Load(pdf);

        Assert.Equal(2, info.PageCount);
        Assert.Equal(240, info.Pages[0].Width, 1);
        Assert.Equal(320, info.Pages[0].Height, 1);
        Assert.Equal(420, info.Pages[1].Width, 1);
        Assert.Equal(240, info.Pages[1].Height, 1);
        Assert.Contains("Small first section", read.Pages[0].ExtractText(), StringComparison.Ordinal);
        Assert.Contains("Landscape second section", read.Pages[1].ExtractText(), StringComparison.Ordinal);
        Assert.Contains("Continuous after landscape", read.Pages[1].ExtractText(), StringComparison.Ordinal);
    }

    [Fact]
    public void RtfString_ToPdfDocument_Applies_Parsed_Section_PageSetup() {
        const string rtf = @"{\rtf1\ansi\sectd\sbkpage\pgwsxn4800\pghsxn6400\pard Parsed small\par\sect\sectd\sbkpage\pgwsxn4800\pghsxn8400\lndscpsxn\pard Parsed landscape\par}";

        byte[] pdf = rtf.SaveAsPdf();
        PdfCore.PdfDocumentInfo info = PdfCore.PdfInspector.Inspect(pdf);

        Assert.Equal(2, info.PageCount);
        Assert.Equal(240, info.Pages[0].Width, 1);
        Assert.Equal(320, info.Pages[0].Height, 1);
        Assert.Equal(420, info.Pages[1].Width, 1);
        Assert.Equal(240, info.Pages[1].Height, 1);
    }

    [Fact]
    public void RtfDocument_ToPdfDocument_Renders_Document_Page_Border() {
        RtfDocument document = RtfDocument.Create();
        document.PageSetup.SetPaperSize(4800, 6400);
        int red = document.AddColor(255, 0, 0);
        document.PageSetup.PageBorders.Top.Set(RtfPageBorderStyle.Single, width: 16, space: 24, colorIndex: red);
        document.AddParagraph("Bordered document");

        byte[] pdf = document.SaveAsPdf();
        string content = ExtractPdfContentStreams(pdf);

        Assert.Contains("1 0 0 RG", content, StringComparison.Ordinal);
        Assert.Contains("2 w", content, StringComparison.Ordinal);
        Assert.Contains("24 24 192 272 re", content, StringComparison.Ordinal);
    }

    [Fact]
    public void RtfDocument_ToPdfDocument_Renders_Section_Page_Border_Override() {
        RtfDocument document = RtfDocument.Create();
        document.PageSetup.SetPaperSize(4800, 6400);
        int red = document.AddColor(255, 0, 0);
        int blue = document.AddColor(0, 0, 255);
        document.PageSetup.PageBorders.Top.Set(RtfPageBorderStyle.Single, width: 16, space: 24, colorIndex: red);

        RtfSection first = document.AddSection();
        first.AddParagraph("First border section");

        RtfSection second = document.AddSection(RtfSectionBreakKind.NextPage);
        second.PageSetup.SetPaperSize(4800, 6400);
        second.PageSetup.PageBorders.Left.Set(RtfPageBorderStyle.Dotted, width: 8, space: 12, colorIndex: blue);
        second.AddParagraph("Second border section");

        byte[] pdf = document.SaveAsPdf();
        string content = ExtractPdfContentStreams(pdf);

        Assert.Contains("1 0 0 RG", content, StringComparison.Ordinal);
        Assert.Contains("24 24 192 272 re", content, StringComparison.Ordinal);
        Assert.Contains("0 0 1 RG", content, StringComparison.Ordinal);
        Assert.Contains("1 w", content, StringComparison.Ordinal);
        Assert.Contains("12 12 216 296 re", content, StringComparison.Ordinal);
    }

    [Fact]
    public void RtfString_ToPdfDocument_Renders_Paragraph_Indentation_And_Spacing() {
        const string rtf = @"{\rtf1\ansi\paperw12240\paperh15840\margl720\margr720\margt720\margb720\pard Plain\par\pard\li1440\ri720\fi720\sb720\sa0 Indented\par}";

        byte[] pdf = rtf.SaveAsPdf();
        using PdfPigDocument read = PdfPigDocument.Open(pdf);
        var words = read.GetPage(1).GetWords().ToList();
        var plain = Assert.Single(words, word => word.Text == "Plain");
        var indented = Assert.Single(words, word => word.Text == "Indented");

        Assert.True(indented.BoundingBox.Left > plain.BoundingBox.Left + 90D, $"Expected RTF left and first-line indents to move PDF text right. Plain={plain.BoundingBox.Left}; Indented={indented.BoundingBox.Left}.");
        Assert.True(plain.BoundingBox.Bottom > indented.BoundingBox.Bottom + 40D, $"Expected RTF spacing-before to increase the vertical gap between paragraphs. Plain={plain.BoundingBox.Bottom}; Indented={indented.BoundingBox.Bottom}.");
    }

    [Fact]
    public void RtfString_ToPdfDocument_Renders_Explicit_Tab_Stop_Alignment_And_Leader() {
        const string rtf = @"{\rtf1\ansi\paperw12240\paperh15840\margl720\margr720\margt720\margb720\pard\tqr\tldot\tx3600 Name\tab 12.34\par}";

        byte[] pdf = rtf.SaveAsPdf();
        using PdfPigDocument read = PdfPigDocument.Open(pdf);
        var letters = read.GetPage(1).Letters.OrderBy(letter => letter.StartBaseLine.X).ToList();
        var label = FindTextBounds(letters, "Name");
        var amount = FindTextBounds(letters, "12.34");
        double expectedRight = 36D + 180D;

        Assert.True(amount.Left > label.Right, $"Expected tabbed amount to render after label. LabelRight={label.Right}; AmountLeft={amount.Left}.");
        Assert.InRange(amount.Right, expectedRight - 10D, expectedRight + 10D);

        string content = ExtractPdfContentStreams(pdf);
        Assert.Contains("2E2E2E", content, StringComparison.Ordinal);

        static (double Left, double Right) FindTextBounds(IReadOnlyList<UglyToad.PdfPig.Content.Letter> letters, string text) {
            for (int index = 0; index <= letters.Count - text.Length; index++) {
                string candidate = string.Concat(letters.Skip(index).Take(text.Length).Select(letter => letter.Value));
                if (!string.Equals(candidate, text, StringComparison.Ordinal)) {
                    continue;
                }

                var matched = letters.Skip(index).Take(text.Length).ToList();
                return (matched.Min(letter => letter.StartBaseLine.X), matched.Max(letter => letter.EndBaseLine.X));
            }

            throw new Xunit.Sdk.XunitException("Expected text '" + text + "' in PDF letters. Text=" + string.Concat(letters.Select(letter => letter.Value)));
        }
    }

    [Fact]
    public void RtfDocument_ToPdfDocument_Renders_Paragraph_Line_Spacing() {
        RtfDocument document = RtfDocument.Create();
        document.PageSetup.SetPaperSize(12240, 15840);
        document.PageSetup.SetMargins(leftTwips: 720, rightTwips: 720, topTwips: 720, bottomTwips: 720);
        RtfParagraph paragraph = document.AddParagraph();
        paragraph.SetLineSpacing(480, multiple: true);
        paragraph.AddText("LineOne");
        paragraph.AddLineBreak();
        paragraph.AddText("LineTwo");

        byte[] pdf = document.SaveAsPdf();
        using PdfPigDocument read = PdfPigDocument.Open(pdf);
        var words = read.GetPage(1).GetWords().ToList();
        var lineOne = Assert.Single(words, word => word.Text == "LineOne");
        var lineTwo = Assert.Single(words, word => word.Text == "LineTwo");

        double lineGap = lineOne.BoundingBox.Bottom - lineTwo.BoundingBox.Bottom;
        Assert.True(lineGap > 20D, $"Expected RTF double line spacing to increase line advance in PDF output. Gap={lineGap}.");
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
    public void RtfDocument_ToPdfDocument_Renders_Notes_And_Can_Skip_Note_Bodies() {
        RtfDocument document = RtfDocument.Create();
        RtfParagraph paragraph = document.AddParagraph();
        paragraph.AddText("Body ");
        paragraph.AddFootnote("1", "Footnote body");
        paragraph.AddText(" and ");
        paragraph.AddEndnote("2", "Endnote body");
        RtfRun annotationRun = paragraph.AddAnnotation("3", "Annotation body");
        annotationRun.Note!.Author = "Alice";

        string defaultText = PdfCore.PdfReadDocument.Load(document.SaveAsPdf()).ExtractText();
        string skippedText = PdfCore.PdfReadDocument.Load(document.SaveAsPdf(new RtfPdfSaveOptions {
            IncludeNotes = false
        })).ExtractText();

        Assert.Contains("Body", defaultText, StringComparison.Ordinal);
        Assert.Contains("Footnote 1:", defaultText, StringComparison.Ordinal);
        Assert.Contains("Footnote body", defaultText, StringComparison.Ordinal);
        Assert.Contains("Endnote 2:", defaultText, StringComparison.Ordinal);
        Assert.Contains("Endnote body", defaultText, StringComparison.Ordinal);
        Assert.Contains("Annotation 3 (Alice):", defaultText, StringComparison.Ordinal);
        Assert.Contains("Annotation body", defaultText, StringComparison.Ordinal);
        Assert.DoesNotContain("Footnote body", skippedText, StringComparison.Ordinal);
        Assert.DoesNotContain("Endnote body", skippedText, StringComparison.Ordinal);
        Assert.DoesNotContain("Annotation body", skippedText, StringComparison.Ordinal);
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

    private static string ExtractPdfContentStreams(byte[] pdf) {
        string raw = Encoding.GetEncoding("ISO-8859-1").GetString(pdf);
        StringBuilder streams = new StringBuilder();
        int searchIndex = 0;
        while (true) {
            int streamStart = raw.IndexOf("stream", searchIndex, StringComparison.Ordinal);
            if (streamStart < 0) {
                return streams.ToString();
            }

            int dataStart = GetPdfStreamDataStart(raw, streamStart + "stream".Length);
            int dataEnd = raw.IndexOf("endstream", dataStart, StringComparison.Ordinal);
            if (dataEnd < 0) {
                return streams.ToString();
            }

            string streamData = raw.Substring(dataStart, dataEnd - dataStart);
            if (!TryInflatePdfStream(streamData, out string inflated)) {
                inflated = streamData;
            }

            streams.AppendLine(inflated);
            searchIndex = dataEnd + "endstream".Length;
        }
    }

    private static int GetPdfStreamDataStart(string raw, int index) {
        if (index < raw.Length && raw[index] == '\r' && index + 1 < raw.Length && raw[index + 1] == '\n') {
            return index + 2;
        }

        if (index < raw.Length && raw[index] == '\n') {
            return index + 1;
        }

        return index;
    }

    private static bool TryInflatePdfStream(string streamData, out string content) {
        byte[] compressed = Encoding.GetEncoding("ISO-8859-1").GetBytes(streamData);
        try {
            using MemoryStream input = new MemoryStream(compressed);
            using System.IO.Compression.DeflateStream deflate = new System.IO.Compression.DeflateStream(input, System.IO.Compression.CompressionMode.Decompress);
            using StreamReader reader = new StreamReader(deflate, Encoding.GetEncoding("ISO-8859-1"));
            content = reader.ReadToEnd();
            return !string.IsNullOrEmpty(content);
        } catch (InvalidDataException) {
            content = string.Empty;
            return false;
        }
    }
}

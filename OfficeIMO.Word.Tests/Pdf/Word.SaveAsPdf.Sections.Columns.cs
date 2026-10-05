using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using System;
using System.IO;
using System.Linq;
using System.Text;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Tests;

public partial class Word {
    [Theory]
    [InlineData("docx")]
    [InlineData("doc")]
    public void SaveAsPdf_Continuous_Sections_With_Equal_Columns_Balance_The_Preceding_Section(string extension) {
        string source = Path.Combine(_directoryWithFiles, "PdfContinuousSameColumns." + extension);
        string target = Path.Combine(_directoryWithFiles, "PdfContinuousSameColumns-" + extension + ".pdf");
        using (WordDocument document = WordDocument.Create(source)) {
            document.Sections[0].ColumnCount = 2;
            document.Sections[0].ColumnsSpace = 400;
            AddLines("FirstSection", 10);
            var second = document.AddSection(WordSectionBreakType.Continuous);
            second.ColumnCount = 2;
            second.ColumnsSpace = 400;
            AddLines("NextSection", 20);
            document.Save();

            void AddLines(string prefix, int count) {
                var paragraph = document.AddParagraph(string.Join("\n", Enumerable.Range(1, count).Select(index => prefix + index.ToString("D3"))));
                paragraph.FontFamily = "Arial"; paragraph.FontSize = 12;
                paragraph.LineSpacing = 400; paragraph.LineSpacingRule = WordLineSpacingRule.Exact;
                paragraph.LineSpacingBeforePoints = 0; paragraph.LineSpacingAfterPoints = 0;
                paragraph.AvoidWidowAndOrphanOverride = false;
            }
        }
        using (WordDocument document = WordDocument.Load(source)) {
            document.SaveAsPdf(target, new WordToPdfOptions { IncludePageNumbers = false,
                PageSize = new PdfCore.PageSize(500, 400), Margins = PdfCore.PageMargins.Uniform(40) });
        }
        using var pdf = PdfPigDocument.Open(File.ReadAllBytes(target));
        Assert.Equal(1, pdf.NumberOfPages);
        Assert.InRange(FindWordStartX(pdf.GetPage(1), "FirstSection006"), 259.9, 260.1);
        Assert.InRange(FindWordStartX(pdf.GetPage(1), "NextSection007"), 39.9, 40.1);
        Assert.InRange(FindWordStartX(pdf.GetPage(1), "NextSection020"), 259.9, 260.1);
    }

    [Fact]
    public void SaveAsPdf_OfficeIMOEngine_Maps_Word_Section_Columns_To_RowColumn_Flow() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfNativeSectionColumns.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeSectionColumns.pdf");

        using (WordDocument document = WordDocument.Create(docPath)) {
            WordSection section = document.Sections[0];
            section.ColumnCount = 2;
            section.ColumnsSpace = 720;

            document.AddParagraph("LeftColumnMarker starts in the first Word section column.")
                .AddBreak(WordBreakType.Column);
            document.AddParagraph("RightColumnMarker starts in the second Word section column.");

            document.Save();
            document.SaveAsPdf(pdfPath, new WordToPdfOptions {
                IncludePageNumbers = false,
                PageSize = new PdfCore.PageSize(612, 792),
                Margins = PdfCore.PageMargins.Uniform(36)
            });
        }

        byte[] bytes = File.ReadAllBytes(pdfPath);
        using PdfPigDocument pdf = PdfPigDocument.Open(bytes);
        var page = pdf.GetPage(1);
        string text = page.Text;
        Assert.Contains("LeftColumnMarker", text);
        Assert.Contains("RightColumnMarker", text);

        double leftX = FindWordStartX(page, "LeftColumnMarker");
        double rightX = FindWordStartX(page, "RightColumnMarker");
        Assert.InRange(leftX, 35D, 48D);
        Assert.True(rightX > leftX + 250D, $"Expected the second Word section column to render to the right of the first. Left x: {leftX:0.##}, right x: {rightX:0.##}.");
    }

    [Fact]
    public void SaveAsPdf_OfficeIMOEngine_Maps_Unequal_Word_Section_Column_Widths() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfNativeUnequalSectionColumns.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeUnequalSectionColumns.pdf");

        using (WordDocument document = WordDocument.Create(docPath)) {
            WordSection section = document.Sections[0];
            section.ColumnCount = 2;
            section.ColumnsSpace = 720;
            Columns columns = section._sectionProperties.GetFirstChild<Columns>()!;
            columns.EqualWidth = false;
            columns.RemoveAllChildren<Column>();
            columns.Append(
                new Column { Width = "1440", Space = "720" },
                new Column { Width = "4320" });

            WordParagraph left = document.AddParagraph("LeftMarker starts in the explicitly narrow first Word section column.");
            left.FontFamily = "Arial"; left.FontSize = 12;
            left.AddBreak(WordBreakType.Column);
            WordParagraph right = document.AddParagraph("RightMarker starts in the wider second Word section column.");
            right.FontFamily = "Arial"; right.FontSize = 12;

            document.Save();
            document.SaveAsPdf(pdfPath, new WordToPdfOptions {
                IncludePageNumbers = false,
                PageSize = new PdfCore.PageSize(612, 792),
                Margins = PdfCore.PageMargins.Uniform(36)
            });
        }

        byte[] bytes = File.ReadAllBytes(pdfPath);
        using PdfPigDocument pdf = PdfPigDocument.Open(bytes);
        var page = pdf.GetPage(1);
        Assert.Contains("LeftMarker", page.Text);
        Assert.Contains("RightMarker", page.Text);

        double leftX = FindWordStartX(page, "LeftMarker");
        double rightX = FindWordStartX(page, "RightMarker");

        Assert.InRange(leftX, 35D, 48D);
        // Explicit widths are points after twip conversion; unused page width is not redistributed.
        Assert.InRange(rightX - leftX, 107.9D, 108.1D);
    }

    [Fact]
    public void SaveAsPdf_OfficeIMOEngine_Maps_Word_Section_Column_Separator() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfNativeSectionColumnSeparator.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeSectionColumnSeparator.pdf");

        using (WordDocument document = WordDocument.Create(docPath)) {
            WordSection section = document.Sections[0];
            section.ColumnCount = 2;
            section.ColumnsSpace = 720;
            section.HasColumnSeparator = true;

            document.AddParagraph("SeparatorLeftMarker starts in the first Word section column.")
                .AddBreak(WordBreakType.Column);
            document.AddParagraph("SeparatorRightMarker starts in the second Word section column.");

            document.Save();
            document.SaveAsPdf(pdfPath, new WordToPdfOptions {
                IncludePageNumbers = false,
                PageSize = new PdfCore.PageSize(612, 792),
                Margins = PdfCore.PageMargins.Uniform(36)
            });
        }

        byte[] bytes = File.ReadAllBytes(pdfPath);
        string rawPdf = PdfOperatorSearchText.From(bytes);
        using PdfPigDocument pdf = PdfPigDocument.Open(bytes);
        var page = pdf.GetPage(1);

        Assert.Contains("SeparatorLeftMarker", page.Text);
        Assert.Contains("SeparatorRightMarker", page.Text);
        Assert.Contains("0.5 w", rawPdf, StringComparison.Ordinal);
        Assert.Contains("306 756 m 306 ", rawPdf, StringComparison.Ordinal);
    }

    [Fact]
    public void SaveAsPdf_OfficeIMOEngine_Fills_First_Word_Section_Column_Before_Advancing() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfNativeAutomaticSectionColumns.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeAutomaticSectionColumns.pdf");

        using (WordDocument document = WordDocument.Create(docPath)) {
            WordSection section = document.Sections[0];
            section.ColumnCount = 2;
            section.ColumnsSpace = 720;

            document.AddParagraph("AutoLeftColumnMarker starts in the first automatic Word section column.");
            document.AddParagraph("AutoRightColumnMarker starts in the second automatic Word section column.");

            document.Save();
            document.SaveAsPdf(pdfPath, new WordToPdfOptions {
                IncludePageNumbers = false,
                PageSize = new PdfCore.PageSize(612, 792),
                Margins = PdfCore.PageMargins.Uniform(36)
            });
        }

        byte[] bytes = File.ReadAllBytes(pdfPath);
        using PdfPigDocument pdf = PdfPigDocument.Open(bytes);
        var page = pdf.GetPage(1);
        string text = page.Text;
        Assert.Contains("AutoLeftColumnMarker", text);
        Assert.Contains("AutoRightColumnMarker", text);

        double leftX = FindWordStartX(page, "AutoLeftColumnMarker");
        double rightX = FindWordStartX(page, "AutoRightColumnMarker");
        Assert.InRange(leftX, 35D, 48D);
        Assert.InRange(Math.Abs(rightX - leftX), 0D, 0.1D);
    }

    [Fact]
    public void SaveAsPdf_OfficeIMOEngine_Keeps_Automatic_Column_Headings_With_Following_Content() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfNativeAutomaticSectionColumnHeadingKeep.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeAutomaticSectionColumnHeadingKeep.pdf");

        using (WordDocument document = WordDocument.Create(docPath)) {
            WordSection section = document.Sections[0];
            section.ColumnCount = 2;
            section.ColumnsSpace = 720;

            WordParagraph prelude = document.AddParagraph("ColumnKeepPrelude\n" + string.Join("\n", Enumerable.Range(1, 14).Select(index => "Prelude" + index)));
            WordParagraph heading = document.AddParagraph("ColumnKeepHeading").SetStyle(WordParagraphStyles.Heading2);
            WordParagraph body = document.AddParagraph("ColumnKeepBody\nSecond body line\nThird body line");
            foreach (WordParagraph paragraph in new[] { prelude, heading, body }) {
                paragraph.FontSize = 12;
                paragraph.LineSpacing = 400;
                paragraph.LineSpacingRule = WordLineSpacingRule.Exact;
                paragraph.LineSpacingBeforePoints = 0;
                paragraph.LineSpacingAfterPoints = 0;
            }

            document.Save();
            document.SaveAsPdf(pdfPath, new WordToPdfOptions {
                IncludePageNumbers = false,
                PageSize = new PdfCore.PageSize(500, 400),
                Margins = PdfCore.PageMargins.Uniform(40)
            });
        }

        byte[] bytes = File.ReadAllBytes(pdfPath);
        using PdfPigDocument pdf = PdfPigDocument.Open(bytes);
        var page = pdf.GetPage(1);

        double preludeX = FindWordStartX(page, "ColumnKeepPrelude");
        double headingX = FindWordStartX(page, "ColumnKeepHeading");
        double bodyX = FindWordStartX(page, "ColumnKeepBody");

        Assert.InRange(preludeX, 35D, 48D);
        Assert.True(headingX > preludeX + 200D, $"Expected the kept heading to move into the second automatic column. Prelude x: {preludeX:0.##}, heading x: {headingX:0.##}.");
        Assert.InRange(Math.Abs(bodyX - headingX), 0D, 8D);
    }

    [Fact]
    public void SaveAsPdf_OfficeIMOEngine_Keeps_Styled_Automatic_Column_Content_With_Following_Content() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfNativeAutomaticSectionColumnStyledKeep.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeAutomaticSectionColumnStyledKeep.pdf");

        using (WordDocument document = WordDocument.Create(docPath)) {
            WordSection section = document.Sections[0];
            section.ColumnCount = 2;
            section.ColumnsSpace = 720;

            const string styleId = "AutomaticColumnStyledKeepNext";
            Styles styles = document._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
            styles.Append(new Style(
                new StyleName { Val = "Automatic Column Styled Keep Next" },
                new BasedOn { Val = "Normal" },
                new StyleParagraphProperties(new KeepNext()))
            {
                Type = StyleValues.Paragraph,
                StyleId = styleId,
                CustomStyle = true
            });

            WordParagraph prelude = document.AddParagraph("StyledColumnKeepPrelude\n" + string.Join("\n", Enumerable.Range(1, 14).Select(index => "StyledPrelude" + index)));
            WordParagraph heading = document.AddParagraph("StyledColumnKeepHeading").SetStyleId(styleId);
            WordParagraph body = document.AddParagraph("StyledColumnKeepBody\nSecond body line\nThird body line");
            foreach (WordParagraph paragraph in new[] { prelude, heading, body }) {
                paragraph.FontSize = 12;
                paragraph.LineSpacing = 400;
                paragraph.LineSpacingRule = WordLineSpacingRule.Exact;
                paragraph.LineSpacingBeforePoints = 0;
                paragraph.LineSpacingAfterPoints = 0;
            }

            document.Save();
            document.SaveAsPdf(pdfPath, new WordToPdfOptions {
                IncludePageNumbers = false,
                PageSize = new PdfCore.PageSize(500, 400),
                Margins = PdfCore.PageMargins.Uniform(40)
            });
        }

        byte[] bytes = File.ReadAllBytes(pdfPath);
        using PdfPigDocument pdf = PdfPigDocument.Open(bytes);
        var page = pdf.GetPage(1);

        double preludeX = FindWordStartX(page, "StyledColumnKeepPrelude");
        double headingX = FindWordStartX(page, "StyledColumnKeepHeading");
        double bodyX = FindWordStartX(page, "StyledColumnKeepBody");

        Assert.InRange(preludeX, 35D, 48D);
        Assert.True(headingX > preludeX + 200D, $"Expected the styled keep-next paragraph to move into the second automatic column. Prelude x: {preludeX:0.##}, heading x: {headingX:0.##}.");
        Assert.InRange(Math.Abs(bodyX - headingX), 0D, 8D);
    }

    [Fact]
    public void SaveAsPdf_OfficeIMOEngine_Splits_Inline_Word_Column_Breaks() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfNativeInlineSectionColumnBreak.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeInlineSectionColumnBreak.pdf");

        using (WordDocument document = WordDocument.Create(docPath)) {
            WordSection section = document.Sections[0];
            section.ColumnCount = 2;
            section.ColumnsSpace = 720;

            WordParagraph paragraph = document.AddParagraph();
            paragraph.AddText("InlineLeftColumnMarker remains before the inline Word column break.");
            paragraph.AddBreak(WordBreakType.Column);
            paragraph.AddText("InlineRightColumnMarker starts after the inline Word column break.");

            document.Save();
            document.SaveAsPdf(pdfPath, new WordToPdfOptions {
                IncludePageNumbers = false,
                PageSize = new PdfCore.PageSize(612, 792),
                Margins = PdfCore.PageMargins.Uniform(36)
            });
        }

        byte[] bytes = File.ReadAllBytes(pdfPath);
        using PdfPigDocument pdf = PdfPigDocument.Open(bytes);
        var page = pdf.GetPage(1);
        string text = page.Text;
        Assert.Contains("InlineLeftColumnMarker", text);
        Assert.Contains("InlineRightColumnMarker", text);

        double leftX = FindWordStartX(page, "InlineLeftColumnMarker");
        double rightX = FindWordStartX(page, "InlineRightColumnMarker");
        Assert.InRange(leftX, 35D, 48D);
        Assert.True(rightX > leftX + 250D, $"Expected text after an inline Word column break to render in the next section column. Left x: {leftX:0.##}, right x: {rightX:0.##}.");
    }
}

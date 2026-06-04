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
    [Fact]
    public void SaveAsPdf_OfficeIMOEngine_Maps_Word_Section_Columns_To_RowColumn_Flow() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfNativeSectionColumns.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeSectionColumns.pdf");

        using (WordDocument document = WordDocument.Create(docPath)) {
            WordSection section = document.Sections[0];
            section.ColumnCount = 2;
            section.ColumnsSpace = 720;

            document.AddParagraph("LeftColumnMarker starts in the first Word section column.")
                .AddBreak(BreakValues.Column);
            document.AddParagraph("RightColumnMarker starts in the second Word section column.");

            document.Save();
            document.SaveAsPdf(pdfPath, new PdfSaveOptions {
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

            document.AddParagraph("NarrowColumnMarker starts in the explicitly narrow first Word section column.")
                .AddBreak(BreakValues.Column);
            document.AddParagraph("WideColumnMarker starts in the wider second Word section column.");

            document.Save();
            document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                IncludePageNumbers = false,
                PageSize = new PdfCore.PageSize(612, 792),
                Margins = PdfCore.PageMargins.Uniform(36)
            });
        }

        byte[] bytes = File.ReadAllBytes(pdfPath);
        using PdfPigDocument pdf = PdfPigDocument.Open(bytes);
        var page = pdf.GetPage(1);
        Assert.Contains("NarrowColumnMarker", page.Text);
        Assert.Contains("WideColumnMarker", page.Text);

        double leftX = FindWordStartX(page, "NarrowColumnMarker");
        double rightX = FindWordStartX(page, "WideColumnMarker");

        Assert.InRange(leftX, 35D, 48D);
        Assert.InRange(rightX - leftX, 145D, 190D);
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
                .AddBreak(BreakValues.Column);
            document.AddParagraph("SeparatorRightMarker starts in the second Word section column.");

            document.Save();
            document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                IncludePageNumbers = false,
                PageSize = new PdfCore.PageSize(612, 792),
                Margins = PdfCore.PageMargins.Uniform(36)
            });
        }

        byte[] bytes = File.ReadAllBytes(pdfPath);
        string rawPdf = Encoding.ASCII.GetString(bytes);
        using PdfPigDocument pdf = PdfPigDocument.Open(bytes);
        var page = pdf.GetPage(1);

        Assert.Contains("SeparatorLeftMarker", page.Text);
        Assert.Contains("SeparatorRightMarker", page.Text);
        Assert.Contains("0.5 w", rawPdf, StringComparison.Ordinal);
        Assert.Contains("306 756 m 306 ", rawPdf, StringComparison.Ordinal);
    }

    [Fact]
    public void SaveAsPdf_OfficeIMOEngine_Distributes_Word_Section_Columns_Without_Explicit_Breaks() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfNativeAutomaticSectionColumns.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeAutomaticSectionColumns.pdf");

        using (WordDocument document = WordDocument.Create(docPath)) {
            WordSection section = document.Sections[0];
            section.ColumnCount = 2;
            section.ColumnsSpace = 720;

            document.AddParagraph("AutoLeftColumnMarker starts in the first automatic Word section column.");
            document.AddParagraph("AutoRightColumnMarker starts in the second automatic Word section column.");

            document.Save();
            document.SaveAsPdf(pdfPath, new PdfSaveOptions {
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
        Assert.True(rightX > leftX + 250D, $"Expected automatic second Word section column content to render to the right of the first. Left x: {leftX:0.##}, right x: {rightX:0.##}.");
    }

    [Fact]
    public void SaveAsPdf_OfficeIMOEngine_Keeps_Automatic_Column_Headings_With_Following_Content() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfNativeAutomaticSectionColumnHeadingKeep.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeAutomaticSectionColumnHeadingKeep.pdf");

        using (WordDocument document = WordDocument.Create(docPath)) {
            WordSection section = document.Sections[0];
            section.ColumnCount = 2;
            section.ColumnsSpace = 720;

            document.AddParagraph("ColumnKeepPrelude " + string.Join(" ", Enumerable.Range(1, 42).Select(index => "prelude" + index.ToString(System.Globalization.CultureInfo.InvariantCulture))));
            document.AddParagraph("ColumnKeepHeading").SetStyle(WordParagraphStyles.Heading2);
            document.AddParagraph("ColumnKeepBody follows the heading and should stay in the same automatic Word section column.");

            document.Save();
            document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                IncludePageNumbers = false,
                PageSize = new PdfCore.PageSize(612, 792),
                Margins = PdfCore.PageMargins.Uniform(36)
            });
        }

        byte[] bytes = File.ReadAllBytes(pdfPath);
        using PdfPigDocument pdf = PdfPigDocument.Open(bytes);
        var page = pdf.GetPage(1);

        double preludeX = FindWordStartX(page, "ColumnKeepPrelude");
        double headingX = FindWordStartX(page, "ColumnKeepHeading");
        double bodyX = FindWordStartX(page, "ColumnKeepBody");

        Assert.InRange(preludeX, 35D, 48D);
        Assert.True(headingX > preludeX + 250D, $"Expected the kept heading to move into the second automatic column. Prelude x: {preludeX:0.##}, heading x: {headingX:0.##}.");
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
            paragraph.AddBreak(BreakValues.Column);
            paragraph.AddText("InlineRightColumnMarker starts after the inline Word column break.");

            document.Save();
            document.SaveAsPdf(pdfPath, new PdfSaveOptions {
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

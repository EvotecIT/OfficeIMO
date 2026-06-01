using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using System;
using System.IO;
using System.Linq;
using System.Text;
using UglyToad.PdfPig;
using Xunit;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Tests;

public partial class Word {
    [Fact]
    public void Test_WordDocument_SaveAsPdf_MultipleSections() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfSections.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, "PdfSections.pdf");

        using (WordDocument document = WordDocument.Create(docPath)) {
            document.AddHeadersAndFooters();
            var defaultHeader = RequireSectionHeader(document, 0, HeaderFooterValues.Default);
            var defaultFooter = RequireSectionFooter(document, 0, HeaderFooterValues.Default);
            defaultHeader.AddParagraph("Header1");
            defaultFooter.AddParagraph("Footer1");
            document.AddParagraph("Section1 Paragraph");

            WordSection section2 = document.AddSection();
            section2.AddHeadersAndFooters();
            var section2Header = RequireSectionHeader(document, 1, HeaderFooterValues.Default);
            var section2Footer = RequireSectionFooter(document, 1, HeaderFooterValues.Default);
            section2Header.AddParagraph("Header2");
            section2Footer.AddParagraph("Footer2");
            document.AddParagraph("Section2 Paragraph");

            document.Save();
            document.SaveAsPdf(pdfPath);
        }

        Assert.True(File.Exists(pdfPath));
    }

    [Fact]
    public void Test_WordDocument_SaveAsPdf_HeaderFooterVariants() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfHeaderVariants.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, "PdfHeaderVariants.pdf");

        using (WordDocument document = WordDocument.Create(docPath)) {
            document.AddHeadersAndFooters();
            var header = RequireSectionHeader(document, 0, HeaderFooterValues.Default);
            var footer = RequireSectionFooter(document, 0, HeaderFooterValues.Default);
            header.AddParagraph("DefaultHeader");
            footer.AddParagraph("DefaultFooter");

            for (int i = 0; i < 100; i++) {
                document.AddParagraph($"Paragraph {i}");
            }

            document.Save();
            document.SaveAsPdf(pdfPath);
        }

        Assert.True(File.Exists(pdfPath));
    }

    [Fact]
    public void SaveAsPdf_OfficeIMOEngine_Renders_Default_Header_And_Footer_Text() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfNativeHeaderFooter.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeHeaderFooter.pdf");

        using (WordDocument document = WordDocument.Create(docPath)) {
            document.AddHeadersAndFooters();
            RequireSectionHeader(document, 0, HeaderFooterValues.Default).AddParagraph("Native Default Header");
            RequireSectionFooter(document, 0, HeaderFooterValues.Default).AddParagraph("Native Default Footer");
            document.AddParagraph("Native body text");
            document.Save();
            document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                IncludePageNumbers = false
            });
        }

        Assert.True(File.Exists(pdfPath));
        using (PdfDocument pdf = PdfDocument.Open(pdfPath)) {
            string allText = string.Concat(pdf.GetPages().Select(page => page.Text));
            Assert.Contains("Native Default Header", allText);
            Assert.Contains("Native Default Footer", allText);
            Assert.Contains("Native body text", allText);
        }
    }

    [Fact]
    public void SaveAsPdf_OfficeIMOEngine_Maps_HeaderFooter_Paragraph_Alignment_To_Zones() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfNativeAlignedHeaderFooter.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeAlignedHeaderFooter.pdf");

        using (WordDocument document = WordDocument.Create(docPath)) {
            document.AddHeadersAndFooters();
            WordParagraph centeredHeader = RequireSectionHeader(document, 0, HeaderFooterValues.Default).AddParagraph("NativeCenterHeader");
            centeredHeader.ParagraphAlignment = JustificationValues.Center;
            WordParagraph rightFooter = RequireSectionFooter(document, 0, HeaderFooterValues.Default).AddParagraph("NativeRightFooter");
            rightFooter.ParagraphAlignment = JustificationValues.Right;
            document.AddParagraph("NativeAlignedBody");

            document.Save();
            document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                IncludePageNumbers = false,
                PageSize = new PdfCore.PageSize(400, 300),
                Margins = PdfCore.PageMargins.Uniform(50)
            });
        }

        Assert.True(File.Exists(pdfPath));
        using (PdfDocument pdf = PdfDocument.Open(pdfPath)) {
            var page = pdf.GetPage(1);
            string allText = page.Text;
            Assert.Contains("NativeCenterHeader", allText);
            Assert.Contains("NativeRightFooter", allText);
            Assert.Contains("NativeAlignedBody", allText);

            double bodyX = FindWordStartX(page, "NativeAlignedBody");
            double headerX = FindWordStartX(page, "NativeCenterHeader");
            double footerX = FindWordStartX(page, "NativeRightFooter");

            Assert.InRange(bodyX, 49D, 61D);
            Assert.True(headerX > bodyX + 45D, $"Expected centered header to render away from the left margin. Header x: {headerX:0.##}, body x: {bodyX:0.##}.");
            Assert.True(footerX > headerX + 55D, $"Expected right footer to render to the right of centered header. Footer x: {footerX:0.##}, header x: {headerX:0.##}.");
        }
    }

    [Fact]
    public void SaveAsPdf_OfficeIMOEngine_Maps_HeaderFooter_Table_Cells_To_Zones() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfNativeTableHeaderFooter.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeTableHeaderFooter.pdf");

        using (WordDocument document = WordDocument.Create(docPath)) {
            document.AddHeadersAndFooters();
            WordTable headerTable = RequireSectionHeader(document, 0, HeaderFooterValues.Default).AddTable(1, 3, WordTableStyle.TableNormal);
            headerTable.Rows[0].Cells[0].Paragraphs[0].Text = "LHdr";
            headerTable.Rows[0].Cells[1].Paragraphs[0].Text = "CHdr";
            headerTable.Rows[0].Cells[2].Paragraphs[0].Text = "RHdr";

            WordTable footerTable = RequireSectionFooter(document, 0, HeaderFooterValues.Default).AddTable(1, 2, WordTableStyle.TableNormal);
            footerTable.Rows[0].Cells[0].Paragraphs[0].Text = "LFtr";
            footerTable.Rows[0].Cells[1].Paragraphs[0].Text = "RFtr";

            document.AddParagraph("NativeTableZoneBody");
            document.Save();
            document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                IncludePageNumbers = false,
                PageSize = new PdfCore.PageSize(520, 320),
                Margins = PdfCore.PageMargins.Uniform(60)
            });
        }

        Assert.True(File.Exists(pdfPath));
        using (PdfDocument pdf = PdfDocument.Open(pdfPath)) {
            var page = pdf.GetPage(1);
            string allText = page.Text;
            Assert.Contains("LHdr", allText);
            Assert.Contains("CHdr", allText);
            Assert.Contains("RHdr", allText);
            Assert.Contains("LFtr", allText);
            Assert.Contains("RFtr", allText);
            Assert.Contains("NativeTableZoneBody", allText);

            double bodyX = FindWordStartX(page, "NativeTableZoneBody");
            double leftHeaderX = FindWordStartX(page, "LHdr");
            double centerHeaderX = FindWordStartX(page, "CHdr");
            double rightHeaderX = FindWordStartX(page, "RHdr");
            double leftFooterX = FindWordStartX(page, "LFtr");
            double rightFooterX = FindWordStartX(page, "RFtr");

            Assert.InRange(bodyX, 58D, 72D);
            Assert.InRange(leftHeaderX, 58D, 72D);
            Assert.InRange(leftFooterX, 58D, 72D);
            Assert.True(centerHeaderX > leftHeaderX + 75D, $"Expected center header table cell to render away from the left zone. Center x: {centerHeaderX:0.##}, left x: {leftHeaderX:0.##}.");
            Assert.True(rightHeaderX > centerHeaderX + 75D, $"Expected right header table cell to render to the right of the center zone. Right x: {rightHeaderX:0.##}, center x: {centerHeaderX:0.##}.");
            Assert.True(rightFooterX > leftFooterX + 150D, $"Expected two-cell footer table to map the last cell to the right zone. Right x: {rightFooterX:0.##}, left x: {leftFooterX:0.##}.");
        }
    }

    [Fact]
    public void SaveAsPdf_OfficeIMOEngine_Uses_Explicit_Pdf_Page_Setup_Options() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfNativeOfficeIMOPageSetup.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeOfficeIMOPageSetup.pdf");

        using (WordDocument document = WordDocument.Create(docPath)) {
            document.AddParagraph("Native page setup marker");
            document.Save();
            document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                IncludePageNumbers = false,
                PageSize = new PdfCore.PageSize(240, 320),
                Margins = new PdfCore.PageMargins(80, 36, 36, 36)
            });
        }

        byte[] bytes = File.ReadAllBytes(pdfPath);
        PdfCore.PdfPageInfo pageInfo = Assert.Single(PdfCore.PdfInspector.Inspect(bytes).Pages);
        Assert.Equal(240, pageInfo.Width, 1);
        Assert.Equal(320, pageInfo.Height, 1);

        using PdfDocument pdf = PdfDocument.Open(bytes);
        var firstLetter = pdf.GetPage(1).Letters.First(letter => letter.Value == "N");
        Assert.InRange(firstLetter.StartBaseLine.X, 78, 92);
    }

    [Fact]
    public void SaveAsPdf_OfficeIMOEngine_Preserves_Explicit_Pdf_Page_Size_Geometry() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfNativeExplicitPageGeometry.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeExplicitPageGeometry.pdf");

        using (WordDocument document = WordDocument.Create(docPath)) {
            document.AddParagraph("Native explicit geometry marker");
            document.Save();
            document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                IncludePageNumbers = false,
                PageSize = new PdfCore.PageSize(420, 240),
                Margins = PdfCore.PageMargins.Uniform(36)
            });
        }

        byte[] bytes = File.ReadAllBytes(pdfPath);
        PdfCore.PdfPageInfo pageInfo = Assert.Single(PdfCore.PdfInspector.Inspect(bytes).Pages);
        Assert.Equal(420, pageInfo.Width, 1);
        Assert.Equal(240, pageInfo.Height, 1);

        using PdfDocument pdf = PdfDocument.Open(bytes);
        var firstLetter = pdf.GetPage(1).Letters.First(letter => letter.Value == "N");
        Assert.InRange(firstLetter.StartBaseLine.X, 35D, 48D);
    }

    [Fact]
    public void SaveAsPdf_OfficeIMOEngine_Uses_Word_Section_Page_Setup_And_Margins() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfNativeWordSectionPageSetup.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeWordSectionPageSetup.pdf");

        using (WordDocument document = WordDocument.Create(docPath)) {
            WordSection firstSection = document.Sections[0];
            firstSection.PageSettings.PageSize = WordPageSize.Letter;
            firstSection.PageOrientation = PageOrientationValues.Portrait;
            firstSection.SetMargins(WordMargin.Narrow);
            document.AddParagraph("NarrowMarginMarker starts from the Word section margin.");

            WordSection secondSection = document.AddSection();
            secondSection.PageSettings.PageSize = WordPageSize.Letter;
            secondSection.PageOrientation = PageOrientationValues.Landscape;
            secondSection.SetMargins(WordMargin.Wide);
            secondSection.AddParagraph("WideMarginMarker starts from the wider Word section margin.");

            document.Save();
            document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                IncludePageNumbers = false
            });
        }

        byte[] bytes = File.ReadAllBytes(pdfPath);
        PdfCore.PdfDocumentInfo info = PdfCore.PdfInspector.Inspect(bytes);
        Assert.Equal(2, info.PageCount);
        Assert.Equal(612, info.Pages[0].Width, 1);
        Assert.Equal(792, info.Pages[0].Height, 1);
        Assert.Equal(792, info.Pages[1].Width, 1);
        Assert.Equal(612, info.Pages[1].Height, 1);

        using PdfDocument pdf = PdfDocument.Open(bytes);
        var firstPage = pdf.GetPage(1);
        var secondPage = pdf.GetPage(2);
        Assert.Contains("NarrowMarginMarker", firstPage.Text);
        Assert.Contains("WideMarginMarker", secondPage.Text);

        double narrowX = FindWordStartX(firstPage, "NarrowMarginMarker");
        double wideX = FindWordStartX(secondPage, "WideMarginMarker");
        Assert.InRange(narrowX, 35D, 48D);
        Assert.InRange(wideX, 140D, 156D);
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
        using PdfDocument pdf = PdfDocument.Open(bytes);
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
        using PdfDocument pdf = PdfDocument.Open(bytes);
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
        using PdfDocument pdf = PdfDocument.Open(bytes);
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
        using PdfDocument pdf = PdfDocument.Open(bytes);
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
        using PdfDocument pdf = PdfDocument.Open(bytes);
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
        using PdfDocument pdf = PdfDocument.Open(bytes);
        var page = pdf.GetPage(1);
        string text = page.Text;
        Assert.Contains("InlineLeftColumnMarker", text);
        Assert.Contains("InlineRightColumnMarker", text);

        double leftX = FindWordStartX(page, "InlineLeftColumnMarker");
        double rightX = FindWordStartX(page, "InlineRightColumnMarker");
        Assert.InRange(leftX, 35D, 48D);
        Assert.True(rightX > leftX + 250D, $"Expected text after an inline Word column break to render in the next section column. Left x: {leftX:0.##}, right x: {rightX:0.##}.");
    }

    [Fact]
    public void SaveAsPdf_OfficeIMOEngine_Renders_First_And_Even_HeaderFooter_Variants() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfNativeHeaderFooterVariants.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeHeaderFooterVariants.pdf");

        using (WordDocument document = WordDocument.Create(docPath)) {
            document.AddHeadersAndFooters();
            document.DifferentFirstPage = true;
            document.DifferentOddAndEvenPages = true;

            RequireSectionHeader(document, 0, HeaderFooterValues.Default).AddParagraph("Native Odd Header");
            RequireSectionFooter(document, 0, HeaderFooterValues.Default).AddParagraph("Native Odd Footer");
            RequireSectionHeader(document, 0, HeaderFooterValues.First).AddParagraph("Native First Header");
            RequireSectionFooter(document, 0, HeaderFooterValues.First).AddParagraph("Native First Footer");
            RequireSectionHeader(document, 0, HeaderFooterValues.Even).AddParagraph("Native Even Header");
            RequireSectionFooter(document, 0, HeaderFooterValues.Even).AddParagraph("Native Even Footer");

            for (int i = 0; i < 240; i++) {
                document.AddParagraph($"Native variant paragraph {i}");
            }

            document.Save();
            document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                IncludePageNumbers = false
            });
        }

        Assert.True(File.Exists(pdfPath));
        using (PdfDocument pdf = PdfDocument.Open(pdfPath)) {
            Assert.True(pdf.NumberOfPages >= 3);
            string firstPageText = pdf.GetPage(1).Text;
            string secondPageText = pdf.GetPage(2).Text;
            string thirdPageText = pdf.GetPage(3).Text;

            Assert.Contains("Native First Header", firstPageText);
            Assert.Contains("Native First Footer", firstPageText);
            Assert.Contains("Native Even Header", secondPageText);
            Assert.Contains("Native Even Footer", secondPageText);
            Assert.Contains("Native Odd Header", thirdPageText);
            Assert.Contains("Native Odd Footer", thirdPageText);
        }
    }

    [Fact]
    public void SaveAsPdf_OfficeIMOEngine_Preserves_Blank_First_And_Even_HeaderFooter_Variants() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfNativeBlankHeaderFooterVariants.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeBlankHeaderFooterVariants.pdf");

        using (WordDocument document = WordDocument.Create(docPath)) {
            document.AddHeadersAndFooters();
            document.DifferentFirstPage = true;
            document.DifferentOddAndEvenPages = true;

            RequireSectionHeader(document, 0, HeaderFooterValues.Default).AddParagraph("Native Odd Header");
            RequireSectionFooter(document, 0, HeaderFooterValues.Default).AddParagraph("Native Odd Footer");

            for (int i = 0; i < 240; i++) {
                document.AddParagraph("Native blank variant body");
            }

            document.Save();
            document.SaveAsPdf(pdfPath);
        }

        using (PdfDocument pdf = PdfDocument.Open(pdfPath)) {
            Assert.True(pdf.NumberOfPages >= 3);
            string firstPageText = pdf.GetPage(1).Text;
            string secondPageText = pdf.GetPage(2).Text;
            string thirdPageText = pdf.GetPage(3).Text;

            Assert.DoesNotContain("Native Odd Header", firstPageText);
            Assert.DoesNotContain("Native Odd Footer", firstPageText);
            Assert.DoesNotContain("Native Odd Header", secondPageText);
            Assert.DoesNotContain("Native Odd Footer", secondPageText);
            Assert.Contains("Native Odd Header", thirdPageText);
            Assert.Contains("Native Odd Footer", thirdPageText);
        }
    }

    [Fact]
    public void SaveAsPdf_OfficeIMOEngine_Records_Warnings_For_Unsupported_HeaderFooter_Content() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfNativeHeaderFooterWarnings.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeHeaderFooterWarnings.pdf");
        var options = new PdfSaveOptions {
            IncludePageNumbers = false
        };
        options.Warnings.Add(new PdfExportWarning("Stale", "test", "This should be cleared before export."));

        using (WordDocument document = WordDocument.Create(docPath)) {
            document.AddHeadersAndFooters();
            WordHeader header = RequireSectionHeader(document, 0, HeaderFooterValues.Default);
            header.AddParagraph("Native warning header text");
            header.AddParagraph().AddTextBox(string.Empty, WrapTextImage.Square);

            WordTable footerTable = RequireSectionFooter(document, 0, HeaderFooterValues.Default).AddTable(1, 1, WordTableStyle.TableNormal);
            footerTable.Rows[0].Cells[0].Paragraphs[0].AddTextBox(string.Empty, WrapTextImage.Square);

            document.AddParagraph("Native warning body text");
            document.Save();
            document.SaveAsPdf(pdfPath, options);
        }

        Assert.DoesNotContain(options.Warnings, warning => warning.Code == "Stale");
        Assert.Contains(options.Warnings, warning =>
            warning.Code == "NativeHeaderFooterTextBoxUnsupported" &&
            warning.Source == "default header");
        Assert.Contains(options.Warnings, warning =>
            warning.Code == "NativeHeaderFooterTextBoxUnsupported" &&
            warning.Source == "default footer table");

        using PdfDocument pdf = PdfDocument.Open(pdfPath);
        string text = pdf.GetPage(1).Text;
        Assert.Contains("Native warning header text", text);
        Assert.Contains("Native warning body text", text);
    }

    [Fact]
    public void SaveAsPdf_OfficeIMOEngine_Maps_HeaderFooter_TextBoxes() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfNativeHeaderFooterTextBoxes.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeHeaderFooterTextBoxes.pdf");
        var options = new PdfSaveOptions {
            IncludePageNumbers = false
        };

        using (WordDocument document = WordDocument.Create(docPath)) {
            document.AddHeadersAndFooters();
            WordTextBox headerTextBox = RequireSectionHeader(document, 0, HeaderFooterValues.Default).AddTextBox("Native header text box");
            headerTextBox.HorizontalAlignment = WordHorizontalAlignmentValues.Center;

            WordParagraph footerParagraph = RequireSectionFooter(document, 0, HeaderFooterValues.Default).AddParagraph();
            WordTextBox footerTextBox = footerParagraph.AddTextBox("Native footer text box", WrapTextImage.Square);
            footerTextBox.HorizontalAlignment = WordHorizontalAlignmentValues.Right;

            document.AddParagraph("Native text box body");
            document.Save();
            document.SaveAsPdf(pdfPath, options);
        }

        Assert.DoesNotContain(options.Warnings, warning => warning.Code == "NativeHeaderFooterTextBoxUnsupported");

        using PdfDocument pdf = PdfDocument.Open(pdfPath);
        string text = pdf.GetPage(1).Text;
        Assert.Contains("Native header text box", text);
        Assert.Contains("Native footer text box", text);
        Assert.Contains("Native text box body", text);
    }

    [Fact]
    public void SaveAsPdf_OfficeIMOEngine_Maps_HeaderFooter_Shapes() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfNativeHeaderFooterShapes.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeHeaderFooterShapes.pdf");
        var options = new PdfSaveOptions {
            IncludePageNumbers = false
        };

        using (WordDocument document = WordDocument.Create(docPath)) {
            document.AddHeadersAndFooters();
            WordHeader header = RequireSectionHeader(document, 0, HeaderFooterValues.Default);
            header.AddShape(ShapeType.Rectangle, 36, 16, "#99ccff", "#003366", 1.5);

            WordFooter footer = RequireSectionFooter(document, 0, HeaderFooterValues.Default);
            WordParagraph footerParagraph = footer.AddParagraph();
            footerParagraph.ParagraphAlignment = JustificationValues.Right;
            footerParagraph.AddShape(ShapeType.Rectangle, 34, 14, "#ffe699", "#663300", 1.25);

            document.AddParagraph("Native header footer shape body");
            document.Save();
            document.SaveAsPdf(pdfPath, options);
        }

        Assert.DoesNotContain(options.Warnings, warning => warning.Code == "NativeHeaderFooterShapeUnsupported");

        using PdfDocument pdf = PdfDocument.Open(pdfPath);
        string text = pdf.GetPage(1).Text;
        Assert.Contains("Native header footer shape body", text);

        string content = Encoding.ASCII.GetString(File.ReadAllBytes(pdfPath));
        Assert.Contains("0.6 0.8 1 rg", content);
        Assert.Contains("1 0.902 0.6 rg", content);
        Assert.Contains("0 0.2 0.4 RG", content);
        Assert.Contains("0.4 0.2 0 RG", content);
        Assert.Contains(" re B", content);
    }

    [Fact]
    public void SaveAsPdf_OfficeIMOEngine_Maps_HeaderFooter_Images() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfNativeHeaderFooterImages.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeHeaderFooterImages.pdf");
        string imagePath = Path.Combine(_directoryWithImages, "EvotecLogo.png");
        var options = new PdfSaveOptions {
            IncludePageNumbers = false
        };

        using (WordDocument document = WordDocument.Create(docPath)) {
            document.AddHeadersAndFooters();
            WordParagraph headerParagraph = RequireSectionHeader(document, 0, HeaderFooterValues.Default).AddParagraph("Native image header");
            headerParagraph.ParagraphAlignment = JustificationValues.Center;
            headerParagraph.AddImage(imagePath, 32, 32);

            WordTable footerTable = RequireSectionFooter(document, 0, HeaderFooterValues.Default).AddTable(1, 3, WordTableStyle.TableNormal);
            footerTable.Rows[0].Cells[2].Paragraphs[0].AddImage(imagePath, 32, 32);

            document.AddParagraph("Native header/footer image body");
            document.Save();
            document.SaveAsPdf(pdfPath, options);
        }

        Assert.DoesNotContain(options.Warnings, warning => warning.Code == "NativeHeaderFooterImageUnsupported");

        byte[] bytes = File.ReadAllBytes(pdfPath);
        string rawPdf = Encoding.ASCII.GetString(bytes);
        int imageObjectCount = rawPdf.Split(new[] { "/Subtype /Image" }, StringSplitOptions.None).Length - 1;
        Assert.True(imageObjectCount >= 2, "Expected native header and footer images to be emitted as image XObjects.");

        using PdfDocument pdf = PdfDocument.Open(bytes);
        string text = pdf.GetPage(1).Text;
        Assert.Contains("Native image header", text);
        Assert.Contains("Native header/footer image body", text);
        Assert.DoesNotContain("Page 1", text, StringComparison.Ordinal);
    }

    [Fact]
    public void SaveAsPdf_OfficeIMOEngine_Maps_HeaderFooter_PictureControls_To_Images() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfNativeHeaderFooterPictureControls.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeHeaderFooterPictureControls.pdf");
        string imagePath = Path.Combine(_directoryWithImages, "EvotecLogo.png");
        var options = new PdfSaveOptions {
            IncludePageNumbers = false
        };

        using (WordDocument document = WordDocument.Create(docPath)) {
            document.AddHeadersAndFooters();
            WordParagraph headerParagraph = RequireSectionHeader(document, 0, HeaderFooterValues.Default).AddParagraph("Native picture-control header");
            headerParagraph.ParagraphAlignment = JustificationValues.Center;
            headerParagraph.AddPictureControl(imagePath, 32, 32, "Header Logo", "HeaderLogo");

            WordTable footerTable = RequireSectionFooter(document, 0, HeaderFooterValues.Default).AddTable(1, 3, WordTableStyle.TableNormal);
            footerTable.Rows[0].Cells[2].Paragraphs[0].AddPictureControl(imagePath, 32, 32, "Footer Logo", "FooterLogo");

            document.AddParagraph("Native header/footer picture-control body");
            document.Save();
            document.SaveAsPdf(pdfPath, options);
        }

        Assert.DoesNotContain(options.Warnings, warning => warning.Code == "NativeHeaderFooterContentControlUnsupported");

        byte[] bytes = File.ReadAllBytes(pdfPath);
        string rawPdf = Encoding.ASCII.GetString(bytes);
        int imageObjectCount = rawPdf.Split(new[] { "/Subtype /Image" }, StringSplitOptions.None).Length - 1;
        Assert.True(imageObjectCount >= 2, "Expected native header and footer picture controls to be emitted as image XObjects.");

        using PdfDocument pdf = PdfDocument.Open(bytes);
        string text = pdf.GetPage(1).Text;
        Assert.Contains("Native picture-control header", text);
        Assert.Contains("Native header/footer picture-control body", text);
    }

    [Fact]
    public void SaveAsPdf_OfficeIMOEngine_Maps_HeaderFooter_RepeatingSections_To_Text_Items() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfNativeHeaderFooterRepeatingSections.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeHeaderFooterRepeatingSections.pdf");
        var options = new PdfSaveOptions {
            IncludePageNumbers = false
        };

        using (WordDocument document = WordDocument.Create(docPath)) {
            document.AddHeadersAndFooters();
            WordRepeatingSection headerRepeating = RequireSectionHeader(document, 0, HeaderFooterValues.Default)
                .AddParagraph("Native header repeating section: ")
                .AddRepeatingSection("HeaderTasks", "HeaderTasks", "HeaderTasksTag");
            headerRepeating.SetTextItems(new[] { "Header item one", "Header item two" });

            WordTable footerTable = RequireSectionFooter(document, 0, HeaderFooterValues.Default).AddTable(1, 3, WordTableStyle.TableNormal);
            WordRepeatingSection footerRepeating = footerTable.Rows[0].Cells[2].Paragraphs[0]
                .AddRepeatingSection("FooterTasks", "FooterTasks", "FooterTasksTag");
            footerRepeating.SetTextItems(new[] { "Footer item one", "Footer item two" });

            document.AddParagraph("Native header/footer repeating-section body");
            document.Save();
            document.SaveAsPdf(pdfPath, options);
        }

        Assert.DoesNotContain(options.Warnings, warning => warning.Code == "NativeHeaderFooterContentControlUnsupported");

        using PdfDocument pdf = PdfDocument.Open(pdfPath);
        string text = pdf.GetPage(1).Text;
        Assert.Contains("Native header repeating section:", text);
        Assert.Contains("Header item one", text);
        Assert.Contains("Header item two", text);
        Assert.Contains("Footer item one", text);
        Assert.Contains("Footer item two", text);
        Assert.Contains("Native header/footer repeating-section body", text);
    }

    [Fact]
    public void SaveAsPdf_OfficeIMOEngine_Maps_HeaderFooter_FormControls_To_Static_Text() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfNativeHeaderFooterFormControls.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeHeaderFooterFormControls.pdf");
        var options = new PdfSaveOptions {
            IncludePageNumbers = false
        };

        using (WordDocument document = WordDocument.Create(docPath)) {
            document.AddHeadersAndFooters();
            WordHeader header = RequireSectionHeader(document, 0, HeaderFooterValues.Default);
            header.AddParagraph("Native header approval: ").AddCheckBox(true, "Header Approval", "HeaderApproval");
            header.AddParagraph("Native header due: ").AddDatePicker(new DateTime(2026, 5, 31), "Header Due", "HeaderDue");

            WordTable footerTable = RequireSectionFooter(document, 0, HeaderFooterValues.Default).AddTable(1, 3, WordTableStyle.TableNormal);
            footerTable.Rows[0].Cells[0].Paragraphs[0].Text = "Native footer region: ";
            footerTable.Rows[0].Cells[0].Paragraphs[0].AddDropDownList(new[] { "North", "South" }, "Footer Region", "FooterRegion");
            footerTable.Rows[0].Cells[2].Paragraphs[0].Text = "Native footer status: ";
            footerTable.Rows[0].Cells[2].Paragraphs[0].AddComboBox(new[] { "Red", "Blue" }, "Footer Status", "FooterStatus", defaultValue: "Blue");

            document.AddParagraph("Native header/footer form-control body");
            document.Save();
            document.SaveAsPdf(pdfPath, options);
        }

        Assert.DoesNotContain(options.Warnings, warning => warning.Code == "NativeHeaderFooterContentControlUnsupported");

        byte[] bytes = File.ReadAllBytes(pdfPath);
        Assert.Empty(PdfCore.PdfInspector.Inspect(bytes).FormFields);

        using PdfDocument pdf = PdfDocument.Open(bytes);
        string text = pdf.GetPage(1).Text;
        Assert.Contains("Native header approval:", text);
        Assert.Contains("[x]", text);
        Assert.Contains("Native header due:", text);
        Assert.Contains("2026-05-31", text);
        Assert.Contains("Native footer region:", text);
        Assert.Contains("North", text);
        Assert.Contains("Native footer status:", text);
        Assert.Contains("Blue", text);
        Assert.Contains("Native header/footer form-control body", text);
    }

    [Fact]
    public void SaveAsPdf_OfficeIMOEngine_Maps_Simple_Equations_To_Static_Text() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfNativeSimpleEquations.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeSimpleEquations.pdf");
        const string headerOmml = "<m:oMathPara xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\"><m:oMath><m:r><m:t>h=2</m:t></m:r></m:oMath></m:oMathPara>";
        const string bodyOmml = "<m:oMathPara xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\"><m:oMath><m:r><m:t>b=3</m:t></m:r></m:oMath></m:oMathPara>";
        const string tableOmml = "<m:oMathPara xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\"><m:oMath><m:r><m:t>c=4</m:t></m:r></m:oMath></m:oMathPara>";
        var options = new PdfSaveOptions {
            IncludePageNumbers = false
        };

        using (WordDocument document = WordDocument.Create(docPath)) {
            document.AddHeadersAndFooters();
            RequireSectionHeader(document, 0, HeaderFooterValues.Default)
                .AddParagraph("Native header equation:")
                .AddEquation(headerOmml);

            document.AddParagraph("Native body equation:").AddEquation(bodyOmml);

            WordTable table = document.AddTable(1, 1, WordTableStyle.TableNormal);
            table.Rows[0].Cells[0].Paragraphs[0].Text = "Native table equation:";
            table.Rows[0].Cells[0].Paragraphs[0].AddEquation(tableOmml);

            document.Save();
            document.SaveAsPdf(pdfPath, options);
        }

        Assert.DoesNotContain(options.Warnings, warning => warning.Code == "NativeHeaderFooterEquationUnsupported");
        Assert.DoesNotContain(options.Warnings, warning => warning.Code == "NativeBodyEquationUnsupported");

        string text = PdfCore.PdfTextExtractor.ExtractAllText(pdfPath);
        Assert.Contains("Native header equation:", text);
        Assert.Contains("h=2", text);
        Assert.Contains("Native body equation:", text);
        Assert.Contains("b=3", text);
        string normalizedText = NormalizePdfText(text);
        Assert.Contains("Native table equation:", normalizedText);
        Assert.Contains("c=4", normalizedText);
    }

    [Fact]
    public void SaveAsPdf_OfficeIMOEngine_Records_Warnings_For_Unsupported_Body_Content() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfNativeBodyWarnings.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeBodyWarnings.pdf");
        var options = new PdfSaveOptions {
            IncludePageNumbers = false
        };

        using (WordDocument document = WordDocument.Create(docPath)) {
            const string omml = "<m:oMathPara xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\"><m:oMath><m:r><m:t>x=1</m:t></m:r></m:oMath></m:oMathPara>";
            document.AddParagraph("Native body control text").AddDropDownList(new[] { "One", "Two" }, "BodyControl", "BodyControlTag");

            WordTable table = document.AddTable(1, 1, WordTableStyle.TableNormal);
            table.Rows[0].Cells[0].Paragraphs[0].Text = "NativeTableControlText";
            table.Rows[0].Cells[0].Paragraphs[0].AddEquation(omml);

            document.AddEmbeddedFragment("<html><body><p>Embedded body fragment</p></body></html>", WordAlternativeFormatImportPartType.Html);
            document.Save();
            document.SaveAsPdf(pdfPath, options);
        }

        Assert.DoesNotContain(options.Warnings, warning =>
            warning.Code == "NativeBodyContentControlUnsupported" &&
            warning.Source == "body paragraph");
        Assert.DoesNotContain(options.Warnings, warning =>
            warning.Code == "NativeBodyEquationUnsupported" &&
            warning.Source == "body table");
        Assert.Contains(options.Warnings, warning =>
            warning.Code == "NativeBodyEmbeddedDocumentUnsupported" &&
            warning.Source == "body");

        using PdfDocument pdf = PdfDocument.Open(pdfPath);
        string text = pdf.GetPage(1).Text;
        Assert.Contains("Native body control text", text);
        Assert.Contains("NativeTableControlText", text);
        Assert.Contains("x=1", text);
    }

    [Fact]
    public void SaveAsPdf_OfficeIMOEngine_Maps_Simple_Text_ContentControls() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfNativeSimpleTextContentControls.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeSimpleTextContentControls.pdf");
        var options = new PdfSaveOptions {
            IncludePageNumbers = false
        };

        using (WordDocument document = WordDocument.Create(docPath)) {
            document.AddHeadersAndFooters();
            RequireSectionHeader(document, 0, HeaderFooterValues.Default)
                .AddParagraph("Native header content control: ")
                .AddStructuredDocumentTag("Header control", "HeaderAlias", "HeaderTag");

            document.AddParagraph("Native body content control: ")
                .AddStructuredDocumentTag("Body control", "BodyAlias", "BodyTag");

            WordTable table = document.AddTable(1, 1, WordTableStyle.TableNormal);
            table.Rows[0].Cells[0].Paragraphs[0].Text = "Native cell content control: ";
            table.Rows[0].Cells[0].Paragraphs[0].AddStructuredDocumentTag("Cell control", "CellAlias", "CellTag");

            document.Save();
            document.SaveAsPdf(pdfPath, options);
        }

        Assert.DoesNotContain(options.Warnings, warning => warning.Code == "NativeBodyContentControlUnsupported");
        Assert.DoesNotContain(options.Warnings, warning => warning.Code == "NativeHeaderFooterContentControlUnsupported");

        using PdfDocument pdf = PdfDocument.Open(pdfPath);
        string text = pdf.GetPage(1).Text;
        Assert.Contains("Native header content control:", text);
        Assert.Contains("Header control", text);
        Assert.Contains("Native body content control:", text);
        Assert.Contains("Body control", text);
        Assert.Contains("Native cell content", text);
        Assert.Contains("Cell control", text);
    }

    [Fact]
    public void SaveAsPdf_OfficeIMOEngine_Maps_Body_DropDown_ComboBox_And_DatePicker_To_AcroForm_Fields() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfNativeBodyContentControlFormFields.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeBodyContentControlFormFields.pdf");
        var options = new PdfSaveOptions {
            IncludePageNumbers = false
        };

        using (WordDocument document = WordDocument.Create(docPath)) {
            document.AddParagraph("Native dropdown: ").AddDropDownList(new[] { "Poland", "Germany" }, "Country", "CountryTag");
            document.AddParagraph("Native combo: ").AddComboBox(new[] { "Red", "Blue" }, "Color", "ColorTag", defaultValue: "Blue");
            document.AddParagraph("Native date: ").AddDatePicker(new DateTime(2026, 5, 29), "Due Date", "DueDateTag");
            document.Save();
            document.SaveAsPdf(pdfPath, options);
        }

        Assert.DoesNotContain(options.Warnings, warning =>
            warning.Code == "NativeBodyContentControlUnsupported" &&
            warning.Source == "body paragraph");

        byte[] bytes = File.ReadAllBytes(pdfPath);
        PdfCore.PdfDocumentInfo info = PdfCore.PdfInspector.Inspect(bytes);
        Assert.Equal(3, info.FormFields.Count);

        PdfCore.PdfFormField country = Assert.Single(info.FormFields, field => field.Name == "CountryTag");
        Assert.Equal(PdfCore.PdfFormFieldKind.Choice, country.Kind);
        Assert.True(country.IsCombo);
        Assert.Equal("Poland", country.Value);
        Assert.Equal(new[] { "Poland", "Germany" }, country.Options.Select(option => option.ExportValue).ToArray());

        PdfCore.PdfFormField color = Assert.Single(info.FormFields, field => field.Name == "ColorTag");
        Assert.Equal(PdfCore.PdfFormFieldKind.Choice, color.Kind);
        Assert.True(color.IsCombo);
        Assert.Equal("Blue", color.Value);

        PdfCore.PdfFormField dueDate = Assert.Single(info.FormFields, field => field.Name == "DueDateTag");
        Assert.Equal(PdfCore.PdfFormFieldKind.Text, dueDate.Kind);
        Assert.Equal("2026-05-29", dueDate.Value);

        using PdfDocument pdf = PdfDocument.Open(bytes);
        string text = pdf.GetPage(1).Text;
        Assert.Contains("Native dropdown:", text);
        Assert.Contains("Native combo:", text);
        Assert.Contains("Native date:", text);
    }

    [Fact]
    public void SaveAsPdf_OfficeIMOEngine_Maps_Table_Cell_DropDown_ComboBox_And_DatePicker_To_AcroForm_Fields() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfNativeTableCellContentControlFormFields.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeTableCellContentControlFormFields.pdf");
        var options = new PdfSaveOptions {
            IncludePageNumbers = false
        };

        using (WordDocument document = WordDocument.Create(docPath)) {
            WordTable table = document.AddTable(1, 1, WordTableStyle.TableGrid);
            table.WidthType = TableWidthUnitValues.Dxa;
            table.Width = 7200;
            table.ColumnWidth = new[] { 7200 }.ToList();
            table.ColumnWidthType = TableWidthUnitValues.Dxa;
            WordParagraph paragraph = table.Rows[0].Cells[0].Paragraphs[0];
            paragraph.Text = "Native table controls:";
            paragraph.AddDropDownList(new[] { "Poland", "Germany" }, "Cell Country", "CellCountry");
            paragraph.AddComboBox(new[] { "Red", "Blue" }, "Cell Color", "CellColor", defaultValue: "Blue");
            paragraph.AddDatePicker(new DateTime(2026, 5, 31), "Cell Due Date", "CellDueDate");
            document.Save();
            document.SaveAsPdf(pdfPath, options);
        }

        Assert.DoesNotContain(options.Warnings, warning =>
            warning.Code == "NativeBodyContentControlUnsupported" &&
            warning.Source == "body table");

        byte[] bytes = File.ReadAllBytes(pdfPath);
        PdfCore.PdfDocumentInfo info = PdfCore.PdfInspector.Inspect(bytes);
        Assert.Equal(3, info.FormFields.Count);
        Assert.Contains(info.FormFields, field => field.Name == "CellCountry" && field.IsChoiceField && field.Value == "Poland");
        Assert.Contains(info.FormFields, field => field.Name == "CellColor" && field.IsChoiceField && field.Value == "Blue");
        Assert.Contains(info.FormFields, field => field.Name == "CellDueDate" && field.IsTextField && field.Value == "2026-05-31");
        Assert.True(info.Pages[0].HasFormWidgets);

        using PdfDocument pdf = PdfDocument.Open(bytes);
        Assert.Contains("Native table controls:", pdf.GetPage(1).Text);
    }

    [Fact]
    public void SaveAsPdf_OfficeIMOEngine_Maps_Body_RepeatingSection_To_Text_Items() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfNativeBodyRepeatingSection.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeBodyRepeatingSection.pdf");
        var options = new PdfSaveOptions {
            IncludePageNumbers = false
        };

        using (WordDocument document = WordDocument.Create(docPath)) {
            document.AddParagraph("Native repeating section:");
            WordRepeatingSection repeatingSection = document.AddParagraph()
                .AddRepeatingSection("Tasks", "Tasks", "TasksTag");
            repeatingSection.SetTextItems(new[] { "Plan roadmap slice", "Validate native PDF output" });
            document.Save();
            document.SaveAsPdf(pdfPath, options);
        }

        Assert.DoesNotContain(options.Warnings, warning =>
            warning.Code == "NativeBodyContentControlUnsupported" &&
            warning.Source == "body paragraph");

        using PdfDocument pdf = PdfDocument.Open(pdfPath);
        string text = pdf.GetPage(1).Text;
        Assert.Contains("Native repeating section:", text);
        Assert.Contains("Plan roadmap slice", text);
        Assert.Contains("Validate native PDF output", text);
    }

    [Fact]
    public void SaveAsPdf_OfficeIMOEngine_Maps_Table_Cell_RepeatingSection_To_Text_Items() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfNativeTableCellRepeatingSection.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeTableCellRepeatingSection.pdf");
        var options = new PdfSaveOptions {
            IncludePageNumbers = false
        };

        using (WordDocument document = WordDocument.Create(docPath)) {
            WordTable table = document.AddTable(1, 1, WordTableStyle.TableGrid);
            table.WidthType = TableWidthUnitValues.Dxa;
            table.Width = 7200;
            table.ColumnWidth = new[] { 7200 }.ToList();
            table.ColumnWidthType = TableWidthUnitValues.Dxa;
            table.Rows[0].Cells[0].Paragraphs[0].Text = "Table tasks";
            WordRepeatingSection repeatingSection = table.Rows[0].Cells[0].Paragraphs[0]
                .AddRepeatingSection("Tasks", "Tasks", "TasksTag");
            repeatingSection.SetTextItems(new[] { "Render cell item", "Keep table warnings clean" });
            document.Save();
            document.SaveAsPdf(pdfPath, options);
        }

        Assert.DoesNotContain(options.Warnings, warning =>
            warning.Code == "NativeBodyContentControlUnsupported" &&
            warning.Source == "body table");

        using PdfDocument pdf = PdfDocument.Open(pdfPath);
        string text = pdf.GetPage(1).Text;
        Assert.Contains("Table tasks", text);
        Assert.Contains("Render cell item", text);
        Assert.Contains("Keep table warnings clean", text);
    }

    [Fact]
    public void SaveAsPdf_OfficeIMOEngine_Maps_Body_CheckBox_To_AcroForm_Field() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfNativeBodyCheckBox.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeBodyCheckBox.pdf");
        var options = new PdfSaveOptions {
            IncludePageNumbers = false
        };

        using (WordDocument document = WordDocument.Create(docPath)) {
            document.AddParagraph("Accept native checkbox").AddCheckBox(true, "Accept Native", "AcceptNative");
            document.Save();
            document.SaveAsPdf(pdfPath, options);
        }

        Assert.DoesNotContain(options.Warnings, warning => warning.Code == "NativeBodyContentControlUnsupported");

        byte[] bytes = File.ReadAllBytes(pdfPath);
        PdfCore.PdfDocumentInfo info = PdfCore.PdfInspector.Inspect(bytes);
        PdfCore.PdfFormField field = Assert.Single(info.FormFields);
        Assert.Equal("AcceptNative", field.Name);
        Assert.Equal(PdfCore.PdfFormFieldKind.Button, field.Kind);
        Assert.True(field.IsCheckBox);
        Assert.Equal("Yes", field.Value);

        using PdfDocument pdf = PdfDocument.Open(bytes);
        Assert.Contains("Accept native checkbox", pdf.GetPage(1).Text);
    }

    [Fact]
    public void SaveAsPdf_OfficeIMOEngine_Maps_Table_Cell_CheckBox_To_AcroForm_Field() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfNativeTableCellCheckBox.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeTableCellCheckBox.pdf");
        var options = new PdfSaveOptions {
            IncludePageNumbers = false
        };

        using (WordDocument document = WordDocument.Create(docPath)) {
            WordTable table = document.AddTable(1, 1, WordTableStyle.TableGrid);
            table.Rows[0].Cells[0].Paragraphs[0].Text = "Table cell approval";
            table.Rows[0].Cells[0].Paragraphs[0].AddCheckBox(true, "Table Cell Approval", "TableCellApproval");
            document.Save();
            document.SaveAsPdf(pdfPath, options);
        }

        Assert.DoesNotContain(options.Warnings, warning =>
            warning.Code == "NativeBodyContentControlUnsupported" &&
            warning.Source == "body table");

        byte[] bytes = File.ReadAllBytes(pdfPath);
        PdfCore.PdfDocumentInfo info = PdfCore.PdfInspector.Inspect(bytes);
        PdfCore.PdfFormField field = Assert.Single(info.FormFields);
        Assert.Equal("TableCellApproval", field.Name);
        Assert.Equal(PdfCore.PdfFormFieldKind.Button, field.Kind);
        Assert.True(field.IsCheckBox);
        Assert.Equal("Yes", field.Value);
        Assert.True(info.Pages[0].HasFormWidgets);

        using PdfDocument pdf = PdfDocument.Open(bytes);
        Assert.Contains("Table cell approval", pdf.GetPage(1).Text);
    }

    private static double FindWordStartX(UglyToad.PdfPig.Content.Page page, string word) {
        var lines = page.Letters
            .Where(letter => !string.IsNullOrWhiteSpace(letter.Value))
            .GroupBy(letter => Math.Round(letter.StartBaseLine.Y, 1));

        foreach (var line in lines) {
            var ordered = line.OrderBy(letter => letter.StartBaseLine.X).ToList();
            string text = string.Concat(ordered.Select(letter => letter.Value));
            int index = text.IndexOf(word, StringComparison.Ordinal);
            if (index >= 0) {
                return ordered[index].StartBaseLine.X;
            }
        }

        throw new InvalidOperationException("Could not find word '" + word + "' in rendered PDF text.");
    }

    private static string NormalizePdfText(string text) =>
        string.Join(" ", text.Split(new[] { ' ', '\r', '\n', '\t', '\f' }, StringSplitOptions.RemoveEmptyEntries));
}

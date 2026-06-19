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
        using (PdfPigDocument pdf = PdfPigDocument.Open(pdfPath)) {
            string allText = string.Concat(pdf.GetPages().Select(page => page.Text));
            Assert.Contains("Native Default Header", allText);
            Assert.Contains("Native Default Footer", allText);
            Assert.Contains("Native body text", allText);
        }
    }

    [Fact]
    public void SaveAsPdf_OfficeIMOEngine_Maps_HeaderFooter_Explicit_Font_Families_To_Page_Text_Fonts() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfNativeHeaderFooterFonts.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeHeaderFooterFonts.pdf");

        using (WordDocument document = WordDocument.Create(docPath)) {
            document.AddHeadersAndFooters();
            WordParagraph header = RequireSectionHeader(document, 0, HeaderFooterValues.Default).AddParagraph("QSerifHeader");
            header.SetFontFamily("Georgia");
            WordParagraph footer = RequireSectionFooter(document, 0, HeaderFooterValues.Default).AddParagraph("XMonoFooter");
            footer.SetFontFamily("Courier New");
            document.AddParagraph("Plain body text");

            document.Save();
            document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                IncludePageNumbers = false
            });
        }

        Assert.True(File.Exists(pdfPath));
        using (PdfPigDocument pdf = PdfPigDocument.Open(pdfPath)) {
            var page = pdf.GetPage(1);
            var headerLetter = page.Letters.Single(letter => letter.Value == "Q");
            var footerLetter = page.Letters.Single(letter => letter.Value == "X");

            Assert.Contains("Times", headerLetter.FontName, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("Courier", footerLetter.FontName, StringComparison.OrdinalIgnoreCase);
        }

        string pdfContent = Encoding.ASCII.GetString(File.ReadAllBytes(pdfPath));
        Assert.Contains("/BaseFont /Times", pdfContent, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("/BaseFont /Courier", pdfContent, StringComparison.OrdinalIgnoreCase);
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
        using (PdfPigDocument pdf = PdfPigDocument.Open(pdfPath)) {
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
        using (PdfPigDocument pdf = PdfPigDocument.Open(pdfPath)) {
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
    public void SaveAsPdf_OfficeIMOEngine_Preserves_HeaderFooter_Table_Cell_Blank_Paragraphs() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfNativeTableHeaderFooterBlankParagraphs.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeTableHeaderFooterBlankParagraphs.pdf");

        using (WordDocument document = WordDocument.Create(docPath)) {
            document.AddHeadersAndFooters();

            WordTable headerTable = RequireSectionHeader(document, 0, HeaderFooterValues.Default).AddTable(1, 2, WordTableStyle.TableNormal);
            headerTable.Rows[0].Cells[0].AddParagraph("ShiftedHeaderLeft");
            headerTable.Rows[0].Cells[1].Paragraphs[0].Text = "PlainHeaderRight";

            WordTable footerTable = RequireSectionFooter(document, 0, HeaderFooterValues.Default).AddTable(1, 2, WordTableStyle.TableNormal);
            footerTable.Rows[0].Cells[0].AddParagraph("ShiftedFooterLeft");
            footerTable.Rows[0].Cells[1].Paragraphs[0].Text = "PlainFooterRight";

            document.AddParagraph("Native blank paragraph header footer body");
            document.Save();
            document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                IncludePageNumbers = false,
                PageSize = new PdfCore.PageSize(520, 340),
                Margins = PdfCore.PageMargins.Uniform(60)
            });
        }

        Assert.True(File.Exists(pdfPath));
        using (PdfPigDocument pdf = PdfPigDocument.Open(pdfPath)) {
            var page = pdf.GetPage(1);
            string text = page.Text;

            Assert.Contains("ShiftedHeaderLeft", text);
            Assert.Contains("PlainHeaderRight", text);
            Assert.Contains("ShiftedFooterLeft", text);
            Assert.Contains("PlainFooterRight", text);

            double shiftedHeaderY = FindWordStartY(page, "ShiftedHeaderLeft");
            double plainHeaderY = FindWordStartY(page, "PlainHeaderRight");
            double shiftedFooterY = FindWordStartY(page, "ShiftedFooterLeft");
            double plainFooterY = FindWordStartY(page, "PlainFooterRight");

            Assert.True(plainHeaderY > shiftedHeaderY + 5D, $"Expected a leading blank header table-cell paragraph to move text down. Plain y: {plainHeaderY:0.##}, shifted y: {shiftedHeaderY:0.##}.");
            Assert.True(plainFooterY > shiftedFooterY + 5D, $"Expected a leading blank footer table-cell paragraph to move text down. Plain y: {plainFooterY:0.##}, shifted y: {shiftedFooterY:0.##}.");
        }
    }

    [Fact]
    public void SaveAsPdf_OfficeIMOEngine_Reserves_Body_Clearance_For_Multiline_Header() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfNativeMultilineHeaderBodyClearance.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeMultilineHeaderBodyClearance.pdf");

        using (WordDocument document = WordDocument.Create(docPath)) {
            document.AddHeadersAndFooters();
            WordTable headerTable = RequireSectionHeader(document, 0, HeaderFooterValues.Default).AddTable(1, 2, WordTableStyle.TableNormal);
            headerTable.Rows[0].Cells[0].AddParagraph("ClearanceHeaderOne");
            headerTable.Rows[0].Cells[0].AddParagraph("ClearanceHeaderTwo");
            headerTable.Rows[0].Cells[1].AddParagraph("ClearanceHeaderRight");

            document.AddParagraph("ClearanceBodyHeading").SetStyle(WordParagraphStyles.Heading1);
            document.AddParagraph("Body text after a multiline Word header.");
            document.Save();
            document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                IncludePageNumbers = false,
                PageSize = PdfCore.PageSizes.Letter
            });
        }

        Assert.True(File.Exists(pdfPath));
        using (PdfPigDocument pdf = PdfPigDocument.Open(pdfPath)) {
            var page = pdf.GetPage(1);
            string text = page.Text;

            Assert.Contains("ClearanceHeaderTwo", text);
            Assert.Contains("ClearanceBodyHeading", text);

            double secondHeaderY = FindWordStartY(page, "ClearanceHeaderTwo");
            double bodyHeadingY = FindWordStartY(page, "ClearanceBodyHeading");

            Assert.True(secondHeaderY > bodyHeadingY + 20D, $"Expected body flow to start below the multiline Word header band. Header y: {secondHeaderY:0.##}, body y: {bodyHeadingY:0.##}.");
        }
    }

    [Fact]
    public void SaveAsPdf_OfficeIMOEngine_Preserves_HeaderFooter_Paragraph_Lines() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfNativeMultilineHeaderFooter.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeMultilineHeaderFooter.pdf");

        using (WordDocument document = WordDocument.Create(docPath)) {
            document.AddHeadersAndFooters();
            WordHeader header = RequireSectionHeader(document, 0, HeaderFooterValues.Default);
            header.AddParagraph("NativeHeaderLineOne");
            header.AddParagraph("NativeHeaderLineTwo");

            WordFooter footer = RequireSectionFooter(document, 0, HeaderFooterValues.Default);
            footer.AddParagraph("NativeFooterLineOne");
            footer.AddParagraph("NativeFooterLineTwo");

            document.AddParagraph("Native multiline header footer body");
            document.Save();
            document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                IncludePageNumbers = false,
                PageSize = new PdfCore.PageSize(440, 340),
                Margins = PdfCore.PageMargins.Uniform(60)
            });
        }

        Assert.True(File.Exists(pdfPath));
        using (PdfPigDocument pdf = PdfPigDocument.Open(pdfPath)) {
            var page = pdf.GetPage(1);
            string text = page.Text;

            Assert.Contains("NativeHeaderLineOne", text);
            Assert.Contains("NativeHeaderLineTwo", text);
            Assert.Contains("NativeFooterLineOne", text);
            Assert.Contains("NativeFooterLineTwo", text);
            Assert.DoesNotContain("NativeHeaderLineOne NativeHeaderLineTwo", text, StringComparison.Ordinal);
            Assert.DoesNotContain("NativeFooterLineOne NativeFooterLineTwo", text, StringComparison.Ordinal);

            double firstHeaderY = FindWordStartY(page, "NativeHeaderLineOne");
            double secondHeaderY = FindWordStartY(page, "NativeHeaderLineTwo");
            double firstFooterY = FindWordStartY(page, "NativeFooterLineOne");
            double secondFooterY = FindWordStartY(page, "NativeFooterLineTwo");

            Assert.True(firstHeaderY > secondHeaderY + 15D, $"Expected separate Word header paragraphs to keep paragraph spacing, not collapse to one line advance. First y: {firstHeaderY:0.##}, second y: {secondHeaderY:0.##}.");
            Assert.True(firstFooterY > secondFooterY + 15D, $"Expected separate Word footer paragraphs to keep paragraph spacing, not collapse to one line advance. First y: {firstFooterY:0.##}, second y: {secondFooterY:0.##}.");
        }
    }

    [Fact]
    public void SaveAsPdf_OfficeIMOEngine_Places_Multiline_Footer_Inside_Word_Footer_Band() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfNativeMultilineFooterPlacement.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeMultilineFooterPlacement.pdf");

        using (WordDocument document = WordDocument.Create(docPath)) {
            document.AddHeadersAndFooters();
            WordFooter footer = RequireSectionFooter(document, 0, HeaderFooterValues.Default);
            footer.AddParagraph("NativeFooterBandLineOne");
            footer.AddParagraph("NativeFooterBandLineTwo");

            document.AddParagraph("Native footer placement body");
            document.Save();
            document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                IncludePageNumbers = false,
                PageSize = PdfCore.PageSizes.Letter
            });
        }

        Assert.True(File.Exists(pdfPath));
        using (PdfPigDocument pdf = PdfPigDocument.Open(pdfPath)) {
            var page = pdf.GetPage(1);
            double firstFooterY = FindWordStartY(page, "NativeFooterBandLineOne");
            double secondFooterY = FindWordStartY(page, "NativeFooterBandLineTwo");

            Assert.InRange(firstFooterY, 67D, 73D);
            Assert.InRange(secondFooterY, 45D, 52D);
        }
    }

    [Fact]
    public void SaveAsPdf_OfficeIMOEngine_Reserves_Body_Clearance_For_Multiline_Footer_Margins() {
        using WordDocument document = WordDocument.Create(Path.Combine(_directoryWithFiles, "PdfNativeMultilineFooterMarginExpansion.docx"));
        document.Sections[0].Margins.Top = 600;
        document.Sections[0].Margins.Bottom = 600;
        document.Sections[0].Margins.Left = 600;
        document.Sections[0].Margins.Right = 600;
        document.AddHeadersAndFooters();
        WordFooter footer = RequireSectionFooter(document, 0, HeaderFooterValues.Default);
        footer.AddParagraph("ClearanceFooterOne");
        footer.AddParagraph("ClearanceFooterTwo");

        var method = typeof(WordPdfConverterExtensions).GetMethod(
            "GetNativeMargins",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static,
            binder: null,
            new[] { typeof(WordSection), typeof(PdfSaveOptions) },
            modifiers: null)!;
        PdfCore.PageMargins margins = Assert.IsType<PdfCore.PageMargins>(method.Invoke(null, new object?[] { document.Sections[0], null }));

        Assert.Equal(30D, margins.Top);
        Assert.Equal(30D, margins.Left);
        Assert.Equal(30D, margins.Right);
        Assert.InRange(margins.Bottom, 47D, 49D);
    }
}

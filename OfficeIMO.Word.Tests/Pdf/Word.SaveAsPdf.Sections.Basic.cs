using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using System;
using System.Globalization;
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
    public void SaveAsPdf_OfficeIMOEngine_Keeps_Compatible_Continuous_Sections_On_Same_Page() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfNativeContinuousSectionSamePage.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeContinuousSectionSamePage.pdf");

        using (WordDocument document = WordDocument.Create(docPath)) {
            document.AddParagraph("ContinuousSectionBefore");
            document.AddSection(SectionMarkValues.Continuous);
            document.AddParagraph("ContinuousSectionAfter");

            document.Save();
            document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                IncludePageNumbers = false,
                PageSize = new PdfCore.PageSize(400, 300),
                Margins = PdfCore.PageMargins.Uniform(50)
            });
        }

        using PdfPigDocument pdf = PdfPigDocument.Open(pdfPath);
        Assert.Equal(1, pdf.NumberOfPages);
        Assert.Contains("ContinuousSectionBefore", pdf.GetPage(1).Text);
        Assert.Contains("ContinuousSectionAfter", pdf.GetPage(1).Text);
    }

    [Fact]
    public void SaveAsPdf_OfficeIMOEngine_Keeps_NextPage_Sections_On_New_Page() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfNativeNextPageSectionBreak.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeNextPageSectionBreak.pdf");

        using (WordDocument document = WordDocument.Create(docPath)) {
            document.AddParagraph("NextPageSectionBefore");
            document.AddSection(SectionMarkValues.NextPage);
            document.AddParagraph("NextPageSectionAfter");

            document.Save();
            document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                IncludePageNumbers = false,
                PageSize = new PdfCore.PageSize(400, 300),
                Margins = PdfCore.PageMargins.Uniform(50)
            });
        }

        using PdfPigDocument pdf = PdfPigDocument.Open(pdfPath);
        Assert.Equal(2, pdf.NumberOfPages);
        Assert.Contains("NextPageSectionBefore", pdf.GetPage(1).Text);
        Assert.Contains("NextPageSectionAfter", pdf.GetPage(2).Text);
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
                IncludePageNumbers = false,
                ResourcePolicy = PdfCore.PdfResourcePolicy.CreateTrustedHost()
            });
        }

        Assert.True(File.Exists(pdfPath));
        using (PdfPigDocument pdf = PdfPigDocument.Open(pdfPath)) {
            var page = pdf.GetPage(1);
            var headerLetter = page.Letters.Single(letter => letter.Value == "Q");
            var footerLetter = page.Letters.Single(letter => letter.Value == "X");

            string expectedHeaderFont = PdfCore.PdfEmbeddedFontFamily.TryFromSystem("Georgia", out _) ? "Georgia" : "Times";
            Assert.Contains(expectedHeaderFont, headerLetter.FontName, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("Courier", footerLetter.FontName, StringComparison.OrdinalIgnoreCase);
        }

        string pdfContent = Encoding.ASCII.GetString(File.ReadAllBytes(pdfPath));
        string expectedHeaderBaseFont = PdfCore.PdfEmbeddedFontFamily.TryFromSystem("Georgia", out _) ? "/BaseFont /Georgia" : "/BaseFont /Times";
        Assert.Contains(expectedHeaderBaseFont, pdfContent, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("/BaseFont /Courier", pdfContent, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void SaveAsPdf_OfficeIMOEngine_Maps_HeaderFooter_Direct_And_Style_Font_Sizes_To_Page_Text() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfNativeHeaderFooterFontSizes.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeHeaderFooterFontSizes.pdf");

        using (WordDocument document = WordDocument.Create(docPath)) {
            const string footerStyleId = "NativeSizedFooter";
            Styles styles = document._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
            styles.Append(
                new Style(
                    new StyleName { Val = "Native Sized Footer" },
                    new StyleRunProperties(
                        new FontSize { Val = "32" }))
                {
                    Type = StyleValues.Paragraph,
                    StyleId = footerStyleId,
                    CustomStyle = true
                });

            document.AddHeadersAndFooters();
            WordParagraph header = RequireSectionHeader(document, 0, HeaderFooterValues.Default).AddParagraph("BigNativeHeader");
            header.FontSize = 18;
            WordParagraph footer = RequireSectionFooter(document, 0, HeaderFooterValues.Default).AddParagraph("StyledNativeFooter");
            footer.SetStyleId(footerStyleId);
            document.AddParagraph("Plain body text");

            document.Save();
            document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                IncludePageNumbers = false
            });
        }

        Assert.True(File.Exists(pdfPath));
        using (PdfPigDocument pdf = PdfPigDocument.Open(pdfPath)) {
            var page = pdf.GetPage(1);
            Assert.Contains("BigNativeHeader", page.Text);
            Assert.Contains("StyledNativeFooter", page.Text);

            var textLines = page.Letters
                .Where(letter => !string.IsNullOrWhiteSpace(letter.Value))
                .GroupBy(letter => Math.Round(letter.StartBaseLine.Y, 1))
                .Select(group => group.OrderBy(letter => letter.StartBaseLine.X).ToList())
                .ToList();
            var headerLetters = textLines.Single(line => string.Concat(line.Select(letter => letter.Value)).Contains("BigNativeHeader", StringComparison.Ordinal));
            var footerLetters = textLines.Single(line => string.Concat(line.Select(letter => letter.Value)).Contains("StyledNativeFooter", StringComparison.Ordinal));
            double headerSize = headerLetters.Average(letter => letter.PointSize);
            double footerSize = footerLetters.Average(letter => letter.PointSize);

            Assert.InRange(headerSize, 17.5D, 18.5D);
            Assert.InRange(footerSize, 15.5D, 16.5D);
        }
    }

    [Fact]
    public void SaveAsPdf_OfficeIMOEngine_Maps_HeaderFooter_Explicit_Text_Colors_To_Page_Text_Colors() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfNativeHeaderFooterColors.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeHeaderFooterColors.pdf");

        using (WordDocument document = WordDocument.Create(docPath)) {
            document.AddHeadersAndFooters();
            WordParagraph header = RequireSectionHeader(document, 0, HeaderFooterValues.Default).AddParagraph("RedNativeHeader");
            header.ColorHex = "FF0000";
            WordParagraph footer = RequireSectionFooter(document, 0, HeaderFooterValues.Default).AddParagraph("BlueNativeFooter");
            footer.ColorHex = "0000FF";
            document.AddParagraph("Plain body text");

            document.Save();
            document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                IncludePageNumbers = false
            });
        }

        byte[] bytes = File.ReadAllBytes(pdfPath);
        string content = ReadPdfPageContent(bytes);
        using (PdfPigDocument pdf = PdfPigDocument.Open(new MemoryStream(bytes))) {
            string text = pdf.GetPage(1).Text;
            Assert.Contains("RedNativeHeader", text);
            Assert.Contains("BlueNativeFooter", text);
        }

        Assert.Contains("1 0 0 rg", content, StringComparison.Ordinal);
        Assert.Contains("0 0 1 rg", content, StringComparison.Ordinal);
    }

    [Fact]
    public void SaveAsPdf_OfficeIMOEngine_Maps_HeaderFooter_Paragraph_Style_Font_And_Color_To_Page_Text() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfNativeHeaderFooterStyleText.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeHeaderFooterStyleText.pdf");

        using (WordDocument document = WordDocument.Create(docPath)) {
            const string headerStyleId = "NativeStyledHeader";
            const string footerStyleId = "NativeStyledFooter";
            Styles styles = document._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
            styles.Append(
                new Style(
                    new StyleName { Val = "Native Styled Header" },
                    new StyleRunProperties(
                        new RunFonts { Ascii = "Georgia", HighAnsi = "Georgia" }))
                {
                    Type = StyleValues.Paragraph,
                    StyleId = headerStyleId,
                    CustomStyle = true
                },
                new Style(
                    new StyleName { Val = "Native Styled Footer" },
                    new StyleRunProperties(
                        new Color { Val = "008000" }))
                {
                    Type = StyleValues.Paragraph,
                    StyleId = footerStyleId,
                    CustomStyle = true
                });

            document.AddHeadersAndFooters();
            WordParagraph header = RequireSectionHeader(document, 0, HeaderFooterValues.Default).AddParagraph("QStyledHeader");
            header.SetStyleId(headerStyleId);
            WordParagraph footer = RequireSectionFooter(document, 0, HeaderFooterValues.Default).AddParagraph("GreenStyledFooter");
            footer.SetStyleId(footerStyleId);
            document.AddParagraph("Plain body text");

            document.Save();
            document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                IncludePageNumbers = false,
                ResourcePolicy = PdfCore.PdfResourcePolicy.CreateTrustedHost()
            });
        }

        Assert.True(File.Exists(pdfPath));
        using (PdfPigDocument pdf = PdfPigDocument.Open(pdfPath)) {
            var page = pdf.GetPage(1);
            Assert.Contains("QStyledHeader", page.Text);
            Assert.Contains("GreenStyledFooter", page.Text);

            var headerLetter = page.Letters.Single(letter => letter.Value == "Q");
            string expectedHeaderFont = PdfCore.PdfEmbeddedFontFamily.TryFromSystem("Georgia", out _) ? "Georgia" : "Times";
            Assert.Contains(expectedHeaderFont, headerLetter.FontName, StringComparison.OrdinalIgnoreCase);
        }

        string content = ReadPdfPageContent(File.ReadAllBytes(pdfPath));
        Assert.Contains("0 0.502 0 rg", content, StringComparison.Ordinal);
    }

    [Fact]
    public void SaveAsPdf_OfficeIMOEngine_Maps_HeaderFooter_Character_Style_Font_And_Color_To_Page_Text() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfNativeHeaderFooterCharacterStyleText.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeHeaderFooterCharacterStyleText.pdf");

        using (WordDocument document = WordDocument.Create(docPath)) {
            const string headerStyleId = "NativeCharacterStyledHeader";
            const string footerStyleId = "NativeCharacterStyledFooter";
            Styles styles = document._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
            styles.Append(
                new Style(
                    new StyleName { Val = "Native Character Styled Header" },
                    new StyleRunProperties(
                        new RunFonts { Ascii = "Georgia", HighAnsi = "Georgia" }))
                {
                    Type = StyleValues.Character,
                    StyleId = headerStyleId,
                    CustomStyle = true
                },
                new Style(
                    new StyleName { Val = "Native Character Styled Footer" },
                    new StyleRunProperties(
                        new Color { Val = "008000" }))
                {
                    Type = StyleValues.Character,
                    StyleId = footerStyleId,
                    CustomStyle = true
                });

            document.AddHeadersAndFooters();
            WordParagraph header = RequireSectionHeader(document, 0, HeaderFooterValues.Default).AddParagraph();
            header.AddText("QCharStyledHeader").SetCharacterStyleId(headerStyleId);
            WordParagraph footer = RequireSectionFooter(document, 0, HeaderFooterValues.Default).AddParagraph();
            footer.AddText("GreenCharStyledFooter").SetCharacterStyleId(footerStyleId);
            document.AddParagraph("Plain body text");

            document.Save();
            document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                IncludePageNumbers = false,
                ResourcePolicy = PdfCore.PdfResourcePolicy.CreateTrustedHost()
            });
        }

        Assert.True(File.Exists(pdfPath));
        using (PdfPigDocument pdf = PdfPigDocument.Open(pdfPath)) {
            var page = pdf.GetPage(1);
            Assert.Contains("QCharStyledHeader", page.Text);
            Assert.Contains("GreenCharStyledFooter", page.Text);

            var headerLetter = page.Letters.Single(letter => letter.Value == "Q");
            string expectedHeaderFont = PdfCore.PdfEmbeddedFontFamily.TryFromSystem("Georgia", out _) ? "Georgia" : "Times";
            Assert.Contains(expectedHeaderFont, headerLetter.FontName, StringComparison.OrdinalIgnoreCase);
        }

        string content = ReadPdfPageContent(File.ReadAllBytes(pdfPath));
        Assert.Contains("0 0.502 0 rg", content, StringComparison.Ordinal);
    }

    [Fact]
    public void SaveAsPdf_OfficeIMOEngine_Maps_HeaderFooter_Bold_And_Italic_To_Page_Text_Fonts() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfNativeHeaderFooterEmphasis.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeHeaderFooterEmphasis.pdf");

        using (WordDocument document = WordDocument.Create(docPath)) {
            document.AddHeadersAndFooters();
            WordParagraph header = RequireSectionHeader(document, 0, HeaderFooterValues.Default).AddParagraph("QBoldSerifHeader");
            header.SetFontFamily("Georgia");
            header.Bold = true;
            WordParagraph footer = RequireSectionFooter(document, 0, HeaderFooterValues.Default).AddParagraph("XItalicMonoFooter");
            footer.SetFontFamily("Courier New");
            footer.Italic = true;
            document.AddParagraph("Plain body text");

            document.Save();
            document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                IncludePageNumbers = false,
                ResourcePolicy = PdfCore.PdfResourcePolicy.CreateTrustedHost()
            });
        }

        Assert.True(File.Exists(pdfPath));
        using (PdfPigDocument pdf = PdfPigDocument.Open(pdfPath)) {
            var page = pdf.GetPage(1);
            var headerLetter = page.Letters.Single(letter => letter.Value == "Q");
            var footerLetter = page.Letters.Single(letter => letter.Value == "X");

            string expectedHeaderFont = PdfCore.PdfEmbeddedFontFamily.TryFromSystem("Georgia", out _) ? "Georgia-Bold" : "Times-Bold";
            Assert.Contains(expectedHeaderFont, headerLetter.FontName, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("Courier", footerLetter.FontName, StringComparison.OrdinalIgnoreCase);
            Assert.True(
                footerLetter.FontName.Contains("Italic", StringComparison.OrdinalIgnoreCase) ||
                footerLetter.FontName.Contains("Oblique", StringComparison.OrdinalIgnoreCase),
                "Expected the explicit footer family to preserve italic emphasis.");
        }

        string pdfContent = Encoding.ASCII.GetString(File.ReadAllBytes(pdfPath));
        string expectedHeaderBaseFont = PdfCore.PdfEmbeddedFontFamily.TryFromSystem("Georgia", out _) ? "/BaseFont /Georgia-Bold" : "/BaseFont /Times-Bold";
        Assert.Contains(expectedHeaderBaseFont, pdfContent, StringComparison.OrdinalIgnoreCase);
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
    public void SaveAsPdf_OfficeIMOEngine_Reserves_Body_Clearance_For_Large_Multiline_Header_Font_Size() {
        string defaultDocPath = Path.Combine(_directoryWithFiles, "PdfNativeDefaultMultilineHeaderClearance.docx");
        string defaultPdfPath = Path.Combine(_directoryWithFiles, "PdfNativeDefaultMultilineHeaderClearance.pdf");
        string largeDocPath = Path.Combine(_directoryWithFiles, "PdfNativeLargeMultilineHeaderClearance.docx");
        string largePdfPath = Path.Combine(_directoryWithFiles, "PdfNativeLargeMultilineHeaderClearance.pdf");

        CreateMultilineHeaderDocument(defaultDocPath, defaultPdfPath, "DefaultHeaderClearanceBody", null);
        CreateMultilineHeaderDocument(largeDocPath, largePdfPath, "LargeHeaderClearanceBody", 24);

        using PdfPigDocument defaultPdf = PdfPigDocument.Open(defaultPdfPath);
        using PdfPigDocument largePdf = PdfPigDocument.Open(largePdfPath);
        double defaultBodyY = FindWordStartY(defaultPdf.GetPage(1), "DefaultHeaderClearanceBody");
        double largeBodyY = FindWordStartY(largePdf.GetPage(1), "LargeHeaderClearanceBody");

        Assert.True(defaultBodyY > largeBodyY + 25D, $"Expected a large multiline Word header font to reserve more body clearance. Default body y: {defaultBodyY:0.##}, large body y: {largeBodyY:0.##}.");

        static void CreateMultilineHeaderDocument(string docPath, string pdfPath, string bodyText, int? headerFontSize) {
            using WordDocument document = WordDocument.Create(docPath);
            document.AddHeadersAndFooters();
            WordHeader header = RequireSectionHeader(document, 0, HeaderFooterValues.Default);
            for (int index = 1; index <= 4; index++) {
                WordParagraph paragraph = header.AddParagraph("ClearanceHeaderLine" + index.ToString(System.Globalization.CultureInfo.InvariantCulture));
                if (headerFontSize.HasValue) {
                    paragraph.FontSize = headerFontSize.Value;
                }
            }

            document.AddParagraph(bodyText).SetStyle(WordParagraphStyles.Heading1);
            document.Save();
            document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                IncludePageNumbers = false,
                PageSize = PdfCore.PageSizes.Letter
            });
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

            Assert.InRange(firstFooterY, 74D, 77D);
            Assert.InRange(secondFooterY, 52D, 56D);
        }
    }

    [Fact]
    public void SaveAsPdf_OfficeIMOEngine_Places_Large_Multiline_Footer_Inside_Page_Bounds() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfNativeLargeMultilineFooterPlacement.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeLargeMultilineFooterPlacement.pdf");

        using (WordDocument document = WordDocument.Create(docPath)) {
            document.AddHeadersAndFooters();
            WordFooter footer = RequireSectionFooter(document, 0, HeaderFooterValues.Default);
            for (int index = 1; index <= 4; index++) {
                WordParagraph paragraph = footer.AddParagraph("LargeFooterLine" + index.ToString(CultureInfo.InvariantCulture));
                paragraph.FontSize = 24;
            }

            document.AddParagraph("Large footer placement body");
            document.Save();
            document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                IncludePageNumbers = false,
                PageSize = PdfCore.PageSizes.Letter
            });
        }

        Assert.True(File.Exists(pdfPath));
        using (PdfPigDocument pdf = PdfPigDocument.Open(pdfPath)) {
            var page = pdf.GetPage(1);
            double firstFooterY = FindWordStartY(page, "LargeFooterLine1");
            double secondFooterY = FindWordStartY(page, "LargeFooterLine2");
            double thirdFooterY = FindWordStartY(page, "LargeFooterLine3");
            double fourthFooterY = FindWordStartY(page, "LargeFooterLine4");

            Assert.True(firstFooterY > secondFooterY + 25D, $"Expected large footer line 1 above line 2. First: {firstFooterY:0.##}, second: {secondFooterY:0.##}.");
            Assert.True(secondFooterY > thirdFooterY + 25D, $"Expected large footer line 2 above line 3. Second: {secondFooterY:0.##}, third: {thirdFooterY:0.##}.");
            Assert.True(thirdFooterY > fourthFooterY + 25D, $"Expected large footer line 3 above line 4. Third: {thirdFooterY:0.##}, fourth: {fourthFooterY:0.##}.");
            Assert.True(fourthFooterY > 20D, $"Expected the final large Word footer line to stay inside the page bounds. Fourth line y: {fourthFooterY:0.##}.");
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
        Assert.InRange(margins.Bottom, 53D, 55D);
    }
}

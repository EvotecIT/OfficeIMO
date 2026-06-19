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
        using (PdfPigDocument pdf = PdfPigDocument.Open(pdfPath)) {
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
    public void SaveAsPdf_OfficeIMOEngine_Ignores_First_And_Even_HeaderFooter_Parts_When_Section_Flags_Are_Off() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfNativeStaleHeaderFooterVariants.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeStaleHeaderFooterVariants.pdf");

        using (WordDocument document = WordDocument.Create(docPath)) {
            document.AddHeadersAndFooters();
            document.DifferentFirstPage = true;
            document.DifferentOddAndEvenPages = true;

            RequireSectionHeader(document, 0, HeaderFooterValues.Default).AddParagraph("Native Default Header");
            RequireSectionFooter(document, 0, HeaderFooterValues.Default).AddParagraph("Native Default Footer");
            RequireSectionHeader(document, 0, HeaderFooterValues.First).AddParagraph("Native Stale First Header");
            RequireSectionFooter(document, 0, HeaderFooterValues.First).AddParagraph("Native Stale First Footer");
            RequireSectionHeader(document, 0, HeaderFooterValues.Even).AddParagraph("Native Stale Even Header");
            RequireSectionFooter(document, 0, HeaderFooterValues.Even).AddParagraph("Native Stale Even Footer");

            for (int i = 0; i < 240; i++) {
                document.AddParagraph($"Native stale variant paragraph {i}");
            }

            document.Save();
        }

        using (WordprocessingDocument package = WordprocessingDocument.Open(docPath, true)) {
            Settings? settings = package.MainDocumentPart!.DocumentSettingsPart!.Settings;
            settings?.RemoveAllChildren<EvenAndOddHeaders>();
            foreach (TitlePage titlePage in package.MainDocumentPart.Document.Body!.Descendants<TitlePage>().ToList()) {
                titlePage.Remove();
            }

            package.Save();
        }

        using (WordDocument document = WordDocument.Load(docPath)) {
            document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                IncludePageNumbers = false
            });
        }

        using (PdfPigDocument pdf = PdfPigDocument.Open(pdfPath)) {
            Assert.True(pdf.NumberOfPages >= 2);
            string allText = string.Concat(pdf.GetPages().Select(page => page.Text));
            Assert.Contains("Native Default Header", pdf.GetPage(1).Text);
            Assert.Contains("Native Default Footer", pdf.GetPage(1).Text);
            Assert.Contains("Native Default Header", pdf.GetPage(2).Text);
            Assert.Contains("Native Default Footer", pdf.GetPage(2).Text);
            Assert.DoesNotContain("Native Stale First Header", allText);
            Assert.DoesNotContain("Native Stale First Footer", allText);
            Assert.DoesNotContain("Native Stale Even Header", allText);
            Assert.DoesNotContain("Native Stale Even Footer", allText);
        }
    }

    [Fact]
    public void SaveAsPdf_OfficeIMOEngine_Ignores_Inactive_First_And_Even_Header_Watermarks() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfNativeStaleHeaderWatermarkVariants.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeStaleHeaderWatermarkVariants.pdf");

        using (WordDocument document = WordDocument.Create(docPath)) {
            document.AddHeadersAndFooters();
            document.DifferentFirstPage = true;
            document.DifferentOddAndEvenPages = true;

            RequireSectionHeader(document, 0, HeaderFooterValues.First)
                .AddWatermark(WordWatermarkStyle.Text, "STALE FIRST WATERMARK");
            RequireSectionHeader(document, 0, HeaderFooterValues.Even)
                .AddWatermark(WordWatermarkStyle.Text, "STALE EVEN WATERMARK");

            for (int i = 0; i < 160; i++) {
                document.AddParagraph($"Native stale watermark paragraph {i}");
            }

            document.Save();
        }

        using (WordprocessingDocument package = WordprocessingDocument.Open(docPath, true)) {
            Settings? settings = package.MainDocumentPart!.DocumentSettingsPart!.Settings;
            settings?.RemoveAllChildren<EvenAndOddHeaders>();
            foreach (TitlePage titlePage in package.MainDocumentPart.Document.Body!.Descendants<TitlePage>().ToList()) {
                titlePage.Remove();
            }

            package.Save();
        }

        using (WordDocument document = WordDocument.Load(docPath)) {
            document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                IncludePageNumbers = false
            });
        }

        using PdfPigDocument pdf = PdfPigDocument.Open(pdfPath);
        string allText = string.Concat(pdf.GetPages().Select(page => page.Text));

        Assert.DoesNotContain("STALE FIRST WATERMARK", allText);
        Assert.DoesNotContain("STALE EVEN WATERMARK", allText);
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

        using (PdfPigDocument pdf = PdfPigDocument.Open(pdfPath)) {
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
    public void SaveAsPdf_OfficeIMOEngine_Preserves_Configured_Header_And_Footer_Offsets() {
        using WordDocument document = WordDocument.Create();
        document.AddHeadersAndFooters();
        WordHeader header = RequireSectionHeader(document, 0, HeaderFooterValues.Default);
        header.AddParagraph("Offset Header Line One");
        header.AddParagraph("Offset Header Line Two");
        header.AddParagraph("Offset Header Line Three");
        RequireSectionFooter(document, 0, HeaderFooterValues.Default).AddParagraph("Offset Footer");
        document.AddParagraph("Offset body");

        byte[] lowOffsetBytes = document.SaveAsPdf(new PdfSaveOptions {
            IncludePageNumbers = false,
            PdfOptions = new PdfCore.PdfOptions {
                HeaderOffsetY = 6,
                FooterOffsetY = 8
            }
        });
        byte[] highOffsetBytes = document.SaveAsPdf(new PdfSaveOptions {
            IncludePageNumbers = false,
            PdfOptions = new PdfCore.PdfOptions {
                HeaderOffsetY = 30,
                FooterOffsetY = 28
            }
        });

        using PdfPigDocument lowOffsetPdf = PdfPigDocument.Open(new MemoryStream(lowOffsetBytes));
        using PdfPigDocument highOffsetPdf = PdfPigDocument.Open(new MemoryStream(highOffsetBytes));
        double lowHeaderY = FindHeaderFooterWordStartY(lowOffsetPdf.GetPage(1), "Offset", lowest: false);
        double highHeaderY = FindHeaderFooterWordStartY(highOffsetPdf.GetPage(1), "Offset", lowest: false);
        double lowFooterY = FindHeaderFooterWordStartY(lowOffsetPdf.GetPage(1), "Offset", lowest: true);
        double highFooterY = FindHeaderFooterWordStartY(highOffsetPdf.GetPage(1), "Offset", lowest: true);

        Assert.True(highHeaderY > lowHeaderY + 15D, $"Expected custom HeaderOffsetY to move the native Word PDF header. Low: {lowHeaderY:0.##}, high: {highHeaderY:0.##}.");
        Assert.True(lowFooterY > highFooterY + 15D, $"Expected custom FooterOffsetY to move the native Word PDF footer. Low: {lowFooterY:0.##}, high: {highFooterY:0.##}.");
    }

    private static double FindHeaderFooterWordStartY(UglyToad.PdfPig.Content.Page page, string word, bool lowest) {
        var words = page.GetWords()
            .Where(item => item.Text == word)
            .OrderBy(item => item.BoundingBox.Bottom)
            .ToList();

        Assert.NotEmpty(words);
        return lowest ? words[0].BoundingBox.Bottom : words[words.Count - 1].BoundingBox.Bottom;
    }
}

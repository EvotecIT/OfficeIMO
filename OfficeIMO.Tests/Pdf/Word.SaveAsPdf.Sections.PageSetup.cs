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

        using PdfPigDocument pdf = PdfPigDocument.Open(bytes);
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

        using PdfPigDocument pdf = PdfPigDocument.Open(bytes);
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

        using PdfPigDocument pdf = PdfPigDocument.Open(bytes);
        var firstPage = pdf.GetPage(1);
        var secondPage = pdf.GetPage(2);
        Assert.Contains("NarrowMarginMarker", firstPage.Text);
        Assert.Contains("WideMarginMarker", secondPage.Text);

        double narrowX = FindWordStartX(firstPage, "NarrowMarginMarker");
        double wideX = FindWordStartX(secondPage, "WideMarginMarker");
        Assert.InRange(narrowX, 35D, 48D);
        Assert.InRange(wideX, 140D, 156D);
    }
}

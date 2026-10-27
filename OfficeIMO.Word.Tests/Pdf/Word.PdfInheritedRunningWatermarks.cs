using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using Xunit;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Tests;

public partial class Word {
    [Theory]
    [InlineData(false, 0)]
    [InlineData(true, 0)]
    [InlineData(false, 5)]
    [InlineData(true, 5)]
    [InlineData(false, 6)]
    [InlineData(true, 6)]
    public void SaveAsPdf_InheritedDefaultRunningWatermarkSuppressesMissingEvenStory(bool image, int start) {
        using WordDocument document = CreateInheritedRunningStoryControl(false, "default", "next");
        W.Settings settings = document._wordprocessingDocument.MainDocumentPart!.DocumentSettingsPart!.Settings;
        settings.RemoveAllChildren<W.EvenAndOddHeaders>();
        settings.AddChild(new W.EvenAndOddHeaders(), true);
        WordHeader story = document.Sections[0].Header!.Default!;
        foreach (var child in story.ChildElements.ToArray()) child.Remove();
        story.AddWatermark(image ? WordWatermarkStyle.Image : WordWatermarkStyle.Text,
            image ? Path.Combine(_directoryWithImages, "Kulek.jpg") : "DEFAULTMARK");
        if (start > 0) document.Sections[1]._sectionProperties.AddChild(new W.PageNumberType { Start = start }, true);
        Assert.Null(document.Sections[0].Header?.Even);
        Assert.Null(document.Sections[1].Header?.Even);
        using var pdf = PdfPigDocument.Open(document.ToPdf(EmptyRunningOptions()).ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true }));
        Assert.Equal(start == 5 ? 4 : 3, pdf.NumberOfPages);
        for (int page = 1; page <= pdf.NumberOfPages; page++) {
            bool expected = page == 1 || page == 3;
            if (image) Assert.Equal(expected ? 1 : 0, pdf.GetPage(page).GetImages().Count());
            else Assert.Equal(expected, pdf.GetPage(page).Text.Contains("DEFAULTMARK"));
        }
    }

    [Theory]
    [InlineData(false, 0)]
    [InlineData(true, 0)]
    [InlineData(false, 5)]
    [InlineData(true, 5)]
    [InlineData(false, 6)]
    [InlineData(true, 6)]
    public void SaveAsPdf_InheritedEvenRunningWatermarkUsesDisplayedNumber(bool image, int start) {
        using WordDocument document = CreateInheritedRunningStoryControl(false, "even", "next");
        WordHeader story = document.Sections[0].Header!.Even!;
        foreach (var child in story.ChildElements.ToArray()) child.Remove();
        story.AddWatermark(image ? WordWatermarkStyle.Image : WordWatermarkStyle.Text,
            image ? Path.Combine(_directoryWithImages, "Kulek.jpg") : "EVENMARK");
        if (start > 0) document.Sections[1]._sectionProperties.AddChild(new W.PageNumberType { Start = start }, true);
        using var pdf = PdfPigDocument.Open(document.ToPdf(EmptyRunningOptions()).ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true }));
        Assert.Equal(start == 5 ? 4 : 3, pdf.NumberOfPages);
        for (int page = 1; page <= pdf.NumberOfPages; page++) {
            bool expected = start == 5 ? page == 4 : page == 2;
            if (image) Assert.Equal(expected ? 1 : 0, pdf.GetPage(page).GetImages().Count());
            else Assert.Equal(expected, pdf.GetPage(page).Text.Contains("EVENMARK"));
        }
    }
}

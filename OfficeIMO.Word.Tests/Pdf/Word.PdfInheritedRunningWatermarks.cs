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
    public void SaveAsPdf_InheritedEvenRunningWatermarkUsesDisplayedNumber(bool image, int start) {
        using WordDocument document = CreateInheritedRunningStoryControl(false, "even", "next");
        WordHeader story = document.Sections[0].Header!.Even!;
        foreach (var child in story.ChildElements.ToArray()) child.Remove();
        story.AddWatermark(image ? WordWatermarkStyle.Image : WordWatermarkStyle.Text,
            image ? Path.Combine(_directoryWithImages, "Kulek.jpg") : "EVENMARK");
        if (start > 0) document.Sections[1]._sectionProperties.AddChild(new W.PageNumberType { Start = start }, true);
        using var pdf = PdfPigDocument.Open(document.ToPdfBytes(EmptyRunningOptions()));
        Assert.Equal(start == 5 ? 4 : 3, pdf.NumberOfPages);
        for (int page = 1; page <= pdf.NumberOfPages; page++) {
            bool expected = start == 5 ? page == 4 : page == 2;
            if (image) Assert.Equal(expected ? 1 : 0, pdf.GetPage(page).GetImages().Count());
            else Assert.Equal(expected, pdf.GetPage(page).Text.Contains("EVENMARK"));
        }
    }
}

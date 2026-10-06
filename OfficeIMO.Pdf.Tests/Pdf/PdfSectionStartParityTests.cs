using OfficeIMO.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfSectionStartParityTests {
    [Theory]
    [InlineData(false, 3)]
    [InlineData(true, 2)]
    public void StartParityCanUsePhysicalIndexOrContinuingNumber(bool useContinuingNumber, int expectedPages) {
        var document = PdfDocument.Create(builder => builder.Section(page => page.PageNumberStart(2)
            .Content(content => content.Text("First"))));
        document.Section(page => page.StartOnPageParity(PdfPageParity.Odd, useContinuingNumber)
            .PageNumberStart(8).Content(content => content.Text("Second")));
        using PdfPigDocument pdf = PdfPigDocument.Open(document.ToBytes());
        Assert.Equal(expectedPages, pdf.NumberOfPages);
        Assert.Contains("Second", pdf.GetPage(expectedPages).Text);
    }

    [Theory]
    [InlineData(false, 2, "plain")]
    [InlineData(true, 3, "plain")]
    [InlineData(false, 2, "container")]
    [InlineData(true, 3, "container")]
    [InlineData(false, 2, "columns")]
    [InlineData(true, 3, "columns")]
    public void ExplicitBreaksCanPreserveAnEmptyPage(bool preserveEmpty, int expectedPages, string mode) {
        var document = PdfDocument.Create(builder => builder.Content(content => {
            void AddBreaks(PdfContentBuilder flow) {
                flow.Text("First");
                flow.PageBreak();
                flow.PageBreak(preserveEmpty);
                flow.Text("Last");
            }
            switch (mode) {
                case "container": content.Element(element => element.Content(AddBreaks)); break;
                case "columns": content.Columns(AddBreaks); break;
                default: AddBreaks(content); break;
            }
        }));
        using PdfPigDocument pdf = PdfPigDocument.Open(document.ToBytes());
        Assert.Equal(expectedPages, pdf.NumberOfPages);
        Assert.Contains("Last", pdf.GetPage(expectedPages).Text);
        if (preserveEmpty) Assert.True(string.IsNullOrWhiteSpace(pdf.GetPage(2).Text));
    }
}

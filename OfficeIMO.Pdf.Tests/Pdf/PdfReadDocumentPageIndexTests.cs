using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfReadDocumentPageIndexTests {
    [Fact]
    public void ManyNamedDestinationsResolveToTheirPages() {
        const int pageCount = 40;
        byte[] bytes = PdfDocument.Create(pdf => pdf.Content(content => {
            for (int page = 1; page <= pageCount; page++) {
                if (page > 1) content.PageBreak();
                content.Bookmark("Page " + page.ToString("D3", System.Globalization.CultureInfo.InvariantCulture));
                content.Paragraph(paragraph => paragraph.Text("Content " + page));
            }
        })).ToBytes();

        PdfReadDocument read = PdfReadDocument.Open(bytes);
        Assert.Equal(pageCount, read.NamedDestinations.Count);
        for (int index = 0; index < pageCount; index++) {
            Assert.Equal(index + 1, read.GetPageNumberForObject(read.Pages[index].ObjectNumber));
            Assert.Contains(read.NamedDestinations, destination =>
                destination.Name == "Page " + (index + 1).ToString("D3", System.Globalization.CultureInfo.InvariantCulture) &&
                destination.PageNumber == index + 1);
        }
        Assert.Null(read.GetPageNumberForObject(int.MaxValue));
    }
}

using System;
using System.IO;
using OfficeIMO;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfPageLengthTests {
    [Fact]
    public void MixedUnitsRetainExactPageAndMarginLengthsUntilPdfSerialization() {
        var size = new PdfPageSize(OfficeLength.Emus(2540001), OfficeLength.Inches(6));
        var margins = new PdfPageMargins(OfficeLength.Centimeters(1), OfficeLength.Inches(0.25),
            OfficeLength.Points(18), OfficeLength.Emus(127001));
        Assert.Equal(2540001, size.Portrait().Width.ToEmus());
        Assert.Equal(2540001, size.Landscape().Height.ToEmus());
        Assert.Equal(360000, margins.Left.ToEmus());
        Assert.Equal(127001, margins.Bottom.ToEmus());
        using var document = PdfDocument.Create();
        document.Compose(compose => compose.Page(page => {
            page.Size(size).Margin(margins);
            page.Content(content => content.Column(column => column.Item().Paragraph(text => text.Text("Explicit lengths"))));
        }));
        using var saved = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(document.ToBytes()));
        Assert.InRange(saved.GetPage(1).Width, size.Width.ToPoints() - 0.001, size.Width.ToPoints() + 0.001);
        Assert.Equal(432, saved.GetPage(1).Height);
        Assert.Contains("Explicit", saved.GetPage(1).Text);
        Assert.Throws<ArgumentOutOfRangeException>(() => new PdfPageSize(OfficeLength.Emus(0), size.Height));
        Assert.Throws<ArgumentOutOfRangeException>(() => PdfPageMargins.Uniform(OfficeLength.Emus(-1)));
    }
}

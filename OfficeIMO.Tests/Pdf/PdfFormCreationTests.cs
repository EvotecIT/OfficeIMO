using System.Text;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfFormCreationTests {
    [Fact]
    public void TextField_CreatesInspectableAcroFormField() {
        byte[] pdf = PdfDoc.Create()
            .Meta(title: "Generated form")
            .Paragraph(p => p.Text("Generated field:"))
            .TextField("Person.Name", width: 180, height: 24, value: "Ada Lovelace", spacingAfter: 12)
            .ToBytes();

        string raw = Encoding.ASCII.GetString(pdf);
        PdfDocumentInfo info = PdfInspector.Inspect(pdf);
        PdfDocumentPreflight preflight = PdfInspector.Preflight(pdf);

        Assert.Contains("/AcroForm", raw);
        Assert.Contains("/Subtype /Widget", raw);
        Assert.Contains("/FT /Tx", raw);
        Assert.Contains("/AP << /N", raw);
        Assert.True(info.HasReadableFormFields);
        PdfFormField field = Assert.Single(info.FormFields);
        Assert.Equal("Person.Name", field.Name);
        Assert.Equal(PdfFormFieldKind.Text, field.Kind);
        Assert.Equal("Ada Lovelace", field.Value);
        PdfFormWidget widget = Assert.Single(field.Widgets);
        Assert.Equal(1, widget.PageNumber);
        Assert.True(widget.Width > 170);
        Assert.True(widget.Height > 20);
        Assert.True(preflight.CanFillSimpleFormFields);
        Assert.True(preflight.CanFlattenSimpleFormFields);
        Assert.True(preflight.CanFillAndFlattenSimpleFormFields);
    }

    [Fact]
    public void TextField_CanBeFilledAndFlattened() {
        byte[] pdf = PdfDoc.Create()
            .TextField("Person.Name", width: 180, height: 24, value: "Original")
            .ToBytes();

        byte[] flattened = PdfFormFiller.FillAndFlattenFields(pdf, new Dictionary<string, string> {
            ["Person.Name"] = "Created fill"
        });

        string raw = Encoding.ASCII.GetString(flattened);
        PdfDocumentInfo info = PdfInspector.Inspect(flattened);

        Assert.False(info.HasReadableFormFields);
        Assert.DoesNotContain("/AcroForm", raw);
        Assert.Contains("(Created fill) Tj", raw);
    }

    [Fact]
    public void TextField_ValidatesFlowGeometry() {
        Assert.Throws<ArgumentException>(() => PdfDoc.Create().TextField(" "));
        Assert.Throws<ArgumentOutOfRangeException>(() => PdfDoc.Create().TextField("Name", width: 0));
        Assert.Throws<ArgumentOutOfRangeException>(() => PdfDoc.Create().TextField("Name", height: -1));
        Assert.Throws<ArgumentException>(() => PdfDoc.Create().TextField("Name", align: PdfAlign.Justify));
    }
}

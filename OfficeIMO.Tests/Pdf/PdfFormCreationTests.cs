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
    public void CheckBox_CreatesInspectableAcroFormField() {
        byte[] pdf = PdfDoc.Create()
            .Paragraph(p => p.Text("Generated checkbox:"))
            .CheckBox("AcceptTerms", isChecked: true, size: 16, spacingAfter: 12)
            .ToBytes();

        string raw = Encoding.ASCII.GetString(pdf);
        PdfDocumentInfo info = PdfInspector.Inspect(pdf);
        PdfDocumentPreflight preflight = PdfInspector.Preflight(pdf);

        Assert.Contains("/AcroForm", raw);
        Assert.Contains("/Subtype /Widget", raw);
        Assert.Contains("/FT /Btn", raw);
        Assert.Contains("/AS /Yes", raw);
        Assert.Contains("/AP << /N << /Off", raw);
        PdfFormField field = Assert.Single(info.FormFields);
        Assert.Equal("AcceptTerms", field.Name);
        Assert.Equal(PdfFormFieldKind.Button, field.Kind);
        Assert.True(field.IsCheckBox);
        Assert.Equal("Yes", field.Value);
        PdfFormWidget widget = Assert.Single(field.Widgets);
        Assert.Equal("Yes", widget.AppearanceState);
        Assert.True(widget.HasNormalAppearanceState("Off"));
        Assert.True(widget.HasNormalAppearanceState("Yes"));
        Assert.True(preflight.CanFillSimpleFormFields);
        Assert.True(preflight.CanFlattenSimpleFormFields);
        Assert.True(preflight.CanFillAndFlattenSimpleFormFields);
    }

    [Fact]
    public void CheckBox_CanBeFilledAndFlattened() {
        byte[] pdf = PdfDoc.Create()
            .CheckBox("AcceptTerms")
            .ToBytes();

        byte[] filled = PdfFormFiller.FillFields(pdf, new Dictionary<string, string> {
            ["AcceptTerms"] = "Yes"
        });
        PdfDocumentInfo filledInfo = PdfInspector.Inspect(filled);

        Assert.Equal("Yes", Assert.Single(filledInfo.FormFields).Value);
        Assert.Equal("Yes", Assert.Single(filledInfo.FormWidgets).AppearanceState);

        byte[] flattened = PdfFormFiller.FillAndFlattenFields(pdf, new Dictionary<string, string> {
            ["AcceptTerms"] = "Yes"
        });
        string raw = Encoding.ASCII.GetString(flattened);

        Assert.False(PdfInspector.Inspect(flattened).HasReadableFormFields);
        Assert.DoesNotContain("/AcroForm", raw);
        Assert.DoesNotContain("/Subtype /Widget", raw);
        Assert.Contains("/OfficeIMOForm1 Do", raw);
    }

    [Fact]
    public void GeneratedFields_ValidateFlowGeometry() {
        Assert.Throws<ArgumentException>(() => PdfDoc.Create().TextField(" "));
        Assert.Throws<ArgumentOutOfRangeException>(() => PdfDoc.Create().TextField("Name", width: 0));
        Assert.Throws<ArgumentOutOfRangeException>(() => PdfDoc.Create().TextField("Name", height: -1));
        Assert.Throws<ArgumentException>(() => PdfDoc.Create().TextField("Name", align: PdfAlign.Justify));
        Assert.Throws<ArgumentException>(() => PdfDoc.Create().CheckBox(" "));
        Assert.Throws<ArgumentOutOfRangeException>(() => PdfDoc.Create().CheckBox("AcceptTerms", size: 0));
        Assert.Throws<ArgumentException>(() => PdfDoc.Create().CheckBox("AcceptTerms", align: PdfAlign.Justify));
        Assert.Throws<ArgumentException>(() => PdfDoc.Create().CheckBox("AcceptTerms", checkedValueName: " "));
        Assert.Throws<ArgumentException>(() => PdfDoc.Create().CheckBox("AcceptTerms", checkedValueName: "Off"));
    }
}

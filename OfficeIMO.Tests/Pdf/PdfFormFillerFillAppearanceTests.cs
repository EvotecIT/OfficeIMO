using System.Text;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfFormFillerTests {
    [Fact]
    public void FillFields_UpdatesSimpleTextAndButtonValues() {
        byte[] filled = PdfFormFiller.FillFields(BuildHierarchicalFormPdf(), new Dictionary<string, string> {
            ["Person.Name"] = "Evotec",
            ["AcceptTerms"] = "Off"
        });

        PdfDocumentInfo info = PdfInspector.Inspect(filled);

        Assert.True(info.HasReadableFormFields);
        Assert.Equal(new[] { "Person.Name", "AcceptTerms" }, info.FormFieldNames);
        Assert.Equal("Evotec", info.FormFields[0].Value);
        Assert.Equal("Off", info.FormFields[1].Value);
        Assert.Contains("/NeedAppearances true", Encoding.ASCII.GetString(filled));
        Assert.False(PdfInspector.Preflight(filled).CanRewrite);
    }

    [Fact]
    public void FillFields_GeneratesSimpleTextWidgetAppearance() {
        byte[] filled = PdfFormFiller.FillFields(BuildTextWidgetFormPdf(), new Dictionary<string, string> {
            ["Name"] = "Visible value"
        });

        string output = Encoding.ASCII.GetString(filled);
        PdfDocumentInfo info = PdfInspector.Inspect(filled);

        Assert.Equal("Visible value", info.FormFields[0].Value);
        Assert.Contains("/Subtype /Form", output);
        Assert.Contains("/AP << /N", output);
        Assert.Contains("/Helv", output);
        Assert.Contains("<56697369626C652076616C7565> Tj", output);
    }

    [Fact]
    public void FillFields_PreservesUnicodeTextStringsWhenRewriting() {
        byte[] filled = PdfFormFiller.FillFields(BuildUnicodeFieldNameFormPdf(), new Dictionary<string, string> {
            ["名"] = "Visible value"
        });

        PdfFormField field = Assert.Single(PdfInspector.Inspect(filled).FormFields);
        string output = Encoding.ASCII.GetString(filled);

        Assert.Equal("名", field.Name);
        Assert.Equal("Visible value", field.Value);
        Assert.Contains("/T <FEFF540D>", output, StringComparison.Ordinal);
    }

    [Fact]
    public void FillFields_GeneratesSimpleButtonWidgetAppearances() {
        byte[] filled = PdfFormFiller.FillFields(BuildCheckboxWidgetWithoutAppearancePdf(), new Dictionary<string, string> {
            ["AcceptTerms"] = "Yes"
        });

        string output = Encoding.ASCII.GetString(filled);
        PdfDocumentInfo info = PdfInspector.Inspect(filled);

        Assert.Equal("Yes", info.FormFields[0].Value);
        Assert.Contains("/AS /Yes", output);
        Assert.Contains("/AP << /N <<", output);
        Assert.Contains("/Off", output);
        Assert.Contains("/Yes", output);
        Assert.Contains("1.25 w", output);
    }
}

using System.Text;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfFormFillerTests {
    [Fact]
    public void FlattenFields_PaintsTextWidgetAndRemovesFormAnnotations() {
        byte[] filled = PdfFormFiller.FillFields(BuildTextWidgetFormPdf(), new Dictionary<string, string> {
            ["Name"] = "Flattened value"
        });

        byte[] flattened = PdfFormFiller.FlattenFields(filled);

        string output = Encoding.ASCII.GetString(flattened);
        PdfDocumentInfo info = PdfInspector.Inspect(flattened);

        Assert.False(info.HasForms);
        Assert.False(info.HasReadableFormFields);
        Assert.DoesNotContain("/AcroForm", output);
        Assert.DoesNotContain("/Subtype /Widget", output);
        Assert.DoesNotContain("/Annots", output);
        Assert.Contains("/OfficeIMOForm1", output);
        Assert.Contains("/OfficeIMOForm1 Do", output);
        Assert.Contains("<466C617474656E65642076616C7565> Tj", output);
    }

    [Fact]
    public void FlattenFields_FlattensReferencedContentArraysBeforeAppendingAppearanceStream() {
        byte[] filled = PdfFormFiller.FillFields(BuildTextWidgetFormPdfWithReferencedContentArray(), new Dictionary<string, string> {
            ["Name"] = "Flattened value"
        });

        byte[] flattened = PdfFormFiller.FlattenFields(filled);

        var (objects, _) = PdfSyntax.ParseObjects(flattened);
        var page = Assert.IsType<PdfDictionary>(objects.Values.First(indirect =>
            indirect.Value is PdfDictionary dictionary &&
            dictionary.Get<PdfName>("Type")?.Name == "Page").Value);
        var contents = Assert.IsType<PdfArray>(page.Items["Contents"]);

        Assert.True(contents.Items.Count >= 2);
        foreach (var item in contents.Items) {
            var reference = Assert.IsType<PdfReference>(item);
            Assert.IsType<PdfStream>(objects[reference.ObjectNumber].Value);
        }
    }

    [Fact]
    public void FillAndFlattenFields_UpdatesValueAndFlattensInOneCall() {
        byte[] flattened = PdfFormFiller.FillAndFlattenFields(BuildTextWidgetFormPdf(), new Dictionary<string, string> {
            ["Name"] = "Single pass"
        });

        string output = Encoding.ASCII.GetString(flattened);
        PdfDocumentInfo info = PdfInspector.Inspect(flattened);

        Assert.False(info.HasForms);
        Assert.DoesNotContain("/AcroForm", output);
        Assert.Contains("<53696E676C652070617373> Tj", output);
        Assert.Contains("/OfficeIMOForm1 Do", output);
    }

    [Fact]
    public void FlattenFields_PaintsSelectedButtonWidgetAppearance() {
        byte[] flattened = PdfFormFiller.FlattenFields(BuildCheckboxWidgetFormPdf());

        string output = Encoding.ASCII.GetString(flattened);
        PdfDocumentInfo info = PdfInspector.Inspect(flattened);

        Assert.False(info.HasForms);
        Assert.DoesNotContain("/AcroForm", output);
        Assert.DoesNotContain("/Subtype /Widget", output);
        Assert.DoesNotContain("/Annots", output);
        Assert.Contains("/OfficeIMOForm1 Do", output);
        Assert.Contains("Checked appearance", GetFlattenedAppearanceStreamText(flattened));
    }

    [Fact]
    public void FillAndFlattenFields_UpdatesButtonStateBeforeFlattening() {
        byte[] flattened = PdfFormFiller.FillAndFlattenFields(BuildCheckboxWidgetFormPdf(), new Dictionary<string, string> {
            ["AcceptTerms"] = "Off"
        });

        string output = Encoding.ASCII.GetString(flattened);
        PdfDocumentInfo info = PdfInspector.Inspect(flattened);

        Assert.False(info.HasForms);
        Assert.DoesNotContain("/AcroForm", output);
        Assert.Contains("/OfficeIMOForm1 Do", output);
        Assert.Contains("Unchecked appearance", GetFlattenedAppearanceStreamText(flattened));
    }

    [Fact]
    public void FillAndFlattenFields_GeneratesMissingButtonAppearanceBeforeFlattening() {
        byte[] flattened = PdfFormFiller.FillAndFlattenFields(BuildCheckboxWidgetWithoutAppearancePdf(), new Dictionary<string, string> {
            ["AcceptTerms"] = "Yes"
        });

        string output = Encoding.ASCII.GetString(flattened);
        string appearance = GetFlattenedAppearanceStreamText(flattened);

        Assert.False(PdfInspector.Inspect(flattened).HasForms);
        Assert.DoesNotContain("/AcroForm", output);
        Assert.DoesNotContain("/Subtype /Widget", output);
        Assert.Contains("/OfficeIMOForm1 Do", output);
        Assert.Contains("1.25 w", appearance);
        Assert.Contains(" l S", appearance);
    }

    [Fact]
    public void FillAndFlattenFields_FlattensOnlySelectedRadioWidget() {
        byte[] flattened = PdfFormFiller.FillAndFlattenFields(BuildRadioWidgetGroupWithoutOffAppearancePdf(), new Dictionary<string, string> {
            ["Payment.Method"] = "Wire"
        });

        string output = Encoding.ASCII.GetString(flattened);
        string appearances = string.Concat(GetFlattenedAppearanceStreamTexts(flattened));

        Assert.False(PdfInspector.Inspect(flattened).HasForms);
        Assert.DoesNotContain("/AcroForm", output);
        Assert.DoesNotContain("/Subtype /Widget", output);
        Assert.Contains("Wire selected", appearances);
        Assert.DoesNotContain("Card selected", appearances);
        Assert.DoesNotContain("Cash selected", appearances);
        Assert.DoesNotContain("1.25 w", appearances);
    }

    [Fact]
    public void FlattenFields_GeneratesMissingButtonAppearanceFromExistingState() {
        byte[] flattened = PdfFormFiller.FlattenFields(BuildCheckboxWidgetWithoutAppearancePdf("Yes"));

        string output = Encoding.ASCII.GetString(flattened);
        string appearance = GetFlattenedAppearanceStreamText(flattened);

        Assert.False(PdfInspector.Inspect(flattened).HasForms);
        Assert.DoesNotContain("/AcroForm", output);
        Assert.DoesNotContain("/Subtype /Widget", output);
        Assert.Contains("/OfficeIMOForm1 Do", output);
        Assert.Contains("1.25 w", appearance);
        Assert.Contains(" l S", appearance);
    }
}

using System.Text;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfFormFillerTests {
    [Fact]
    public void FillAndFlattenFields_PaintsChoiceOptionDisplayText() {
        byte[] filled = PdfFormFiller.FillFields(BuildChoiceWidgetFormPdf(), new Dictionary<string, string> {
            ["Country"] = "PL"
        });

        PdfDocumentInfo filledInfo = PdfInspector.Inspect(filled);

        Assert.True(filledInfo.HasReadableFormFields);
        PdfFormField filledField = Assert.Single(filledInfo.FormFields);
        Assert.Equal("PL", filledField.Value);
        Assert.Equal("Poland", Assert.Single(filledField.SelectedOptions).DisplayText);
        Assert.Contains("<506F6C616E64> Tj", Encoding.ASCII.GetString(filled), StringComparison.Ordinal);

        byte[] flattened = PdfFormFiller.FlattenFields(filled);

        string output = Encoding.ASCII.GetString(flattened);
        PdfDocumentInfo flattenedInfo = PdfInspector.Inspect(flattened);

        Assert.False(flattenedInfo.HasForms);
        Assert.DoesNotContain("/AcroForm", output);
        Assert.DoesNotContain("/Subtype /Widget", output);
        Assert.DoesNotContain("/Annots", output);
        Assert.Contains("/OfficeIMOForm1 Do", output);
        Assert.Contains("<506F6C616E64> Tj", GetFlattenedAppearanceStreamText(flattened), StringComparison.Ordinal);
    }

    [Fact]
    public void FillFields_ChoiceDisplayTextStoresExportValueAndPaintsDisplayText() {
        byte[] filled = PdfFormFiller.FillFields(BuildChoiceWidgetFormPdf(), new Dictionary<string, string> {
            ["Country"] = "United States"
        });

        PdfDocumentInfo filledInfo = PdfInspector.Inspect(filled);

        Assert.True(filledInfo.HasReadableFormFields);
        PdfFormField filledField = Assert.Single(filledInfo.FormFields);
        Assert.Equal("US", filledField.Value);
        Assert.Equal("United States", Assert.Single(filledField.SelectedOptions).DisplayText);
        Assert.Contains("<556E6974656420537461746573> Tj", Encoding.ASCII.GetString(filled), StringComparison.Ordinal);
    }

    [Fact]
    public void FillFields_ChoicePrefersExactExportValueOverEarlierDisplayText() {
        byte[] filled = PdfFormFiller.FillFields(BuildOverlappingChoiceWidgetFormPdf(), new Dictionary<string, string> {
            ["Choice"] = "B"
        });

        PdfFormField filledField = Assert.Single(PdfInspector.Inspect(filled).FormFields);

        Assert.Equal("B", filledField.Value);
        Assert.Equal("C", Assert.Single(filledField.SelectedOptions).DisplayText);
    }

    [Fact]
    public void FillFields_ChoiceDisplayTextUsesFirstMatchingOption() {
        byte[] filled = PdfFormFiller.FillFields(BuildDuplicateChoiceWidgetFormPdf(), new Dictionary<string, string> {
            ["Choice"] = "Same"
        });

        PdfFormField filledField = Assert.Single(PdfInspector.Inspect(filled).FormFields);

        Assert.Equal("A", filledField.Value);
        Assert.Equal("Same", Assert.Single(filledField.SelectedOptions).DisplayText);
    }

    [Fact]
    public void FillFields_ChoiceExportValueUsesFirstMatchingOption() {
        byte[] filled = PdfFormFiller.FillFields(BuildDuplicateChoiceWidgetFormPdf(), new Dictionary<string, string> {
            ["Choice"] = "C"
        });

        PdfFormField filledField = Assert.Single(PdfInspector.Inspect(filled).FormFields);

        Assert.Equal("C", filledField.Value);
        Assert.Equal(new[] { "First C", "Second C" }, filledField.SelectedOptions.Select(option => option.DisplayText).ToArray());
        Assert.Contains("<46697273742043> Tj", Encoding.ASCII.GetString(filled), StringComparison.Ordinal);
    }

    [Fact]
    public void FillFields_ChoiceDisplayTextUsesInheritedOptions() {
        byte[] filled = PdfFormFiller.FillFields(BuildInheritedChoiceWidgetFormPdf(), new Dictionary<string, string> {
            ["Selection.Country"] = "United States"
        });

        PdfDocumentInfo filledInfo = PdfInspector.Inspect(filled);

        PdfFormField filledField = Assert.Single(filledInfo.FormFields);
        Assert.Equal("Selection.Country", filledField.Name);
        Assert.Equal("US", filledField.Value);
        Assert.Equal("United States", Assert.Single(filledField.SelectedOptions).DisplayText);
        Assert.Contains("<556E6974656420537461746573> Tj", Encoding.ASCII.GetString(filled), StringComparison.Ordinal);
    }

    [Fact]
    public void FillFields_ChoiceRejectsUnknownValueWhenNotEditable() {
        var ex = Assert.Throws<ArgumentException>(() => PdfFormFiller.FillFields(BuildChoiceWidgetFormPdf(), new Dictionary<string, string> {
            ["Country"] = "Untied States"
        }));

        Assert.Contains("PDF choice field value does not match an available option: Untied States", ex.Message);
    }

    [Fact]
    public void FillFields_EditableChoiceAllowsUnknownValue() {
        byte[] filled = PdfFormFiller.FillFields(BuildEditableChoiceWidgetFormPdf(), new Dictionary<string, string> {
            ["Country"] = "Atlantis"
        });

        PdfFormField filledField = Assert.Single(PdfInspector.Inspect(filled).FormFields);

        Assert.Equal("Atlantis", filledField.Value);
        Assert.Contains("<41746C616E746973> Tj", Encoding.ASCII.GetString(filled), StringComparison.Ordinal);
    }

    [Fact]
    public void FillFields_RejectsMultipleScalarChoiceValues() {
        var ex = Assert.Throws<ArgumentException>(() => PdfFormFiller.FillFields(BuildChoiceWidgetFormPdf(), new Dictionary<string, PdfFormFieldValue> {
            ["Country"] = PdfFormFieldValue.FromValues("Poland", "Germany")
        }));

        Assert.Contains("PDF scalar choice field cannot be filled with multiple values.", ex.Message);
    }

    [Fact]
    public void FillFields_MultiSelectChoiceValuesStoreExportArrayAndPaintDisplayText() {
        byte[] filled = PdfFormFiller.FillFields(BuildMultiSelectChoiceWidgetFormPdf(), new Dictionary<string, PdfFormFieldValue> {
            ["Country"] = PdfFormFieldValue.FromValues("Poland", "United States")
        });

        PdfDocumentInfo filledInfo = PdfInspector.Inspect(filled);

        Assert.True(filledInfo.HasReadableFormFields);
        PdfFormField filledField = Assert.Single(filledInfo.FormFields);
        Assert.Equal(new[] { "PL", "US" }, filledField.Values);
        Assert.Equal(new[] { "Poland", "United States" }, filledField.SelectedOptions.Select(option => option.DisplayText).ToArray());
        string output = Encoding.ASCII.GetString(filled);
        Assert.Contains("<506F6C616E64> Tj", output, StringComparison.Ordinal);
        Assert.Contains("<556E6974656420537461746573> Tj", output, StringComparison.Ordinal);
        Assert.DoesNotContain("<506F6C616E642C20556E6974656420537461746573> Tj", output, StringComparison.Ordinal);
    }

    [Fact]
    public void FillAndFlattenFields_MultiSelectChoiceValuesPaintDisplayText() {
        byte[] flattened = PdfFormFiller.FillAndFlattenFields(BuildMultiSelectChoiceWidgetFormPdf(), new Dictionary<string, PdfFormFieldValue> {
            ["Country"] = PdfFormFieldValue.FromValues("Germany", "United States")
        });

        string output = Encoding.ASCII.GetString(flattened);

        Assert.False(PdfInspector.Inspect(flattened).HasForms);
        Assert.DoesNotContain("/AcroForm", output);
        Assert.DoesNotContain("/Subtype /Widget", output);
        Assert.Contains("/OfficeIMOForm1 Do", output);
        string appearanceText = GetFlattenedAppearanceStreamText(flattened);
        Assert.Contains("<4765726D616E79> Tj", appearanceText, StringComparison.Ordinal);
        Assert.Contains("<556E6974656420537461746573> Tj", appearanceText, StringComparison.Ordinal);
        Assert.DoesNotContain("<4765726D616E792C20556E6974656420537461746573> Tj", appearanceText, StringComparison.Ordinal);
    }

    [Fact]
    public void FlattenFields_PaintsMultiSelectChoiceOptionDisplayText() {
        byte[] source = BuildMultiSelectChoiceWidgetFormPdf();
        PdfDocumentInfo sourceInfo = PdfInspector.Inspect(source);

        PdfFormField sourceField = Assert.Single(sourceInfo.FormFields);
        Assert.Equal(new[] { "PL", "US" }, sourceField.Values);
        Assert.Equal(new[] { "Poland", "United States" }, sourceField.SelectedOptions.Select(option => option.DisplayText).ToArray());

        byte[] flattened = PdfFormFiller.FlattenFields(source);

        string output = Encoding.ASCII.GetString(flattened);
        PdfDocumentInfo flattenedInfo = PdfInspector.Inspect(flattened);

        Assert.False(flattenedInfo.HasForms);
        Assert.DoesNotContain("/AcroForm", output);
        Assert.DoesNotContain("/Subtype /Widget", output);
        Assert.DoesNotContain("/Annots", output);
        Assert.Contains("/OfficeIMOForm1 Do", output);
        string appearanceText = GetFlattenedAppearanceStreamText(flattened);
        Assert.Contains("<506F6C616E64> Tj", appearanceText, StringComparison.Ordinal);
        Assert.Contains("<556E6974656420537461746573> Tj", appearanceText, StringComparison.Ordinal);
        Assert.DoesNotContain("<506F6C616E642C20556E6974656420537461746573> Tj", appearanceText, StringComparison.Ordinal);
    }

    [Fact]
    public void FlattenFields_UsesInheritedChoiceOptionsForDisplayText() {
        byte[] flattened = PdfFormFiller.FlattenFields(BuildInheritedChoiceValueWidgetFormPdf());

        string output = Encoding.ASCII.GetString(flattened);

        Assert.False(PdfInspector.Inspect(flattened).HasForms);
        Assert.DoesNotContain("/AcroForm", output);
        Assert.Contains("<556E6974656420537461746573> Tj", GetFlattenedAppearanceStreamText(flattened), StringComparison.Ordinal);
    }

    [Fact]
    public void Preflight_AllowsChoiceFieldFlattening() {
        PdfDocumentPreflight preflight = PdfInspector.Preflight(BuildChoiceWidgetFormPdf());

        Assert.True(preflight.CanFlattenSimpleFormFields);
        Assert.True(preflight.CanFillAndFlattenSimpleFormFields);
        Assert.Empty(preflight.GetCapabilityDiagnostics(PdfPreflightCapability.FlattenSimpleFormFields));
    }
}

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
        Assert.Contains("<437265617465642066696C6C> Tj", raw);
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
    public void ChoiceField_CreatesInspectableAcroFormField() {
        byte[] pdf = PdfDoc.Create()
            .Paragraph(p => p.Text("Generated choice:"))
            .ChoiceField("Country", new[] { "Poland", "Germany", "United States" }, value: "Germany", width: 180, height: 24, spacingAfter: 12)
            .ToBytes();

        string raw = Encoding.ASCII.GetString(pdf);
        PdfDocumentInfo info = PdfInspector.Inspect(pdf);
        PdfDocumentPreflight preflight = PdfInspector.Preflight(pdf);

        Assert.Contains("/AcroForm", raw);
        Assert.Contains("/Subtype /Widget", raw);
        Assert.Contains("/FT /Ch", raw);
        Assert.Contains("/Opt [ <506F6C616E64> <4765726D616E79> <556E6974656420537461746573> ]", raw);
        Assert.Contains("/Ff 131072", raw);
        PdfFormField field = Assert.Single(info.FormFields);
        Assert.Equal("Country", field.Name);
        Assert.Equal(PdfFormFieldKind.Choice, field.Kind);
        Assert.True(field.IsCombo);
        Assert.Equal("Germany", field.Value);
        Assert.Equal(new[] { "Poland", "Germany", "United States" }, field.Options.Select(option => option.DisplayText).ToArray());
        Assert.Equal("Germany", Assert.Single(field.SelectedOptions).DisplayText);
        PdfFormWidget widget = Assert.Single(field.Widgets);
        Assert.Equal(1, widget.PageNumber);
        Assert.True(widget.Width > 170);
        Assert.True(widget.Height > 20);
        Assert.True(preflight.CanFillSimpleFormFields);
        Assert.True(preflight.CanFlattenSimpleFormFields);
        Assert.True(preflight.CanFillAndFlattenSimpleFormFields);
    }

    [Fact]
    public void ChoiceField_CanBeFilledAndFlattened() {
        byte[] pdf = PdfDoc.Create()
            .ChoiceField("Country", new[] { "Poland", "Germany", "United States" }, value: "Poland")
            .ToBytes();

        byte[] filled = PdfFormFiller.FillFields(pdf, new Dictionary<string, string> {
            ["Country"] = "United States"
        });
        PdfDocumentInfo filledInfo = PdfInspector.Inspect(filled);
        PdfFormField filledField = Assert.Single(filledInfo.FormFields);

        Assert.Equal("United States", filledField.Value);
        Assert.Equal("United States", Assert.Single(filledField.SelectedOptions).DisplayText);
        Assert.Contains("<556E6974656420537461746573> Tj", Encoding.ASCII.GetString(filled), StringComparison.Ordinal);

        byte[] flattened = PdfFormFiller.FillAndFlattenFields(pdf, new Dictionary<string, string> {
            ["Country"] = "United States"
        });
        string raw = Encoding.ASCII.GetString(flattened);

        Assert.False(PdfInspector.Inspect(flattened).HasReadableFormFields);
        Assert.DoesNotContain("/AcroForm", raw);
        Assert.DoesNotContain("/Subtype /Widget", raw);
        Assert.Contains("/OfficeIMOForm1 Do", raw);
    }

    [Fact]
    public void MultiSelectChoiceField_CreatesInspectableAcroFormField() {
        byte[] pdf = PdfDoc.Create()
            .Paragraph(p => p.Text("Generated multi-select choice:"))
            .MultiSelectChoiceField("Countries", new[] { "Poland", "Germany", "United States" }, values: new[] { "Poland", "United States" }, width: 190, height: 72)
            .ToBytes();

        string raw = Encoding.ASCII.GetString(pdf);
        PdfDocumentInfo info = PdfInspector.Inspect(pdf);
        PdfDocumentPreflight preflight = PdfInspector.Preflight(pdf);

        Assert.Contains("/FT /Ch", raw);
        Assert.Contains("/V [<506F6C616E64> <556E6974656420537461746573>]", raw);
        Assert.Contains("/DV [<506F6C616E64> <556E6974656420537461746573>]", raw);
        Assert.Contains("/Ff 2097152", raw);
        PdfFormField field = Assert.Single(info.FormFields);
        Assert.Equal("Countries", field.Name);
        Assert.Equal(PdfFormFieldKind.Choice, field.Kind);
        Assert.False(field.IsCombo);
        Assert.True(field.AllowsMultipleSelection);
        Assert.Equal(new[] { "Poland", "United States" }, field.Values);
        Assert.Equal(new[] { "Poland", "United States" }, field.SelectedOptions.Select(option => option.DisplayText).ToArray());
        PdfFormWidget widget = Assert.Single(field.Widgets);
        Assert.True(widget.Width > 180);
        Assert.True(widget.Height > 70);
        Assert.True(preflight.CanFillSimpleFormFields);
        Assert.True(preflight.CanFlattenSimpleFormFields);
        Assert.True(preflight.CanFillAndFlattenSimpleFormFields);
    }

    [Fact]
    public void MultiSelectChoiceField_CanBeFilledAndFlattened() {
        byte[] pdf = PdfDoc.Create()
            .MultiSelectChoiceField("Countries", new[] { "Poland", "Germany", "United States" }, values: new[] { "Poland" })
            .ToBytes();

        byte[] filled = PdfFormFiller.FillFields(pdf, new Dictionary<string, PdfFormFieldValue> {
            ["Countries"] = PdfFormFieldValue.FromValues("Germany", "United States")
        });
        PdfFormField filledField = Assert.Single(PdfInspector.Inspect(filled).FormFields);

        Assert.Equal(new[] { "Germany", "United States" }, filledField.Values);
        Assert.Equal(new[] { "Germany", "United States" }, filledField.SelectedOptions.Select(option => option.DisplayText).ToArray());
        Assert.Contains("<4765726D616E792C20556E6974656420537461746573> Tj", Encoding.ASCII.GetString(filled), StringComparison.Ordinal);

        byte[] scalarFilled = PdfFormFiller.FillFields(pdf, new Dictionary<string, string> {
            ["Countries"] = "Germany"
        });
        string scalarRaw = Encoding.ASCII.GetString(scalarFilled);
        PdfFormField scalarField = Assert.Single(PdfInspector.Inspect(scalarFilled).FormFields);

        Assert.Equal(new[] { "Germany" }, scalarField.Values);
        Assert.Contains("/V [", scalarRaw, StringComparison.Ordinal);
        Assert.Contains("<4765726D616E79>", scalarRaw, StringComparison.Ordinal);

        byte[] flattened = PdfFormFiller.FillAndFlattenFields(pdf, new Dictionary<string, PdfFormFieldValue> {
            ["Countries"] = PdfFormFieldValue.FromValues("Germany", "United States")
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
        Assert.Throws<ArgumentException>(() => PdfDoc.Create().ChoiceField(" ", new[] { "One" }));
        Assert.Throws<ArgumentNullException>(() => PdfDoc.Create().ChoiceField("Country", null!));
        Assert.Throws<ArgumentException>(() => PdfDoc.Create().ChoiceField("Country", Array.Empty<string>()));
        Assert.Throws<ArgumentException>(() => PdfDoc.Create().ChoiceField("Country", new[] { "One", "One" }));
        Assert.Throws<ArgumentException>(() => PdfDoc.Create().ChoiceField("Country", new[] { "One", " " }));
        Assert.Throws<ArgumentException>(() => PdfDoc.Create().ChoiceField("Country", new[] { "One" }, value: "Two"));
        Assert.Throws<ArgumentOutOfRangeException>(() => PdfDoc.Create().ChoiceField("Country", new[] { "One" }, width: 0));
        Assert.Throws<ArgumentOutOfRangeException>(() => PdfDoc.Create().ChoiceField("Country", new[] { "One" }, height: -1));
        Assert.Throws<ArgumentException>(() => PdfDoc.Create().ChoiceField("Country", new[] { "One" }, align: PdfAlign.Justify));
        Assert.Throws<ArgumentException>(() => PdfDoc.Create().MultiSelectChoiceField("Countries", Array.Empty<string>()));
        Assert.Throws<ArgumentException>(() => PdfDoc.Create().MultiSelectChoiceField("Countries", new[] { "One" }, values: Array.Empty<string>()));
        Assert.Throws<ArgumentException>(() => PdfDoc.Create().MultiSelectChoiceField("Countries", new[] { "One" }, values: new[] { "Two" }));
        Assert.Throws<ArgumentException>(() => PdfDoc.Create().MultiSelectChoiceField("Countries", new[] { "One" }, values: new[] { "One", "One" }));
        Assert.Throws<ArgumentOutOfRangeException>(() => PdfDoc.Create().MultiSelectChoiceField("Countries", new[] { "One" }, height: 0));
    }
}

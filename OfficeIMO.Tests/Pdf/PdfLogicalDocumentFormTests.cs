using System.Text;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfLogicalDocumentTests {
    [Fact]
    public void Load_ExposesSimpleAcroFormFields() {
        PdfLogicalDocument logical = PdfLogicalDocument.Load(BuildHierarchicalFormPdf());

        Assert.True(logical.HasFormFields);
        Assert.True(logical.HasAcroFormNeedAppearances);
        Assert.Equal(true, logical.AcroFormNeedAppearances);
        Assert.True(logical.RequiresAcroFormAppearanceRegeneration);
        Assert.True(logical.HasAcroFormSignatureFlags);
        Assert.Equal(1, logical.AcroFormSignatureFlags);
        Assert.True(logical.AcroFormSignaturesExist);
        Assert.False(logical.AcroFormAppendOnly);
        Assert.True(logical.HasAcroFormDefaultAppearance);
        Assert.Equal("/Helv 7 Tf 0.5 g", logical.AcroFormDefaultAppearance);
        Assert.Equal(new[] { "Person.Name", "AcceptTerms", "Selection.Country" }, logical.FormFields.Select(field => field.Name).ToArray());
        Assert.Equal("OfficeIMO", logical.FormFields[0].Value);
        Assert.Equal("InheritedDraft", logical.FormFields[0].DefaultValue);
        Assert.Equal(64, logical.FormFields[0].MaxLength);
        Assert.True(logical.FormFields[0].IsReadOnly);
        Assert.Equal("/Helv 10 Tf 0 g", logical.FormFields[0].DefaultAppearance);
        Assert.Equal(2, logical.FormFields[0].Quadding);
        Assert.Equal(PdfFormFieldTextAlignment.Right, logical.FormFields[0].TextAlignment);
        Assert.Equal("Yes", logical.FormFields[1].Value);
        Assert.Equal("DE", logical.FormFields[2].Value);
        Assert.Equal("PL", logical.FormFields[2].DefaultValue);
        Assert.Equal("/Helv 8 Tf 0 0 1 rg", logical.FormFields[2].DefaultAppearance);
        Assert.Equal(PdfFormFieldTextAlignment.Center, logical.FormFields[2].TextAlignment);
        Assert.Equal(2, logical.FormFields[2].OptionCount);
        Assert.Equal(new[] { "DE" }, logical.FormFields[2].SelectedOptions.Select(option => option.ExportValue).ToArray());
        Assert.Equal(new[] { "PL" }, logical.FormFields[2].DefaultSelectedOptions.Select(option => option.ExportValue).ToArray());
        Assert.Equal(3, logical.FormFieldsByName.Count);
        Assert.Contains("Person.Name", logical.FormFieldNames);
        Assert.Contains("AcceptTerms", logical.FormFieldNames);
        Assert.Contains("Selection.Country", logical.FormFieldNames);

        Assert.True(logical.TryGetFormField("Person.Name", out PdfFormField? nameField));
        Assert.Equal("OfficeIMO", nameField!.Value);
        Assert.Equal(new[] { "InheritedDraft" }, nameField.DefaultValues);
        Assert.True(nameField.HasDefaultAppearance);
        Assert.True(logical.TryGetFormField("AcceptTerms", out PdfFormField? acceptField));
        Assert.Equal("Yes", acceptField!.Value);
        Assert.True(logical.TryGetFormField("Selection.Country", out PdfFormField? countryField));
        Assert.True(countryField!.IsChoiceField);
        Assert.False(logical.TryGetFormField("Missing", out PdfFormField? missingField));
        Assert.Null(missingField);
    }

    [Fact]
    public void Load_ExposesAcroFormFieldKindsAndFlags() {
        PdfLogicalDocument logical = PdfLogicalDocument.Load(BuildFieldKindFormPdf());

        PdfFormField text = logical.FormFieldsByName["Notes"];
        Assert.Equal(PdfFormFieldKind.Text, text.Kind);
        Assert.True(text.IsTextField);
        Assert.True(text.IsReadOnly);
        Assert.True(text.IsRequired);
        Assert.True(text.IsNoExport);
        Assert.True(text.IsMultiline);
        Assert.True(text.IsPassword);
        Assert.Equal(42, text.MaxLength);
        Assert.Equal(new[] { "Secret" }, text.Values);
        Assert.Equal("Draft", text.DefaultValue);
        Assert.Equal(new[] { "Draft" }, text.DefaultValues);
        Assert.True(text.HasDefaultValues);
        Assert.Equal("/Helv 9 Tf 0 g", text.DefaultAppearance);
        Assert.Equal(PdfFormFieldTextAlignment.Left, text.TextAlignment);
        Assert.False(text.HasOptions);
        Assert.False(text.HasDefaultSelectedOptions);
        Assert.False(text.IsButtonField);
        Assert.False(text.IsChoiceField);

        PdfFormField checkBox = logical.FormFieldsByName["Accept"];
        Assert.Equal(PdfFormFieldKind.Button, checkBox.Kind);
        Assert.True(checkBox.IsButtonField);
        Assert.True(checkBox.IsCheckBox);
        Assert.False(checkBox.IsRadioButton);
        Assert.False(checkBox.IsPushButton);

        PdfFormField radio = logical.FormFieldsByName["Choice"];
        Assert.True(radio.IsRadioButton);
        Assert.True(radio.IsNoToggleToOff);
        Assert.False(radio.IsCheckBox);

        PdfFormField pushButton = logical.FormFieldsByName["Submit"];
        Assert.True(pushButton.IsPushButton);
        Assert.False(pushButton.IsCheckBox);

        PdfFormField choice = logical.FormFieldsByName["Country"];
        Assert.Equal(PdfFormFieldKind.Choice, choice.Kind);
        Assert.Equal("[PL US]", choice.Value);
        Assert.Equal(new[] { "PL", "US" }, choice.Values);
        Assert.Equal("[DE US]", choice.DefaultValue);
        Assert.Equal(new[] { "DE", "US" }, choice.DefaultValues);
        Assert.Equal("/Helv 8 Tf 0 0 1 rg", choice.DefaultAppearance);
        Assert.Equal(1, choice.Quadding);
        Assert.Equal(PdfFormFieldTextAlignment.Center, choice.TextAlignment);
        Assert.True(choice.IsChoiceField);
        Assert.True(choice.IsCombo);
        Assert.True(choice.IsEditableChoice);
        Assert.True(choice.IsSortedChoice);
        Assert.True(choice.AllowsMultipleSelection);
        Assert.True(choice.DoesNotSpellCheck);
        Assert.True(choice.CommitsOnSelectionChange);
        Assert.True(choice.HasOptions);
        Assert.Equal(3, choice.OptionCount);
        Assert.Equal("PL", choice.Options[0].ExportValue);
        Assert.Equal("Poland", choice.Options[0].DisplayText);
        Assert.True(choice.Options[0].HasSeparateDisplayText);
        Assert.Equal("DE", choice.Options[1].ExportValue);
        Assert.Equal("DE", choice.Options[1].DisplayText);
        Assert.False(choice.Options[1].HasSeparateDisplayText);
        Assert.Equal("US", choice.Options[2].ExportValue);
        Assert.Equal("United States", choice.Options[2].DisplayText);
        Assert.True(choice.HasSelectedOptions);
        Assert.Equal(2, choice.SelectedOptionCount);
        Assert.Equal(new[] { "PL", "US" }, choice.SelectedOptions.Select(option => option.ExportValue).ToArray());
        Assert.True(choice.HasDefaultSelectedOptions);
        Assert.Equal(2, choice.DefaultSelectedOptionCount);
        Assert.Equal(new[] { "DE", "US" }, choice.DefaultSelectedOptions.Select(option => option.ExportValue).ToArray());

        PdfFormField signature = logical.FormFieldsByName["Approval"];
        Assert.Equal(PdfFormFieldKind.Signature, signature.Kind);
        Assert.True(signature.IsSignatureField);

        Assert.Same(text, Assert.Single(logical.GetFormFields(PdfFormFieldKind.Text)));
        Assert.Equal(new[] { "Accept", "Choice", "Submit" }, logical.GetFormFields(PdfFormFieldKind.Button).Select(field => field.Name).ToArray());
        Assert.Same(choice, Assert.Single(logical.FormFieldsByKind[PdfFormFieldKind.Choice]));
        Assert.Same(signature, Assert.Single(logical.GetFormFields(PdfFormFieldKind.Signature)));
        Assert.Empty(logical.GetFormFields(PdfFormFieldKind.Unknown));
    }

    [Fact]
    public void Load_ExposesAcroFormWidgetGeometry() {
        PdfLogicalDocument logical = PdfLogicalDocument.Load(BuildWidgetFormPdf());

        PdfFormField field = Assert.Single(logical.FormFields);
        Assert.Equal("AcceptTerms", field.Name);
        Assert.Equal("Btn", field.FieldType);
        Assert.Equal("Yes", field.Value);
        Assert.True(field.HasWidgets);
        Assert.True(field.HasPageNumbers);
        Assert.Equal(1, field.PageNumberCount);
        Assert.Equal(new[] { 1 }, field.PageNumbers);
        Assert.Same(field, Assert.Single(logical.FormFieldsByPageNumber[1]));
        Assert.Same(field, Assert.Single(logical.GetFormFields(1)));
        Assert.Empty(logical.GetFormFields(2));
        Assert.Throws<ArgumentOutOfRangeException>(() => logical.GetFormFields(0));

        PdfFormWidget widget = Assert.Single(field.Widgets);
        Assert.Same(widget, Assert.Single(field.WidgetsByPageNumber[1]));
        Assert.Same(widget, Assert.Single(field.GetWidgets(1)));
        Assert.Empty(field.GetWidgets(2));
        Assert.Throws<ArgumentOutOfRangeException>(() => field.GetWidgets(0));
        Assert.Equal("AcceptTerms", widget.FieldName);
        Assert.Equal(8, widget.ObjectNumber);
        Assert.Equal(1, widget.PageNumber);
        Assert.Equal(20, widget.X1);
        Assert.Equal(100, widget.Y1);
        Assert.Equal(36, widget.X2);
        Assert.Equal(116, widget.Y2);
        Assert.Equal(16, widget.Width);
        Assert.Equal(16, widget.Height);
        Assert.Equal("Yes", widget.AppearanceState);
        Assert.Equal(4, widget.Flags);
        Assert.Equal(new[] { "Off", "Yes" }, widget.NormalAppearanceStates);
        Assert.True(widget.HasNormalAppearanceState("Yes"));

        PdfLogicalPage page = Assert.Single(logical.Pages);
        PdfLogicalFormWidget logicalWidget = Assert.Single(page.FormWidgets);
        Assert.Same(field, logicalWidget.Field);
        Assert.Same(widget, logicalWidget.SourceWidget);
        Assert.Equal(PdfLogicalElementKind.FormWidget, logicalWidget.Kind);
        Assert.Equal("AcceptTerms", logicalWidget.FieldName);
        Assert.Equal("Btn", logicalWidget.FieldType);
        Assert.Equal("Yes", logicalWidget.Value);
        Assert.Equal(8, logicalWidget.ObjectNumber);
        Assert.Equal(1, logicalWidget.PageNumber);
        Assert.Equal(20, logicalWidget.X1);
        Assert.Equal(100, logicalWidget.Y1);
        Assert.Equal(36, logicalWidget.X2);
        Assert.Equal(116, logicalWidget.Y2);
        Assert.True(logicalWidget.IsPrint);
        Assert.False(logicalWidget.IsHidden);
        Assert.False(logicalWidget.IsNoView);
        Assert.False(logicalWidget.IsLocked);
        Assert.True(logicalWidget.HasNormalAppearanceStates);
        Assert.Equal(2, logicalWidget.NormalAppearanceStateCount);
        Assert.Equal(new[] { "Off", "Yes" }, logicalWidget.NormalAppearanceStates);
        Assert.True(logicalWidget.HasNormalAppearanceState("Off"));
        Assert.True(logical.HasFormWidgets);
        Assert.Same(logicalWidget, Assert.Single(logical.FormWidgets));
        Assert.Same(logicalWidget, Assert.Single(logical.FormWidgetsByFieldName["AcceptTerms"]));
        Assert.Same(logicalWidget, Assert.Single(logical.FormWidgetsByPageNumber[1]));
        Assert.Same(logicalWidget, Assert.Single(logical.GetFormWidgets("AcceptTerms")));
        Assert.Same(logicalWidget, Assert.Single(logical.GetFormWidgets(1)));
        Assert.Empty(logical.GetFormWidgets("Missing"));
        Assert.Empty(logical.GetFormWidgets(2));
        Assert.Throws<ArgumentOutOfRangeException>(() => logical.GetFormWidgets(0));
        Assert.Contains(page.Elements, element => element.Kind == PdfLogicalElementKind.FormWidget);
        Assert.Contains(logical.Elements, element => element.Kind == PdfLogicalElementKind.FormWidget);
    }
}

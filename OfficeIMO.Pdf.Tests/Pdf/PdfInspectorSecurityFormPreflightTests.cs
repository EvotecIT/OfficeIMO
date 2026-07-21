using System;
using System.IO;
using System.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfInspectorTests {
    [Fact]
    public void Preflight_BlocksEncryptedPdfBeforeFullInspection() {
        PdfDocumentPreflight report = PdfInspector.Preflight(BuildEncryptedPdfMarker());

        Assert.False(report.CanRead);
        Assert.False(report.CanRewrite);
        Assert.False(report.CanExtractText);
        Assert.False(report.CanExtractImages);
        Assert.False(report.CanReadLogicalObjects);
        Assert.False(report.CanManipulatePages);
        Assert.False(report.CanFillSimpleFormFields);
        Assert.False(report.CanFlattenSimpleFormFields);
        Assert.False(report.CanFillAndFlattenSimpleFormFields);
        Assert.False(report.Can(PdfPreflightCapability.ExtractText));
        Assert.False(report.Can(PdfPreflightCapability.ExtractImages));
        Assert.False(report.Can(PdfPreflightCapability.ReadLogicalObjects));
        Assert.False(report.Can(PdfPreflightCapability.ManipulatePages));
        Assert.False(report.Can(PdfPreflightCapability.FillSimpleFormFields));
        Assert.Contains("PDF encryption dictionary could not be read.", report.GetCapabilityDiagnostics(PdfPreflightCapability.ExtractText));
        Assert.Contains("PDF encryption dictionary could not be read.", report.GetCapabilityDiagnostics(PdfPreflightCapability.ExtractImages));
        Assert.Contains("PDF encryption dictionary could not be read.", report.GetCapabilityDiagnostics(PdfPreflightCapability.ReadLogicalObjects));
        Assert.Contains("Encrypted input requires operation-specific planning. Authenticated unsigned PDFs support proven page extraction, merge, page-tree, and page-content rewrites; other rewrites remain blocked, and security changes require owner authorization.", report.GetCapabilityDiagnostics(PdfPreflightCapability.ManipulatePages));
        Assert.Contains("PDF encryption dictionary could not be read.", report.GetCapabilityDiagnostics(PdfPreflightCapability.FillSimpleFormFields));
        Assert.Null(report.DocumentInfo);
        Assert.True(report.Probe.HasEncryption);
        Assert.Contains("PDF encryption dictionary could not be read.", report.Diagnostics);
        AssertReadBlocker(report, PdfReadBlockerKind.Encryption, "PDF encryption dictionary could not be read.");
        AssertRewriteBlocker(report, PdfRewriteBlockerKind.Encryption, "Encrypted input requires operation-specific planning. Authenticated unsigned PDFs support proven page extraction, merge, page-tree, and page-content rewrites; other rewrites remain blocked, and security changes require owner authorization.");
    }

    [Fact]
    public void Preflight_AllowsSignedPdfReadButBlocksRewrite() {
        PdfDocumentPreflight report = PdfInspector.Preflight(BuildSignedPdfMarker());

        Assert.True(report.CanRead);
        Assert.False(report.CanRewrite);
        Assert.True(report.CanExtractText);
        Assert.True(report.CanExtractImages);
        Assert.True(report.CanReadLogicalObjects);
        Assert.False(report.CanManipulatePages);
        Assert.False(report.CanFillSimpleFormFields);
        Assert.False(report.CanFlattenSimpleFormFields);
        Assert.False(report.CanFillAndFlattenSimpleFormFields);
        Assert.False(report.Can(PdfPreflightCapability.FillSimpleFormFields));
        Assert.False(report.Can(PdfPreflightCapability.FlattenSimpleFormFields));
        Assert.Contains(
            "Signed PDF files are not supported for form filling or flattening by OfficeIMO.Pdf yet.",
            report.GetCapabilityDiagnostics(PdfPreflightCapability.FillSimpleFormFields));
        Assert.NotNull(report.DocumentInfo);
        Assert.True(report.Probe.HasSignatures);
        Assert.True(report.DocumentInfo!.HasSignatures);
        Assert.True(report.Probe.HasForms);
        Assert.True(report.DocumentInfo.HasForms);
        Assert.True(report.DocumentInfo.HasAcroFormSignatureFlags);
        Assert.Equal(3, report.DocumentInfo.AcroFormSignatureFlags);
        Assert.True(report.DocumentInfo.AcroFormSignaturesExist);
        Assert.True(report.DocumentInfo.AcroFormAppendOnly);
        Assert.Empty(report.ReadBlockers);
        Assert.Contains("Signed PDF files are not supported for rewriting by OfficeIMO.Pdf yet.", report.Diagnostics);
        AssertRewriteBlocker(report, PdfRewriteBlockerKind.Signatures, "Signed PDF files are not supported for rewriting by OfficeIMO.Pdf yet.");
        AssertRewriteBlocker(report, PdfRewriteBlockerKind.Forms, "PDF form fields are not supported for rewriting by OfficeIMO.Pdf yet.");
    }

    [Fact]
    public void Preflight_AllowsFormPdfReadButBlocksRewrite() {
        PdfDocumentPreflight report = PdfInspector.Preflight(BuildFormPdfMarker());

        Assert.True(report.CanRead);
        Assert.False(report.CanRewrite);
        Assert.True(report.CanExtractText);
        Assert.True(report.CanExtractImages);
        Assert.True(report.CanReadLogicalObjects);
        Assert.False(report.CanManipulatePages);
        Assert.True(report.CanFillSimpleFormFields);
        Assert.False(report.CanFlattenSimpleFormFields);
        Assert.False(report.CanFillAndFlattenSimpleFormFields);
        Assert.True(report.Can(PdfPreflightCapability.FillSimpleFormFields));
        Assert.False(report.Can(PdfPreflightCapability.FlattenSimpleFormFields));
        Assert.False(report.Can(PdfPreflightCapability.FillAndFlattenSimpleFormFields));
        Assert.Empty(report.GetCapabilityDiagnostics(PdfPreflightCapability.FillSimpleFormFields));
        Assert.Contains(
            "PDF does not contain named text, choice, or button AcroForm widgets with readable page-backed rectangles supported for simple form flattening by OfficeIMO.Pdf.",
            report.GetCapabilityDiagnostics(PdfPreflightCapability.FlattenSimpleFormFields));
        Assert.NotNull(report.DocumentInfo);
        Assert.False(report.Probe.HasSignatures);
        Assert.True(report.Probe.HasForms);
        Assert.True(report.DocumentInfo!.HasForms);
        Assert.True(report.DocumentInfo.HasAcroFormNeedAppearances);
        Assert.Equal(true, report.DocumentInfo.AcroFormNeedAppearances);
        Assert.True(report.DocumentInfo.RequiresAcroFormAppearanceRegeneration);
        Assert.False(report.DocumentInfo.HasAcroFormSignatureFlags);
        Assert.False(report.DocumentInfo.AcroFormSignaturesExist);
        Assert.False(report.DocumentInfo.AcroFormAppendOnly);
        Assert.True(report.DocumentInfo.HasAcroFormDefaultAppearance);
        Assert.Equal("/Helv 11 Tf 0 g", report.DocumentInfo.AcroFormDefaultAppearance);
        Assert.True(report.DocumentInfo.HasReadableFormFields);
        Assert.Equal(1, report.DocumentInfo.FormFieldCount);
        Assert.Equal("Name", report.DocumentInfo.FormFields[0].Name);
        Assert.Equal("Tx", report.DocumentInfo.FormFields[0].FieldType);
        Assert.Equal("OfficeIMO", report.DocumentInfo.FormFields[0].Value);
        Assert.Equal("/Helv 11 Tf 0 g", report.DocumentInfo.FormFields[0].DefaultAppearance);
        Assert.Equal(new[] { "Name" }, report.DocumentInfo.FormFieldNames);
        Assert.Empty(report.ReadBlockers);
        Assert.Contains("PDF form fields are not supported for rewriting by OfficeIMO.Pdf yet.", report.Diagnostics);
        AssertRewriteBlocker(report, PdfRewriteBlockerKind.Forms, "PDF form fields are not supported for rewriting by OfficeIMO.Pdf yet.");
    }

    [Fact]
    public void Preflight_ReportsSimpleFormMutationGatesForWrappers() {
        PdfDocumentPreflight report = PdfInspector.Preflight(BuildWidgetFormPdf());

        Assert.True(report.CanRead);
        Assert.False(report.CanRewrite);
        Assert.True(report.CanExtractText);
        Assert.True(report.CanExtractImages);
        Assert.True(report.CanReadLogicalObjects);
        Assert.False(report.CanManipulatePages);
        Assert.True(report.CanFillSimpleFormFields);
        Assert.True(report.CanFlattenSimpleFormFields);
        Assert.True(report.CanFillAndFlattenSimpleFormFields);
        Assert.True(report.Can(PdfPreflightCapability.FillSimpleFormFields));
        Assert.True(report.Can(PdfPreflightCapability.FlattenSimpleFormFields));
        Assert.True(report.Can(PdfPreflightCapability.FillAndFlattenSimpleFormFields));
        Assert.Empty(report.GetCapabilityDiagnostics(PdfPreflightCapability.FillSimpleFormFields));
        Assert.Empty(report.GetCapabilityDiagnostics(PdfPreflightCapability.FlattenSimpleFormFields));
        Assert.Empty(report.GetCapabilityDiagnostics(PdfPreflightCapability.FillAndFlattenSimpleFormFields));
        Assert.Contains("PDF form fields are not supported for rewriting by OfficeIMO.Pdf yet.", report.GetCapabilityDiagnostics(PdfPreflightCapability.ManipulatePages));
        Assert.NotNull(report.DocumentInfo);
        Assert.True(report.DocumentInfo!.HasReadableFormFields);
        Assert.True(report.DocumentInfo.HasFormWidgets);
        Assert.Single(report.DocumentInfo.FormFields);
        Assert.Single(report.DocumentInfo.FormWidgets);
        Assert.Empty(report.ReadBlockers);
        AssertRewriteBlocker(report, PdfRewriteBlockerKind.Forms, "PDF form fields are not supported for rewriting by OfficeIMO.Pdf yet.");
    }

    [Fact]
    public void Inspect_ReadsHierarchicalAcroFormFieldNamesForWrappers() {
        PdfDocumentInfo info = PdfInspector.Inspect(BuildHierarchicalFormPdf());

        Assert.True(info.HasForms);
        Assert.True(info.HasReadableFormFields);
        Assert.True(info.HasAcroFormNeedAppearances);
        Assert.Equal(false, info.AcroFormNeedAppearances);
        Assert.False(info.RequiresAcroFormAppearanceRegeneration);
        Assert.False(info.HasAcroFormSignatureFlags);
        Assert.False(info.AcroFormSignaturesExist);
        Assert.False(info.AcroFormAppendOnly);
        Assert.True(info.HasAcroFormDefaultAppearance);
        Assert.Equal("/Helv 7 Tf 0.5 g", info.AcroFormDefaultAppearance);
        Assert.Equal(3, info.FormFieldCount);
        Assert.Equal(3, info.FormFieldsByName.Count);
        Assert.Equal(new[] { "Person.Name", "AcceptTerms", "Selection.Country" }, info.FormFieldNames);
        Assert.Equal("Person.Name", info.FormFields[0].Name);
        Assert.Equal("Name", info.FormFields[0].PartialName);
        Assert.Equal("Tx", info.FormFields[0].FieldType);
        Assert.Equal("OfficeIMO", info.FormFields[0].Value);
        Assert.Equal("Display name", info.FormFields[0].AlternateName);
        Assert.Equal("ExportName", info.FormFields[0].MappingName);
        Assert.Equal(1, info.FormFields[0].Flags);
        Assert.Equal(64, info.FormFields[0].MaxLength);
        Assert.Equal("InheritedDraft", info.FormFields[0].DefaultValue);
        Assert.Equal(new[] { "InheritedDraft" }, info.FormFields[0].DefaultValues);
        Assert.Equal("/Helv 10 Tf 0 g", info.FormFields[0].DefaultAppearance);
        Assert.True(info.FormFields[0].HasDefaultAppearance);
        Assert.Equal(2, info.FormFields[0].Quadding);
        Assert.Equal(PdfFormFieldTextAlignment.Right, info.FormFields[0].TextAlignment);
        Assert.Equal("AcceptTerms", info.FormFields[1].Name);
        Assert.Equal("Btn", info.FormFields[1].FieldType);
        Assert.Equal("Yes", info.FormFields[1].Value);
        Assert.Equal("Selection.Country", info.FormFields[2].Name);
        Assert.Equal("Ch", info.FormFields[2].FieldType);
        Assert.Equal("DE", info.FormFields[2].Value);
        Assert.Equal("PL", info.FormFields[2].DefaultValue);
        Assert.Equal(new[] { "DE" }, info.FormFields[2].Values);
        Assert.Equal(new[] { "PL" }, info.FormFields[2].DefaultValues);
        Assert.Equal(2, info.FormFields[2].OptionCount);
        Assert.Equal(new[] { "DE" }, info.FormFields[2].SelectedOptions.Select(option => option.ExportValue).ToArray());
        Assert.Equal(new[] { "PL" }, info.FormFields[2].DefaultSelectedOptions.Select(option => option.ExportValue).ToArray());
        Assert.True(info.TryGetFormField("Person.Name", out PdfFormField? nameField));
        Assert.Same(info.FormFields[0], nameField);
        Assert.Equal("OfficeIMO", nameField!.Value);
        Assert.True(info.TryGetFormField("AcceptTerms", out PdfFormField? acceptField));
        Assert.Same(info.FormFields[1], acceptField);
        Assert.True(info.TryGetFormField("Selection.Country", out PdfFormField? countryField));
        Assert.True(countryField!.IsChoiceField);
        Assert.False(info.TryGetFormField("Missing", out PdfFormField? missingField));
        Assert.Null(missingField);
        Assert.Same(info.FormFields[0], Assert.Single(info.GetFormFields(PdfFormFieldKind.Text)));
        Assert.Same(info.FormFields[1], Assert.Single(info.GetFormFields(PdfFormFieldKind.Button)));
        Assert.Same(info.FormFields[2], Assert.Single(info.FormFieldsByKind[PdfFormFieldKind.Choice]));
        Assert.Empty(info.GetFormFields(PdfFormFieldKind.Signature));
    }

    [Fact]
    public void Inspect_ReadsAcroFormChoiceOptionsAndTextConstraintsForWrappers() {
        PdfDocumentInfo info = PdfInspector.Inspect(BuildChoiceAndTextConstraintFormPdf());

        PdfFormField text = Assert.Single(info.FormFields, field => field.Name == "Notes");
        Assert.Equal(PdfFormFieldKind.Text, text.Kind);
        Assert.Equal(42, text.MaxLength);
        Assert.Equal(new[] { "Secret" }, text.Values);
        Assert.Equal("Draft", text.DefaultValue);
        Assert.Equal(new[] { "Draft" }, text.DefaultValues);
        Assert.True(text.HasDefaultValues);
        Assert.Equal("/Helv 9 Tf 0 g", text.DefaultAppearance);
        Assert.Equal(0, text.Quadding);
        Assert.Equal(PdfFormFieldTextAlignment.Left, text.TextAlignment);
        Assert.False(text.HasOptions);
        Assert.False(text.HasDefaultSelectedOptions);

        PdfFormField choice = Assert.Single(info.FormFields, field => field.Name == "Country");
        Assert.Equal(PdfFormFieldKind.Choice, choice.Kind);
        Assert.Equal("[PL US]", choice.Value);
        Assert.Equal(new[] { "PL", "US" }, choice.Values);
        Assert.Equal("[DE US]", choice.DefaultValue);
        Assert.Equal(new[] { "DE", "US" }, choice.DefaultValues);
        Assert.Equal("/Helv 8 Tf 0 0 1 rg", choice.DefaultAppearance);
        Assert.Equal(1, choice.Quadding);
        Assert.Equal(PdfFormFieldTextAlignment.Center, choice.TextAlignment);
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
        Assert.Equal(2, choice.SelectedOptionCount);
        Assert.Equal(new[] { "PL", "US" }, choice.SelectedOptions.Select(option => option.ExportValue).ToArray());
        Assert.Equal(2, choice.DefaultSelectedOptionCount);
        Assert.Equal(new[] { "DE", "US" }, choice.DefaultSelectedOptions.Select(option => option.ExportValue).ToArray());
    }

    [Fact]
    public void Inspect_ReadsAcroFormWidgetGeometryForWrappers() {
        PdfDocumentInfo info = PdfInspector.Inspect(BuildWidgetFormPdf());

        PdfFormField field = Assert.Single(info.FormFields);
        Assert.Equal("AcceptTerms", field.Name);
        Assert.True(field.HasWidgets);
        Assert.True(field.HasPageNumbers);
        Assert.Equal(1, field.PageNumberCount);
        Assert.Equal(new[] { 1 }, field.PageNumbers);
        Assert.True(info.HasFormWidgets);
        Assert.Equal(1, info.FormWidgetCount);
        Assert.True(info.Pages[0].HasFormWidgets);
        Assert.Same(field, Assert.Single(info.FormFieldsByPageNumber[1]));
        Assert.Same(field, Assert.Single(info.GetFormFields(1)));
        Assert.Empty(info.GetFormFields(2));
        Assert.Throws<ArgumentOutOfRangeException>(() => info.GetFormFields(0));

        PdfFormWidget widget = Assert.Single(field.Widgets);
        Assert.Same(widget, Assert.Single(field.WidgetsByPageNumber[1]));
        Assert.Same(widget, Assert.Single(field.GetWidgets(1)));
        Assert.Empty(field.GetWidgets(2));
        Assert.Throws<ArgumentOutOfRangeException>(() => field.GetWidgets(0));
        Assert.Same(widget, Assert.Single(info.FormWidgets));
        Assert.Same(widget, Assert.Single(info.Pages[0].FormWidgets));
        Assert.Same(widget, Assert.Single(info.FormWidgetsByFieldName["AcceptTerms"]));
        Assert.Same(widget, Assert.Single(info.FormWidgetsByPageNumber[1]));
        Assert.Same(widget, Assert.Single(info.GetFormWidgets("AcceptTerms")));
        Assert.Same(widget, Assert.Single(info.GetFormWidgets(1)));
        Assert.Empty(info.GetFormWidgets("Missing"));
        Assert.Empty(info.GetFormWidgets(2));
        Assert.Throws<ArgumentOutOfRangeException>(() => info.GetFormWidgets(0));
        Assert.Equal("AcceptTerms", widget.FieldName);
        Assert.Equal(8, widget.ObjectNumber);
        Assert.Equal(1, widget.PageNumber);
        Assert.Equal(20, widget.X1);
        Assert.Equal(100, widget.Y1);
        Assert.Equal(36, widget.X2);
        Assert.Equal(116, widget.Y2);
        Assert.Equal("Yes", widget.AppearanceState);
        Assert.Equal(4, widget.Flags);
        Assert.False(widget.IsInvisible);
        Assert.False(widget.IsHidden);
        Assert.True(widget.IsPrint);
        Assert.False(widget.IsNoZoom);
        Assert.False(widget.IsNoRotate);
        Assert.False(widget.IsNoView);
        Assert.False(widget.IsReadOnly);
        Assert.False(widget.IsLocked);
        Assert.False(widget.IsToggleNoView);
        Assert.False(widget.IsLockedContents);
        Assert.True(widget.HasNormalAppearanceStates);
        Assert.Equal(2, widget.NormalAppearanceStateCount);
        Assert.Equal(new[] { "Off", "Yes" }, widget.NormalAppearanceStates);
        Assert.True(widget.HasNormalAppearanceState("Yes"));
        Assert.True(widget.HasNormalAppearanceState("Off"));
        Assert.False(widget.HasNormalAppearanceState("Maybe"));
    }


}

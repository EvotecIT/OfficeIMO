using System.Globalization;
using System.Text;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfITextInspiredCoverageTests {
    [Fact]
    public void Inspect_ReportsFormsAnnotationsPageBoxesTaggedFontsAndAppendPlan() {
        byte[] pdf = BuildCoveragePdf();

        byte[] appendablePdf = PdfDocument.Create()
            .Paragraph(paragraph => paragraph.Text("Append-only metadata plan"))
            .ToBytes();
        PdfAppendOnlyMutationReport appendPlan = PdfIncrementalUpdater.AnalyzeAppendOnlyMutation(appendablePdf);
        Assert.True(appendPlan.CanAppendMetadata);
        Assert.True(appendPlan.CanAppendFormFields);
        Assert.Contains("Metadata", appendPlan.SupportedActions);
        Assert.Contains("FormFill", appendPlan.SupportedActions);

        PdfDocumentInfo info = PdfInspector.Inspect(pdf);
        Assert.Equal(1, info.TextFormFieldCount);
        Assert.Equal(1, info.RequiredFormFieldCount);
        Assert.Equal(1, info.ReadOnlyFormFieldCount);
        Assert.Equal(1, info.FormWidgetCount);
        Assert.True(info.HasProductionPageBoxes);
        Assert.Equal(1, info.TrimBoxPageCount);
        Assert.Equal(1, info.BleedBoxPageCount);
        Assert.Equal(1, info.ArtBoxPageCount);
        Assert.Equal(1, info.ActiveAnnotationCount);
        Assert.Equal(1, info.RiskyAnnotationActionCount);
        Assert.True(info.AnnotationSubtypeCounts["Text"] >= 1);

        PdfAnnotation note = Assert.Single(info.GetAnnotationsBySubtype("Text"));
        Assert.Equal("Review note", note.Contents);
        Assert.Equal("Note-1", note.Name);
        Assert.Equal("Reviewer", note.Title);
        Assert.True(note.IsLocked);
        Assert.True(note.HasColor);
        Assert.Equal("Launch", Assert.Single(note.AdditionalActions).ActionType);

        PdfTaggedContentInfo tagged = Assert.IsType<PdfTaggedContentInfo>(info.TaggedContent);
        Assert.True(tagged.HasRoleMap);
        Assert.True(tagged.HasDeepTaggedPdfEvidence);
        Assert.Equal(1, tagged.LanguageElementCount);
        Assert.Equal(0, tagged.AlternateTextElementCount);
        Assert.Equal(1, tagged.FigureWithoutAlternateTextCount);

        PdfDiagnosticReport diagnostics = PdfDiagnostics.Analyze(pdf);
        Assert.True(diagnostics.FontCount >= 2);
        Assert.Contains(diagnostics.Fonts, font => font.ObjectNumber == 4 && font.IsStandardBase14Font);
        Assert.Contains(diagnostics.Fonts, font => font.ObjectNumber == 14 && font.HasEmbeddedFontFile && font.EmbeddedFontFileKind == "FontFile2");
        Assert.Contains(diagnostics.Fonts, font => font.ObjectNumber == 14 && font.RepairReadiness == "Ready");

        PdfOptimizationReport optimization = PdfDiagnostics.AnalyzeOptimization(pdf);
        Assert.True(optimization.StreamCount > 0);
        Assert.True(optimization.TotalStreamBytes > 0);
        Assert.True(optimization.LargestStreamBytes > 0);
        Assert.True(optimization.FindingCount >= 0);
    }

    [Fact]
    public void AssessProof_ReportsMissingExternalValidationStatus() {
        var options = new PdfOptions {
            ComplianceProfile = PdfComplianceProfile.PdfA3B
        };

        PdfComplianceProofReport proof = PdfComplianceAnalyzer.AssessProof(PdfComplianceProfile.PdfA3B, options);

        Assert.True(proof.RequiresExternalValidation);
        Assert.True(proof.RequiredExternalValidatorCount > 0);
        Assert.Equal("InternalGaps", proof.ProofStatus);
        Assert.Contains("Missing external validation", proof.ExternalProofSummary, StringComparison.Ordinal);
        Assert.False(proof.CanClaimConformance);
    }

    [Fact]
    public void IncrementalUpdater_AppendsSimpleFormFieldRevision() {
        byte[] pdf = BuildCoveragePdf();

        byte[] updated = PdfIncrementalUpdater.UpdateFormFields(pdf, new Dictionary<string, string> {
            ["Name"] = "Grace"
        });

        Assert.True(updated.Length > pdf.Length);
        PdfDocumentInfo info = PdfInspector.Inspect(updated);
        PdfFormField field = Assert.Single(info.FormFields);
        Assert.Equal("Grace", field.Value);
        Assert.True(info.AcroFormNeedAppearances);
        Assert.True(info.Security.HasIncrementalUpdates);
    }

    [Fact]
    public void IncrementalUpdater_AppendsFormAppearanceStreams() {
        byte[] pdf = BuildCoveragePdf();

        byte[] updated = PdfIncrementalUpdater.UpdateFormFields(pdf, new Dictionary<string, string> {
            ["Name"] = "Grace"
        }, new PdfIncrementalFormFieldUpdateOptions {
            GenerateAppearanceStreams = true,
            KeepNeedAppearances = false
        });

        string text = PdfEncoding.Latin1GetString(updated);
        PdfDocumentInfo info = PdfInspector.Inspect(updated);
        PdfFormField field = Assert.Single(info.FormFields);
        Assert.Equal("Grace", field.Value);
        Assert.Equal(false, info.AcroFormNeedAppearances);
        Assert.Contains("/AP", text, StringComparison.Ordinal);
        Assert.Contains("/Subtype /Form", text, StringComparison.Ordinal);
        Assert.Contains("/BaseFont /Helvetica", text, StringComparison.Ordinal);
    }

    [Fact]
    public void IncrementalUpdater_WrapsGeneratedMultilineFormAppearanceStreams() {
        byte[] pdf = BuildMultilineFormPdf();

        byte[] updated = PdfIncrementalUpdater.UpdateFormFields(pdf, new Dictionary<string, string> {
            ["Notes"] = "Alpha Bravo Charlie Delta"
        }, new PdfIncrementalFormFieldUpdateOptions {
            GenerateAppearanceStreams = true,
            KeepNeedAppearances = false
        });

        string appended = PdfEncoding.Latin1GetString(updated).Substring(PdfEncoding.Latin1GetString(pdf).Length);
        PdfDocumentInfo info = PdfInspector.Inspect(updated);
        PdfFormField field = Assert.Single(info.FormFields);

        Assert.True(field.IsMultiline);
        Assert.Equal("Alpha Bravo Charlie Delta", field.Value);
        Assert.Equal(false, info.AcroFormNeedAppearances);
        Assert.Contains("/AP", appended, StringComparison.Ordinal);
        Assert.DoesNotContain("<416C70686120427261766F20436861726C69652044656C7461> Tj", appended, StringComparison.Ordinal);
        Assert.True(System.Text.RegularExpressions.Regex.Matches(appended, @"BT /Helv 12 Tf .* Tj ET").Count > 1);
    }

    [Fact]
    public void IncrementalUpdater_UsesCmykDefaultAppearanceTextColor() {
        byte[] pdf = BuildDefaultAppearanceFormPdf("/Helv 12 Tf 0.1 0.2 0 0.1 k");

        byte[] updated = PdfIncrementalUpdater.UpdateFormFields(pdf, new Dictionary<string, string> {
            ["Name"] = "Color"
        }, new PdfIncrementalFormFieldUpdateOptions {
            GenerateAppearanceStreams = true,
            KeepNeedAppearances = false
        });

        string appended = PdfEncoding.Latin1GetString(updated).Substring(PdfEncoding.Latin1GetString(pdf).Length);
        PdfFormField field = Assert.Single(PdfInspector.Inspect(updated).FormFields);

        Assert.Equal("Color", field.Value);
        Assert.Contains("0.81 0.72 0.9 rg", appended, StringComparison.Ordinal);
    }

    [Fact]
    public void IncrementalUpdater_UsesDefaultAppearanceFontSize() {
        byte[] pdf = BuildDefaultAppearanceFormPdf("/Helv 7.5 Tf 0.1 0.2 0 0.1 k");

        byte[] updated = PdfIncrementalUpdater.UpdateFormFields(pdf, new Dictionary<string, string> {
            ["Name"] = "Small"
        }, new PdfIncrementalFormFieldUpdateOptions {
            GenerateAppearanceStreams = true,
            KeepNeedAppearances = false
        });

        string appended = PdfEncoding.Latin1GetString(updated).Substring(PdfEncoding.Latin1GetString(pdf).Length);
        PdfFormField field = Assert.Single(PdfInspector.Inspect(updated).FormFields);

        Assert.Equal("Small", field.Value);
        Assert.Contains("BT /Helv 7.5 Tf 0.81 0.72 0.9 rg", appended, StringComparison.Ordinal);
        Assert.DoesNotContain("BT /Helv 12 Tf 0.81 0.72 0.9 rg", appended, StringComparison.Ordinal);
    }

    [Fact]
    public void IncrementalUpdater_PreservesWidgetBorderArrayWidthInGeneratedAppearance() {
        byte[] pdf = BuildDefaultAppearanceFormPdf(
            "/Helv 12 Tf 0 g",
            " /MK << /BC [0 0 1] /BG [0.95 0.95 1] >> /Border [0 0 2.5]");

        byte[] updated = PdfIncrementalUpdater.UpdateFormFields(pdf, new Dictionary<string, string> {
            ["Name"] = "Border"
        }, new PdfIncrementalFormFieldUpdateOptions {
            GenerateAppearanceStreams = true,
            KeepNeedAppearances = false
        });

        string appended = PdfEncoding.Latin1GetString(updated).Substring(PdfEncoding.Latin1GetString(pdf).Length);
        PdfFormField field = Assert.Single(PdfInspector.Inspect(updated).FormFields);

        Assert.Equal("Border", field.Value);
        Assert.Contains("0.95 0.95 1 rg", appended, StringComparison.Ordinal);
        Assert.Contains("0 0 1 RG 2.5 w", appended, StringComparison.Ordinal);
        Assert.DoesNotContain("0 0 1 RG 1 w", appended, StringComparison.Ordinal);
    }

    [Fact]
    public void IncrementalUpdater_PreservesWidgetBorderStyleDashPatternInGeneratedAppearance() {
        byte[] pdf = BuildDefaultAppearanceFormPdf(
            "/Helv 12 Tf 0 g",
            " /MK << /BC [0 0 1] /BG [0.95 0.95 1] >> /BS << /S /D /W 2 /D [5 1] >>");

        byte[] updated = PdfIncrementalUpdater.UpdateFormFields(pdf, new Dictionary<string, string> {
            ["Name"] = "Dash"
        }, new PdfIncrementalFormFieldUpdateOptions {
            GenerateAppearanceStreams = true,
            KeepNeedAppearances = false
        });

        string appended = PdfEncoding.Latin1GetString(updated).Substring(PdfEncoding.Latin1GetString(pdf).Length);
        PdfFormField field = Assert.Single(PdfInspector.Inspect(updated).FormFields);

        Assert.Equal("Dash", field.Value);
        Assert.Contains("0 0 1 RG 2 w [5 1] 0 d", appended, StringComparison.Ordinal);
    }

    [Fact]
    public void IncrementalUpdater_PreservesWidgetUnderlineBorderStyleInGeneratedAppearance() {
        byte[] pdf = BuildDefaultAppearanceFormPdf(
            "/Helv 12 Tf 0 g",
            " /MK << /BC [0 0 1] /BG [0.95 0.95 1] >> /BS << /S /U /W 2 >>");

        byte[] updated = PdfIncrementalUpdater.UpdateFormFields(pdf, new Dictionary<string, string> {
            ["Name"] = "Underline"
        }, new PdfIncrementalFormFieldUpdateOptions {
            GenerateAppearanceStreams = true,
            KeepNeedAppearances = false
        });

        string appended = PdfEncoding.Latin1GetString(updated).Substring(PdfEncoding.Latin1GetString(pdf).Length);
        PdfFormField field = Assert.Single(PdfInspector.Inspect(updated).FormFields);

        Assert.Equal("Underline", field.Value);
        Assert.Contains("0 0 1 RG 2 w 0 1 m 130 1 l S", appended, StringComparison.Ordinal);
        Assert.DoesNotContain("0 0 1 RG 2 w 1 1 128 18 re S", appended, StringComparison.Ordinal);
    }

    [Fact]
    public void IncrementalUpdater_PreservesWidgetInsetBorderStyleInGeneratedAppearance() {
        byte[] pdf = BuildDefaultAppearanceFormPdf(
            "/Helv 12 Tf 0 g",
            " /MK << /BC [0 0 0] /BG [0.95 0.95 1] >> /BS << /S /I /W 2 >>");

        byte[] updated = PdfIncrementalUpdater.UpdateFormFields(pdf, new Dictionary<string, string> {
            ["Name"] = "Inset"
        }, new PdfIncrementalFormFieldUpdateOptions {
            GenerateAppearanceStreams = true,
            KeepNeedAppearances = false
        });

        string appended = PdfEncoding.Latin1GetString(updated).Substring(PdfEncoding.Latin1GetString(pdf).Length);
        PdfFormField field = Assert.Single(PdfInspector.Inspect(updated).FormFields);

        Assert.Equal("Inset", field.Value);
        Assert.Contains("0 0 0 RG 2 w 1 1 m 1 19 l 129 19 l S", appended, StringComparison.Ordinal);
        Assert.Contains("0.55 0.55 0.55 RG 2 w 1 1 m 129 1 l 129 19 l S", appended, StringComparison.Ordinal);
        Assert.DoesNotContain("0 0 0 RG 2 w 1 1 128 18 re S", appended, StringComparison.Ordinal);
    }

    [Fact]
    public void IncrementalUpdater_UsesInheritedDefaultAppearanceFontSizeAndTextColor() {
        byte[] pdf = BuildAcroFormDefaultAppearanceFormPdf("/Helv 8.5 Tf 0.1 0.2 0.3 rg");

        byte[] updated = PdfIncrementalUpdater.UpdateFormFields(pdf, new Dictionary<string, string> {
            ["Name"] = "Inherited"
        }, new PdfIncrementalFormFieldUpdateOptions {
            GenerateAppearanceStreams = true,
            KeepNeedAppearances = false
        });

        string appended = PdfEncoding.Latin1GetString(updated).Substring(PdfEncoding.Latin1GetString(pdf).Length);
        PdfFormField field = Assert.Single(PdfInspector.Inspect(updated).FormFields);

        Assert.Equal("Inherited", field.Value);
        Assert.Contains("BT /Helv 8.5 Tf 0.1 0.2 0.3 rg", appended, StringComparison.Ordinal);
        Assert.DoesNotContain("BT /Helv 12 Tf 0.1 0.2 0.3 rg", appended, StringComparison.Ordinal);
    }

    [Fact]
    public void IncrementalUpdater_UsesInheritedDefaultAppearanceFontResourceName() {
        byte[] pdf = BuildAcroFormDefaultAppearanceFormPdf("/F1 8.5 Tf 0.1 0.2 0.3 rg", includeFontResource: true);

        byte[] updated = PdfIncrementalUpdater.UpdateFormFields(pdf, new Dictionary<string, string> {
            ["Name"] = "Resource"
        }, new PdfIncrementalFormFieldUpdateOptions {
            GenerateAppearanceStreams = true,
            KeepNeedAppearances = false
        });

        string appended = PdfEncoding.Latin1GetString(updated).Substring(PdfEncoding.Latin1GetString(pdf).Length);
        PdfFormField field = Assert.Single(PdfInspector.Inspect(updated).FormFields);

        Assert.Equal("Resource", field.Value);
        Assert.Contains("BT /F1 8.5 Tf 0.1 0.2 0.3 rg", appended, StringComparison.Ordinal);
        Assert.DoesNotContain("BT /Helv 8.5 Tf 0.1 0.2 0.3 rg", appended, StringComparison.Ordinal);
        Assert.Contains("/Font << /F1 4 0 R >>", appended, StringComparison.Ordinal);
    }

    [Fact]
    public void IncrementalUpdater_UsesWidgetPageDefaultAppearanceFontResourceName() {
        byte[] pdf = BuildAcroFormDefaultAppearanceFormPdf("/F1 8.5 Tf 0.1 0.2 0.3 rg", includeWidgetPageReference: true);

        byte[] updated = PdfIncrementalUpdater.UpdateFormFields(pdf, new Dictionary<string, string> {
            ["Name"] = "PageResource"
        }, new PdfIncrementalFormFieldUpdateOptions {
            GenerateAppearanceStreams = true,
            KeepNeedAppearances = false
        });

        string appended = PdfEncoding.Latin1GetString(updated).Substring(PdfEncoding.Latin1GetString(pdf).Length);
        PdfFormField field = Assert.Single(PdfInspector.Inspect(updated).FormFields);

        Assert.Equal("PageResource", field.Value);
        Assert.Contains("BT /F1 8.5 Tf 0.1 0.2 0.3 rg", appended, StringComparison.Ordinal);
        Assert.DoesNotContain("BT /Helv 8.5 Tf 0.1 0.2 0.3 rg", appended, StringComparison.Ordinal);
        Assert.Contains("/Font << /F1 4 0 R >>", appended, StringComparison.Ordinal);
        Assert.DoesNotContain("/BaseFont /Helvetica /Encoding /WinAnsiEncoding", appended, StringComparison.Ordinal);
    }

    [Fact]
    public void IncrementalUpdater_UsesWidgetAppearanceDefaultAppearanceFontResourceName() {
        byte[] pdf = BuildAcroFormDefaultAppearanceFormPdf("/FAP 8.5 Tf 0.1 0.2 0.3 rg", includeWidgetAppearanceResource: true);

        byte[] updated = PdfIncrementalUpdater.UpdateFormFields(pdf, new Dictionary<string, string> {
            ["Name"] = "AppearanceResource"
        }, new PdfIncrementalFormFieldUpdateOptions {
            GenerateAppearanceStreams = true,
            KeepNeedAppearances = false
        });

        string appended = PdfEncoding.Latin1GetString(updated).Substring(PdfEncoding.Latin1GetString(pdf).Length);
        PdfFormField field = Assert.Single(PdfInspector.Inspect(updated).FormFields);

        Assert.Equal("AppearanceResource", field.Value);
        Assert.Contains("BT /FAP 8.5 Tf 0.1 0.2 0.3 rg", appended, StringComparison.Ordinal);
        Assert.DoesNotContain("BT /Helv 8.5 Tf 0.1 0.2 0.3 rg", appended, StringComparison.Ordinal);
        Assert.Contains("/Font << /FAP 4 0 R >>", appended, StringComparison.Ordinal);
        Assert.DoesNotContain("/BaseFont /Helvetica /Encoding /WinAnsiEncoding", appended, StringComparison.Ordinal);
    }

    [Fact]
    public void IncrementalUpdater_UpdatesButtonAppearanceStateWithoutRegeneratingAppearances() {
        byte[] pdf = PdfDocument.Create()
            .CheckBox("Accept", isChecked: false)
            .ToBytes();

        byte[] updated = PdfIncrementalUpdater.UpdateFormFields(pdf, new Dictionary<string, string> {
            ["Accept"] = "Yes"
        });

        string appended = PdfEncoding.Latin1GetString(updated).Substring(PdfEncoding.Latin1GetString(pdf).Length);
        PdfFormField field = Assert.Single(PdfInspector.Inspect(updated).FormFields);

        Assert.Equal("Yes", field.Value);
        Assert.Contains("/AS /Yes", appended, StringComparison.Ordinal);
        Assert.DoesNotContain("/Subtype /Form", appended, StringComparison.Ordinal);
    }

    [Fact]
    public void IncrementalUpdater_ResolvesCheckboxTruthyValueToAvailableAppearanceState() {
        byte[] pdf = PdfDocument.Create()
            .CheckBox("Accept", isChecked: false)
            .ToBytes();

        byte[] updated = PdfIncrementalUpdater.UpdateFormFields(pdf, new Dictionary<string, string> {
            ["Accept"] = "true"
        });

        string appended = PdfEncoding.Latin1GetString(updated).Substring(PdfEncoding.Latin1GetString(pdf).Length);
        PdfFormField field = Assert.Single(PdfInspector.Inspect(updated).FormFields);

        Assert.Equal("Yes", field.Value);
        Assert.Contains("/V /Yes", appended, StringComparison.Ordinal);
        Assert.Contains("/AS /Yes", appended, StringComparison.Ordinal);
        Assert.DoesNotContain("/AS /true", appended, StringComparison.Ordinal);
    }

    [Fact]
    public void IncrementalUpdater_PreservesRadioWidgetOnStatesWhenRegeneratingAppearances() {
        byte[] pdf = PdfDocument.Create()
            .RadioButtonGroup("Payment.Method", new[] { "Card", "Cash", "Wire" }, value: "Card")
            .ToBytes();

        byte[] updated = PdfIncrementalUpdater.UpdateFormFields(pdf, new Dictionary<string, string> {
            ["Payment.Method"] = "Cash"
        }, new PdfIncrementalFormFieldUpdateOptions {
            GenerateAppearanceStreams = true,
            KeepNeedAppearances = false
        });

        byte[] updatedAgain = PdfIncrementalUpdater.UpdateFormFields(updated, new Dictionary<string, string> {
            ["Payment.Method"] = "Card"
        }, new PdfIncrementalFormFieldUpdateOptions {
            GenerateAppearanceStreams = true,
            KeepNeedAppearances = false
        });

        string appended = PdfEncoding.Latin1GetString(updatedAgain).Substring(PdfEncoding.Latin1GetString(updated).Length);
        PdfFormField field = Assert.Single(PdfInspector.Inspect(updatedAgain).FormFields);
        Assert.Equal("Card", field.Value);
        Assert.Contains(" c S", appended, StringComparison.Ordinal);
        Assert.Contains(" c f", appended, StringComparison.Ordinal);
        Assert.DoesNotContain("1.25 w", appended, StringComparison.Ordinal);
    }

    [Fact]
    public void IncrementalUpdater_RejectsInvalidNonEditableChoiceValue() {
        byte[] pdf = PdfDocument.Create()
            .ChoiceField("Country", new[] { "PL", "DE" }, value: "PL")
            .ToBytes();

        Assert.Throws<ArgumentException>(() => PdfIncrementalUpdater.UpdateFormFields(pdf, new Dictionary<string, string> {
            ["Country"] = "US"
        }));
    }

    [Fact]
    public void IncrementalUpdater_AppendsMultiSelectChoiceArrayAndRowAppearances() {
        byte[] pdf = PdfDocument.Create()
            .MultiSelectChoiceField("Countries", new[] { "Poland", "Germany", "United States" }, values: new[] { "Poland" }, width: 190, height: 72)
            .ToBytes();

        byte[] updated = PdfIncrementalUpdater.UpdateFormFields(pdf, new Dictionary<string, PdfFormFieldValue> {
            ["Countries"] = PdfFormFieldValue.FromValues("Germany", "United States")
        }, new PdfIncrementalFormFieldUpdateOptions {
            GenerateAppearanceStreams = true,
            KeepNeedAppearances = false
        });

        string appended = PdfEncoding.Latin1GetString(updated).Substring(PdfEncoding.Latin1GetString(pdf).Length);
        PdfFormField field = Assert.Single(PdfInspector.Inspect(updated).FormFields);

        Assert.Equal(new[] { "Germany", "United States" }, field.Values);
        Assert.Equal(new[] { "Germany", "United States" }, field.SelectedOptions.Select(option => option.DisplayText).ToArray());
        Assert.Contains("/V [", appended, StringComparison.Ordinal);
        Assert.Contains("<4765726D616E79> Tj", appended, StringComparison.Ordinal);
        Assert.Contains("<556E6974656420537461746573> Tj", appended, StringComparison.Ordinal);
        Assert.DoesNotContain("<4765726D616E792C20556E6974656420537461746573> Tj", appended, StringComparison.Ordinal);
    }

    [Fact]
    public void IncrementalUpdater_AllowsDocMDPCertifiedFormFillWhenPermissionPermits() {
        byte[] pdf = BuildDocMdpFormPdf(permissionLevel: 2);

        PdfAppendOnlyMutationReport plan = PdfIncrementalUpdater.AnalyzeAppendOnlyMutation(pdf);
        PdfDocumentPreflight preflight = PdfInspector.Preflight(pdf);
        Assert.True(plan.CanAppendFormFields);
        Assert.False(plan.CanAppendMetadata);
        Assert.Contains("SignedDocMDPFormFill", plan.Warnings);
        Assert.True(preflight.RequiresAppendOnlyMutation);
        Assert.True(preflight.CanAppendOnlyMutate);
        Assert.Empty(preflight.AppendOnlyMutationDiagnostics);

        byte[] updated = PdfIncrementalUpdater.UpdateFormFields(pdf, new Dictionary<string, string> {
            ["Name"] = "Grace"
        }, new PdfIncrementalFormFieldUpdateOptions {
            GenerateAppearanceStreams = true,
            KeepNeedAppearances = false
        });

        PdfDocumentInfo info = PdfInspector.Inspect(updated);
        PdfFormField textField = Assert.Single(info.FormFields, static field => field.Name == "Name");
        Assert.Equal("Grace", textField.Value);
        Assert.True(info.Security.HasSignatures);
        Assert.True(info.Security.HasDocMDPPermissions);
        Assert.True(info.Security.HasIncrementalUpdates);
        Assert.Contains("/Prev", PdfEncoding.Latin1GetString(updated), StringComparison.Ordinal);
    }

    [Fact]
    public void IncrementalUpdater_BlocksDocMDPCertifiedFormFillWhenPermissionForbidsChanges() {
        byte[] pdf = BuildDocMdpFormPdf(permissionLevel: 1);

        PdfAppendOnlyMutationReport plan = PdfIncrementalUpdater.AnalyzeAppendOnlyMutation(pdf);

        Assert.False(plan.CanAppendFormFields);
        Assert.Contains("DocMDP", plan.Blockers);
        PdfMutationBlockedException exception = Assert.Throws<PdfMutationBlockedException>(() => PdfIncrementalUpdater.UpdateFormFields(pdf, new Dictionary<string, string> {
            ["Name"] = "Grace"
        }, new PdfIncrementalFormFieldUpdateOptions {
            GenerateAppearanceStreams = true,
            KeepNeedAppearances = false
        }));
        Assert.Contains("AppendOnly.DocMDP", exception.Plan.BlockerCodes);
    }

    [Fact]
    public void IncrementalUpdater_BlocksDocMDPFormFillForLockedRequestedField() {
        byte[] pdf = BuildDocMdpFormPdf(
            permissionLevel: 2,
            lockDictionary: "<< /Type /SigFieldLock /Action /Include /Fields [(Name)] >>");

        PdfMutationBlockedException exception = Assert.Throws<PdfMutationBlockedException>(() => PdfIncrementalUpdater.UpdateFormFields(pdf, new Dictionary<string, string> {
            ["Name"] = "Grace"
        }));

        Assert.Contains("AppendOnly.SignatureFieldLock", exception.Plan.BlockerCodes);
    }

    [Fact]
    public void IncrementalUpdater_AllowsHierarchicalDocMDPFieldLockExclude() {
        byte[] pdf = BuildDocMdpFormPdf(
            permissionLevel: 2,
            lockDictionary: "<< /Type /SigFieldLock /Action /Exclude /Fields [(Parent)] >>",
            fieldName: "Parent.Child");

        byte[] updated = PdfIncrementalUpdater.UpdateFormFields(pdf, new Dictionary<string, string> {
            ["Parent.Child"] = "Grace"
        });

        PdfFormField textField = Assert.Single(PdfInspector.Inspect(updated).FormFields, static field => field.Name == "Parent.Child");
        Assert.Equal("Grace", textField.Value);
    }

    [Fact]
    public void IncrementalUpdater_PreservesObjectGenerationWhenAppendingFormFieldRevision() {
        byte[] pdf = BuildGeneratedFormPdf();

        byte[] updated = PdfIncrementalUpdater.UpdateFormFields(pdf, new Dictionary<string, string> {
            ["Name"] = "Grace"
        });

        string updatedText = PdfEncoding.Latin1GetString(updated);
        int originalLength = PdfEncoding.Latin1GetString(pdf).Length;
        int appendedFieldObject = updatedText.IndexOf("5 2 obj\n", originalLength, StringComparison.Ordinal);

        Assert.True(appendedFieldObject >= originalLength);
        Assert.Contains("\n5 1\n", updatedText, StringComparison.Ordinal);
        Assert.Contains(" 00002 n ", updatedText, StringComparison.Ordinal);
        Assert.Contains("[ 5 2 R ]", updatedText, StringComparison.Ordinal);
        Assert.Equal("Grace", Assert.Single(PdfInspector.Inspect(updated).FormFields).Value);
    }

    [Fact]
    public void IncrementalUpdater_PreservesTrailerReferenceGenerationsWhenAppendingFormFieldRevision() {
        byte[] pdf = BuildGeneratedFormPdfWithTrailerGenerations();

        byte[] updated = PdfIncrementalUpdater.UpdateFormFields(pdf, new Dictionary<string, string> {
            ["Name"] = "Grace"
        });

        string updatedText = PdfEncoding.Latin1GetString(updated);
        int appendedTrailer = updatedText.LastIndexOf("trailer", StringComparison.Ordinal);

        Assert.True(appendedTrailer >= PdfEncoding.Latin1GetString(pdf).Length);
        Assert.Contains("/Root 1 2 R", updatedText.Substring(appendedTrailer), StringComparison.Ordinal);
        Assert.Contains("/Info 8 3 R", updatedText.Substring(appendedTrailer), StringComparison.Ordinal);
        Assert.Equal("Grace", Assert.Single(PdfInspector.Inspect(updated).FormFields).Value);
    }

    [Fact]
    public void IncrementalUpdater_RejectsUnknownRadioButtonState() {
        byte[] pdf = PdfDocument.Create()
            .RadioButtonGroup("Payment.Method", new[] { "Card", "Cash", "Wire" }, value: "Card")
            .ToBytes();

        Assert.Throws<ArgumentException>(() => PdfIncrementalUpdater.UpdateFormFields(pdf, new Dictionary<string, string> {
            ["Payment.Method"] = "Crypto"
        }, new PdfIncrementalFormFieldUpdateOptions {
            GenerateAppearanceStreams = true
        }));
    }

    [Fact]
    public void PageEditor_SetsProductionBoundaryBoxes() {
        byte[] pdf = PdfPageGeometrySupport.BuildPageGeometryPdf();

        byte[] updated = PdfPageEditor.SetPageBox(pdf, "TrimBox", 12, 14, 222, 244);

        PdfPageGeometry geometry = PdfInspector.Inspect(updated).Pages[0].Geometry!;
        Assert.NotNull(geometry.TrimBox);
        Assert.Equal(12, geometry.TrimBox!.Left);
        Assert.Equal(14, geometry.TrimBox.Bottom);
        Assert.Equal(222, geometry.TrimBox.Right);
        Assert.Equal(244, geometry.TrimBox.Top);
    }

    [Fact]
    public void PageEditor_NamedAndFluentBoundaryBoxApisUseSharedEngine() {
        byte[] pdf = PdfPageGeometrySupport.BuildPageGeometryPdf();

        PdfDocument updated = PdfDocument.Open(pdf)
            .Pages.SetMediaBox(0, 0, 260, 280)
            .Pages.SetCropBox(5, 6, 250, 270)
            .Pages.SetBleedBox(7, 8, 248, 268)
            .Pages.SetTrimBox(9, 10, 246, 266)
            .Pages.SetArtBox(11, 12, 244, 264);
        PdfPageGeometry geometry = updated.Inspect().Pages[0].Geometry!;
        PdfPageGeometry typed = PdfInspector.Inspect(
            PdfPageEditor.SetPageBox(pdf, PdfPageBoundaryBox.CropBox, 13, 14, 230, 240)).Pages[0].Geometry!;

        Assert.Equal(260, geometry.MediaBox.Width);
        Assert.Equal(5, geometry.CropBox!.Left);
        Assert.Equal(7, geometry.BleedBox!.Left);
        Assert.Equal(9, geometry.TrimBox!.Left);
        Assert.Equal(11, geometry.ArtBox!.Left);
        Assert.Equal(13, typed.CropBox!.Left);
        Assert.Equal(240, typed.CropBox.Top);
    }

    [Fact]
    public void PageEditor_SetPageBoxPreservesSourceHeaderVersion() {
        byte[] pdf = BuildVersionedPageBoxPdf("2.0");

        byte[] updated = PdfPageEditor.SetPageBox(pdf, "TrimBox", 12, 14, 222, 244);
        PdfDocumentInfo info = PdfInspector.Inspect(updated);

        Assert.StartsWith("%PDF-2.0", PdfEncoding.Latin1GetString(updated), StringComparison.Ordinal);
        Assert.Equal("2.0", info.HeaderVersion);
        Assert.Equal("2.0", info.EffectiveVersion);
        Assert.Equal(12, info.Pages[0].Geometry!.TrimBox!.Left);
    }

    [Fact]
    public void PageEditor_ResizePagesTransformsAnnotationsAndNormalizesProductionBoxes() {
        byte[] pdf = BuildResizableAnnotatedPagePdf();

        byte[] resized = PdfPageEditor.ResizePages(pdf, new PdfPageResizeOptions(new PageSize(600, 600)) {
            Mode = PdfPageResizeMode.Stretch
        });

        PdfDocumentInfo info = PdfInspector.Inspect(resized);
        PdfPageGeometry geometry = info.Pages[0].Geometry!;
        PdfAnnotation annotation = Assert.Single(info.GetAnnotationsBySubtype("Link"));
        string raw = Encoding.ASCII.GetString(resized);

        Assert.Equal(600, geometry.MediaBox.Width);
        Assert.Equal(600, geometry.CropBox!.Width);
        Assert.Equal(600, geometry.TrimBox!.Width);
        Assert.Equal(600, geometry.BleedBox!.Width);
        Assert.Equal(600, geometry.ArtBox!.Width);
        Assert.Equal(120, annotation.X1);
        Assert.Equal(420, annotation.Y1);
        Assert.Equal(240, annotation.X2);
        Assert.Equal(540, annotation.Y2);
        Assert.Contains("/UserUnit 1", raw, StringComparison.Ordinal);
        Assert.Contains("/Rotate 0", raw, StringComparison.Ordinal);
        Assert.Contains("0 -6 6 0 -60 660 cm", raw, StringComparison.Ordinal);
        Assert.Contains("10 10 100 100 re\nW n", raw, StringComparison.Ordinal);
        Assert.Contains("/QuadPoints [ 420 540 420 420 360 540 360 420 ]", raw, StringComparison.Ordinal);
        Assert.Contains("/L [ 120 540 240 420 ]", raw, StringComparison.Ordinal);
        Assert.Contains("/Vertices [ 120 540 240 420 ]", raw, StringComparison.Ordinal);
        Assert.Contains("/InkList [ [ 120 540 240 420 ] ]", raw, StringComparison.Ordinal);

        Assert.NotNull(info.OpenAction);
        Assert.Equal(420, info.OpenAction!.DestinationLeft);
        Assert.Equal(540, info.OpenAction.DestinationTop);
        PdfNamedDestination namedDestination = Assert.Single(info.NamedDestinations, destination => destination.Name == "Target");
        Assert.Equal(420, namedDestination.DestinationLeft);
        Assert.Equal(540, namedDestination.DestinationTop);
        PdfOutlineItem outline = Assert.Single(info.Outlines);
        Assert.Equal(420, outline.DestinationLeft);
        Assert.Equal(540, outline.DestinationTop);
    }

    [Fact]
    public void PageEditor_ResizePagesTransformsDestinationsFromUnresizedPages() {
        byte[] pdf = BuildResizableTwoPageLinkPdf();

        byte[] resized = PdfPageEditor.ResizePages(pdf, new PdfPageResizeOptions(new PageSize(600, 600)) {
            Mode = PdfPageResizeMode.Stretch
        }, 1);

        PdfDocumentInfo info = PdfInspector.Inspect(resized);
        PdfLinkAnnotation link = Assert.Single(info.Pages[1].LinkAnnotations);

        Assert.Equal(60, link.DestinationLeft);
        Assert.Equal(420, link.DestinationTop);
        Assert.Equal(600, info.Pages[0].Geometry!.MediaBox.Width);
        Assert.Equal(300, info.Pages[1].Geometry!.MediaBox.Width);
    }

    [Fact]
    public void PageEditor_ResizePagesConvertsRotatedFitDestinationsToConcretePoints() {
        byte[] pdf = BuildResizableRotatedFitDestinationPdf();

        byte[] resized = PdfPageEditor.ResizePages(pdf, new PdfPageResizeOptions(new PageSize(600, 600)) {
            Mode = PdfPageResizeMode.Stretch
        });

        PdfDocumentInfo info = PdfInspector.Inspect(resized);

        Assert.NotNull(info.OpenAction);
        Assert.Equal(PdfOpenActionDestinationMode.Xyz, info.OpenAction!.DestinationMode);
        Assert.Equal(420, info.OpenAction.DestinationLeft);
        Assert.Equal(600, info.OpenAction.DestinationTop);
    }

    [Fact]
    public void PageEditor_ResizePagesConvertsRotatedPartialXyzDestinationsToConcretePoints() {
        byte[] pdf = BuildResizableRotatedPartialXyzDestinationPdf();

        byte[] resized = PdfPageEditor.ResizePages(pdf, new PdfPageResizeOptions(new PageSize(600, 600)) {
            Mode = PdfPageResizeMode.Stretch
        });

        PdfDocumentInfo info = PdfInspector.Inspect(resized);

        Assert.NotNull(info.OpenAction);
        Assert.Equal(PdfOpenActionDestinationMode.Xyz, info.OpenAction!.DestinationMode);
        Assert.Equal(420, info.OpenAction.DestinationLeft);
        Assert.Equal(600, info.OpenAction.DestinationTop);
    }

    [Fact]
    public void PageEditor_ResizePagesFillClipsToMarginBoxBeforeScaling() {
        byte[] pdf = BuildResizableTallPagePdf();

        byte[] resized = PdfPageEditor.ResizePages(pdf, new PdfPageResizeOptions(new PageSize(600, 600)) {
            Mode = PdfPageResizeMode.Fill,
            Margin = 50
        });
        string raw = PdfEncoding.Latin1GetString(resized);

        int targetClip = raw.IndexOf("50 50 500 500 re\nW n\n", StringComparison.Ordinal);
        int transform = raw.IndexOf("5 0 0 5 50 -450 cm\n", StringComparison.Ordinal);
        int sourceClip = raw.IndexOf("0 0 100 300 re\nW n\n", StringComparison.Ordinal);
        Assert.True(targetClip >= 0, "Expected resized content stream to clip to the target margin box.");
        Assert.True(transform > targetClip, "Expected target-space margin clipping before the resize transform.");
        Assert.True(sourceClip > transform, "Expected source-space clipping after the resize transform.");
    }

    [Fact]
    public void PageEditor_ResizePagesFillClipsAnnotationRectanglesToMarginBox() {
        byte[] pdf = BuildResizableTallPageWithCroppedAnnotationsPdf();

        byte[] resized = PdfPageEditor.ResizePages(pdf, new PdfPageResizeOptions(new PageSize(600, 600)) {
            Mode = PdfPageResizeMode.Fill,
            Margin = 50
        });
        PdfDocumentInfo info = PdfInspector.Inspect(resized);
        string raw = PdfEncoding.Latin1GetString(resized);

        PdfAnnotation annotation = Assert.Single(info.GetAnnotationsBySubtype("Link"));
        Assert.Equal(100, annotation.X1);
        Assert.Equal(50, annotation.Y1);
        Assert.Equal(200, annotation.X2);
        Assert.Equal(100, annotation.Y2);
        Assert.Contains("/Rect [ 100 50 200 100 ]", raw, StringComparison.Ordinal);
        Assert.DoesNotContain("https://example.com/clipped-away", raw, StringComparison.Ordinal);
    }

    [Fact]
    public void PageEditor_ResizePagesTransformsSharedIndirectDestinationsOnce() {
        byte[] pdf = BuildResizableSharedDestinationPdf();

        byte[] resized = PdfPageEditor.ResizePages(pdf, new PdfPageResizeOptions(new PageSize(600, 600)) {
            Mode = PdfPageResizeMode.Stretch
        });

        PdfDocumentInfo info = PdfInspector.Inspect(resized);

        Assert.Equal(2, info.Pages[0].LinkAnnotations.Count);
        Assert.All(info.Pages[0].LinkAnnotations, link => {
            Assert.Equal(60, link.DestinationLeft);
            Assert.Equal(420, link.DestinationTop);
        });
        Assert.NotNull(info.OpenAction);
        Assert.Equal(60, info.OpenAction!.DestinationLeft);
        Assert.Equal(420, info.OpenAction.DestinationTop);
    }

    [Fact]
    public void PageEditor_ResizePagesTransformsIndirectAnnotationGeometryArrays() {
        byte[] pdf = BuildResizableIndirectAnnotationGeometryPdf();

        byte[] resized = PdfPageEditor.ResizePages(pdf, new PdfPageResizeOptions(new PageSize(600, 600)) {
            Mode = PdfPageResizeMode.Stretch
        });
        string raw = PdfEncoding.Latin1GetString(resized);

        Assert.Contains("/Rect [ 60 120 180 240 ]", raw, StringComparison.Ordinal);
        Assert.Contains("/QuadPoints [ 60 420 180 420 60 360 180 360 ]", raw, StringComparison.Ordinal);
        Assert.Contains("/InkList [ [ 60 420 180 360 ] ]", raw, StringComparison.Ordinal);
        Assert.DoesNotContain("/Rect 7 0 R", raw, StringComparison.Ordinal);
        Assert.DoesNotContain("/QuadPoints 8 0 R", raw, StringComparison.Ordinal);
        Assert.DoesNotContain("/InkList 9 0 R", raw, StringComparison.Ordinal);
    }

    [Fact]
    public void PageEditor_ResizePagesRemapsPopupAnnotationReferencesToTransformedClones() {
        byte[] pdf = BuildResizablePopupAnnotationPdf();

        byte[] resized = PdfPageEditor.ResizePages(pdf, new PdfPageResizeOptions(new PageSize(600, 600)) {
            Mode = PdfPageResizeMode.Stretch
        });
        string raw = PdfEncoding.Latin1GetString(resized);

        Assert.DoesNotContain("/Rect [ 50 50 150 120 ]", raw, StringComparison.Ordinal);
        Assert.Contains("/Rect [ 240 240 600 600 ]", raw, StringComparison.Ordinal);
        Assert.Contains("/Popup", raw, StringComparison.Ordinal);
        Assert.Contains("/Parent", raw, StringComparison.Ordinal);
    }

    [Fact]
    public void PageEditor_ResizePagesCollectsClonedAnnotationActionDependencies() {
        byte[] pdf = BuildResizableAnnotationActionPdf();

        byte[] resized = PdfPageEditor.ResizePages(pdf, new PdfPageResizeOptions(new PageSize(600, 600)) {
            Mode = PdfPageResizeMode.Stretch
        });
        string raw = PdfEncoding.Latin1GetString(resized);

        Assert.Contains("/A ", raw, StringComparison.Ordinal);
        Assert.Contains("/URI (https://example.com/review)", raw, StringComparison.Ordinal);
    }

    [Fact]
    public void SecurityInfo_ReportsObjectStreamRewriteReadiness() {
        byte[] pdf = Encoding.ASCII.GetBytes(string.Join("\n", new[] {
            "%PDF-1.7",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 0 /Kids [] >>",
            "endobj",
            "3 0 obj",
            "<< /Type /ObjStm /N 0 /First 0 /Length 0 >>",
            "stream",
            "",
            "endstream",
            "endobj",
            "trailer",
            "<< /Root 1 0 R /Size 4 >>",
            "startxref",
            "123",
            "%%EOF"
        }));

        PdfDocumentSecurityInfo security = PdfInspector.Probe(pdf).Security;

        Assert.True(security.HasObjectStreams);
        Assert.True(security.BlocksOfficeIMOFullRewriteMutation);
    }

    [Fact]
    public void AnnotationEditor_UpdatesAndRemovesAnnotations() {
        byte[] pdf = BuildAnnotationEditPdf();

        PdfAnnotationEditResult updated = PdfAnnotationEditor.UpdateAnnotation(pdf, 6, new PdfAnnotationUpdateOptions {
            Contents = "Updated note",
            Title = "Lead reviewer",
            Name = "Note-2",
            Flags = 4,
            Color = new[] { 0D, 0.5D, 1D },
            RemoveActions = true
        });

        Assert.True(updated.Applied);
        PdfAnnotation annotation = Assert.Single(PdfInspector.Inspect(updated.Bytes).GetAnnotationsBySubtype("Text"));
        Assert.Equal("Updated note", annotation.Contents);
        Assert.Equal("Lead reviewer", annotation.Title);
        Assert.Equal("Note-2", annotation.Name);
        Assert.False(annotation.HasAction);
        Assert.False(annotation.HasAdditionalActions);
        Assert.False(annotation.HasChainedActions);
        Assert.Equal(4, annotation.Flags);

        PdfAnnotationEditResult removed = PdfAnnotationEditor.RemoveAnnotations(updated.Bytes, new PdfAnnotationRemovalOptions {
            Subtype = "Text"
        });

        Assert.True(removed.Applied);
        PdfDocumentInfo info = PdfInspector.Inspect(removed.Bytes);
        Assert.Empty(info.GetAnnotationsBySubtype("Text"));
        Assert.Equal(0, info.AnnotationCount);
    }

    [Fact]
    public void AnnotationEditor_InvalidatesAppearanceWhenUpdatingVisibleAnnotationText() {
        byte[] pdf = BuildAnnotationEditPdf();

        PdfAnnotationEditResult updated = PdfAnnotationEditor.UpdateAnnotation(pdf, 6, new PdfAnnotationUpdateOptions {
            Contents = "Updated note"
        });

        string annotationObject = ExtractObjectBlock(PdfEncoding.Latin1GetString(updated.Bytes), 6);
        PdfAnnotation annotation = Assert.Single(PdfInspector.Inspect(updated.Bytes).GetAnnotationsBySubtype("Text"));

        Assert.True(updated.Applied);
        Assert.Equal("Updated note", annotation.Contents);
        Assert.DoesNotContain("/AP", annotationObject, StringComparison.Ordinal);
    }

    [Fact]
    public void AnnotationEditor_PrunesOrphanedAppearanceAndActionObjectsWhenUpdating() {
        byte[] pdf = BuildAnnotationEditPdf();

        PdfAnnotationEditResult updated = PdfAnnotationEditor.UpdateAnnotation(pdf, 6, new PdfAnnotationUpdateOptions {
            Contents = "Updated note",
            RemoveActions = true
        });
        string raw = PdfEncoding.Latin1GetString(updated.Bytes);

        Assert.True(updated.Applied);
        Assert.DoesNotContain("Old note appearance", raw, StringComparison.Ordinal);
        Assert.DoesNotContain("old-action", raw, StringComparison.Ordinal);
    }

    [Fact]
    public void AnnotationEditor_RejectsNonAnnotationDictionaryWithSubtype() {
        byte[] pdf = BuildAnnotationEditPdf();

        Assert.Throws<ArgumentException>(() => PdfAnnotationEditor.UpdateAnnotation(pdf, 4, new PdfAnnotationUpdateOptions {
            Contents = "Not an annotation"
        }));
    }

    [Fact]
    public void AnnotationEditor_ClearsParentPopupReferencesWhenRemovingPopupAnnotations() {
        byte[] pdf = BuildAnnotationWithPopupPdf();

        PdfAnnotationEditResult removed = PdfAnnotationEditor.RemoveAnnotations(pdf, new PdfAnnotationRemovalOptions {
            Subtype = "Popup"
        });
        string raw = PdfEncoding.Latin1GetString(removed.Bytes);

        Assert.True(removed.Applied);
        Assert.DoesNotContain("/Subtype /Popup", raw, StringComparison.Ordinal);
        Assert.DoesNotContain("/Popup", raw, StringComparison.Ordinal);
        Assert.Single(PdfInspector.Inspect(removed.Bytes).GetAnnotationsBySubtype("Text"));
    }

    [Fact]
    public void PdfOptions_RejectsEncryptionWithPdfABackedGroundwork() {
        Assert.Throws<ArgumentException>(() => PdfDocument.Create(
                new PdfOptions()
                    .ConfigurePdfAGroundwork(PdfComplianceProfile.PdfA4)
                    .SetEncryption("open"))
            .Paragraph(p => p.Text("PDF/A and encryption should not mix."))
            .ToBytes());

        Assert.Throws<ArgumentException>(() => PdfDocument.Create(
                new PdfOptions()
                    .ConfigureFacturXGroundwork(Encoding.UTF8.GetBytes("<rsm:CrossIndustryInvoice xmlns:rsm=\"urn:invoice\"/>"))
                    .SetEncryption("open"))
            .Paragraph(p => p.Text("Factur-X and encryption should not mix."))
            .ToBytes());
    }

    [Fact]
    public void ExternalValidationResult_FromExitCodeImportsProcessOutcome() {
        PdfExternalValidationResult passed = PdfExternalValidationResult.FromExitCode(
            PdfExternalValidatorKind.VeraPdf,
            0,
            "veraPDF",
            "passed",
            profile: "PDF/A-3b",
            executablePath: "verapdf",
            arguments: "--format text file.pdf");

        PdfExternalValidationResult failed = PdfExternalValidationResult.FromExitCode(
            PdfExternalValidatorKind.PdfUaValidator,
            2,
            "pdfua",
            "failed",
            profile: "PDF/UA-1");

        Assert.Equal(PdfExternalValidationStatus.Passed, passed.Status);
        Assert.Equal(0, passed.ExitCode);
        Assert.Equal("verapdf", passed.ExecutablePath);
        Assert.Equal("--format text file.pdf", passed.Arguments);
        Assert.Equal(PdfExternalValidationStatus.Failed, failed.Status);
        Assert.Equal(2, failed.ExitCode);
    }

    private static byte[] BuildCoveragePdf() {
        string longText = new string('A', 512);
        byte[] contentBytes = Encoding.ASCII.GetBytes("BT\n/F1 12 Tf\n72 720 Td\n(" + longText + ") Tj\nET\n");
        byte[] fontBytes = { 1, 2, 3, 4 };
        var objects = new List<string> {
            "<< /Type /Catalog /Pages 2 0 R /AcroForm 8 0 R /MarkInfo << /Marked true >> /StructTreeRoot 10 0 R >>",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 300 300] /CropBox [0 0 290 290] /BleedBox [5 5 295 295] /TrimBox [10 10 290 290] /ArtBox [20 20 280 280] /Resources << /Font << /F1 4 0 R /F2 14 0 R >> >> /Annots [6 0 R 7 0 R] /Contents 5 0 R >>",
            "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>",
            BuildStream(contentBytes),
            "<< /Type /Annot /Subtype /Text /Rect [20 20 40 40] /Contents (Review note) /F 132 /NM (Note-1) /T (Reviewer) /M (D:20260622090000Z) /C [1 0 0] /AA << /E << /S /Launch /F (tool.exe) >> >> >>",
            "<< /Type /Annot /Subtype /Widget /FT /Tx /T (Name) /V (Ada) /Ff 3 /Rect [50 50 180 70] /F 4 >>",
            "<< /Fields [7 0 R] /SigFlags 2 >>",
            "<< /Type /FontDescriptor /FontName /EmbeddedSans /FontFile2 15 0 R >>",
            "<< /Type /StructTreeRoot /K [11 0 R] /ParentTree 13 0 R /ParentTreeNextKey 1 /RoleMap << /Custom /P >> >>",
            "<< /Type /StructElem /S /Document /P 10 0 R /K [12 0 R] /Lang (en-US) >>",
            "<< /Type /StructElem /S /Figure /P 11 0 R /K 0 >>",
            "<< /Nums [0 12 0 R] >>",
            "<< /Type /Font /Subtype /TrueType /BaseFont /EmbeddedSans /FontDescriptor 9 0 R >>",
            BuildStream(fontBytes)
        };

        return Encoding.ASCII.GetBytes(BuildPdf(objects));
    }

    private static byte[] BuildMultilineFormPdf() {
        var objects = new List<string> {
            "<< /Type /Catalog /Pages 2 0 R /AcroForm 6 0 R >>",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 300 300] /Resources << /Font << /F1 4 0 R >> >> /Annots [5 0 R] /Contents 7 0 R >>",
            "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>",
            "<< /Type /Annot /Subtype /Widget /FT /Tx /T (Notes) /V () /Ff 4096 /Rect [50 50 120 122] /F 4 >>",
            "<< /Fields [5 0 R] /SigFlags 2 >>",
            BuildStream(Encoding.ASCII.GetBytes("BT /F1 12 Tf 72 720 Td (Multiline form) Tj ET"))
        };

        return Encoding.ASCII.GetBytes(BuildPdf(objects));
    }

    private static byte[] BuildDefaultAppearanceFormPdf(string defaultAppearance, string widgetEntries = "") {
        var objects = new List<string> {
            "<< /Type /Catalog /Pages 2 0 R /AcroForm 6 0 R >>",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 300 300] /Resources << /Font << /F1 4 0 R >> >> /Annots [5 0 R] /Contents 7 0 R >>",
            "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>",
            "<< /Type /Annot /Subtype /Widget /FT /Tx /T (Name) /V (Ada) /DA (" + defaultAppearance + ") /Rect [50 50 180 70] /F 4" + widgetEntries + " >>",
            "<< /Fields [5 0 R] /SigFlags 2 >>",
            BuildStream(Encoding.ASCII.GetBytes("BT /F1 12 Tf 72 720 Td (Default appearance form) Tj ET"))
        };

        return Encoding.ASCII.GetBytes(BuildPdf(objects));
    }

    private static byte[] BuildAcroFormDefaultAppearanceFormPdf(string defaultAppearance, bool includeFontResource = false, bool includeWidgetPageReference = false, bool includeWidgetAppearanceResource = false) {
        string acroForm = includeFontResource
            ? "<< /Fields [5 0 R] /SigFlags 2 /DA (" + defaultAppearance + ") /DR << /Font << /F1 4 0 R >> >> >>"
            : "<< /Fields [5 0 R] /SigFlags 2 /DA (" + defaultAppearance + ") >>";
        string widget = "<< /Type /Annot /Subtype /Widget /FT /Tx /T (Name) /V (Ada) /Rect [50 50 180 70] /F 4" +
            (includeWidgetPageReference ? " /P 3 0 R" : string.Empty) +
            (includeWidgetAppearanceResource ? " /AP << /N 8 0 R >>" : string.Empty) +
            " >>";
        var objects = new List<string> {
            "<< /Type /Catalog /Pages 2 0 R /AcroForm 6 0 R >>",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 300 300] /Resources << /Font << /F1 4 0 R >> >> /Annots [5 0 R] /Contents 7 0 R >>",
            "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>",
            widget,
            acroForm,
            BuildStream(Encoding.ASCII.GetBytes("BT /F1 12 Tf 72 720 Td (Default appearance form) Tj ET"))
        };
        if (includeWidgetAppearanceResource) {
            objects.Add(BuildStream(
                Encoding.ASCII.GetBytes("BT /FAP 8.5 Tf 0 0 Td (Old appearance) Tj ET"),
                "/Type /XObject /Subtype /Form /BBox [0 0 130 20] /Resources << /Font << /FAP 4 0 R >> >>"));
        }

        return Encoding.ASCII.GetBytes(BuildPdf(objects));
    }

    private static byte[] BuildAnnotationEditPdf() {
        byte[] contentBytes = Encoding.ASCII.GetBytes("BT\n/F1 12 Tf\n72 720 Td\n(Annotation editing) Tj\nET\n");
        byte[] appearanceBytes = Encoding.ASCII.GetBytes("BT /F1 12 Tf 0 0 Td (Old note appearance) Tj ET");
        var objects = new List<string> {
            "<< /Type /Catalog /Pages 2 0 R >>",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 300 300] /Resources << /Font << /F1 4 0 R >> >> /Annots [6 0 R] /Contents 5 0 R >>",
            "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>",
            BuildStream(contentBytes),
            "<< /Type /Annot /Subtype /Text /Rect [20 20 40 40] /Contents (Review note) /F 132 /NM (Note-1) /T (Reviewer) /M (D:20260622090000Z) /C [1 0 0] /AP << /N 7 0 R >> /A 8 0 R >>",
            BuildStream(appearanceBytes),
            "<< /S /URI /URI (https://example.com/old-action) >>"
        };

        return Encoding.ASCII.GetBytes(BuildPdf(objects));
    }

    private static byte[] BuildAnnotationWithPopupPdf() {
        byte[] contentBytes = Encoding.ASCII.GetBytes("BT\n/F1 12 Tf\n72 720 Td\n(Annotation popup cleanup) Tj\nET\n");
        var objects = new List<string> {
            "<< /Type /Catalog /Pages 2 0 R >>",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 300 300] /Resources << /Font << /F1 4 0 R >> >> /Annots [6 0 R 7 0 R] /Contents 5 0 R >>",
            "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>",
            BuildStream(contentBytes),
            "<< /Type /Annot /Subtype /Text /Rect [20 20 40 40] /Contents (Review note) /Popup 7 0 R >>",
            "<< /Type /Annot /Subtype /Popup /Rect [50 50 150 120] /Parent 6 0 R >>"
        };

        return Encoding.ASCII.GetBytes(BuildPdf(objects));
    }

    private static string ExtractObjectBlock(string pdf, int objectNumber) {
        string marker = objectNumber.ToString(CultureInfo.InvariantCulture) + " 0 obj";
        int start = pdf.IndexOf(marker, StringComparison.Ordinal);
        Assert.True(start >= 0);
        int end = pdf.IndexOf("endobj", start, StringComparison.Ordinal);
        Assert.True(end > start);
        return pdf.Substring(start, end - start);
    }

    internal static byte[] BuildDocMdpFormPdf(int permissionLevel, string? lockDictionary = null, string fieldName = "Name") {
        string lockEntry = string.IsNullOrWhiteSpace(lockDictionary) ? string.Empty : " /Lock " + lockDictionary;
        var objects = new List<string> {
            "<< /Type /Catalog /Pages 2 0 R /AcroForm 8 0 R /Perms << /DocMDP 7 0 R >> >>",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 300 300] /Resources << /Font << /F1 4 0 R >> >> /Annots [5 0 R 6 0 R] /Contents 9 0 R >>",
            "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>",
            "<< /Type /Annot /Subtype /Widget /FT /Tx /T (" + fieldName + ") /V (Ada) /Rect [50 50 180 70] /F 4 >>",
            "<< /FT /Sig /T (Approval) /V 7 0 R /Subtype /Widget /Rect [10 10 120 40]" + lockEntry + " >>",
            "<< /Type /Sig /Filter /Adobe.PPKLite /SubFilter /adbe.pkcs7.detached /Name (Alice) /ByteRange [0 10 20 30] /Contents <001122> /Reference [<< /TransformMethod /DocMDP /TransformParams << /Type /TransformParams /V /1.2 /P " + permissionLevel.ToString(CultureInfo.InvariantCulture) + " >> >>] >>",
            "<< /Fields [5 0 R 6 0 R] /SigFlags 3 >>",
            BuildStream(Encoding.ASCII.GetBytes("BT /F1 12 Tf 72 720 Td (Signed form) Tj ET"))
        };

        return Encoding.ASCII.GetBytes(BuildPdf(objects));
    }

    private static byte[] BuildGeneratedFormPdf() {
        var entries = new List<(int ObjectNumber, int Generation, string Body)> {
            (1, 0, "<< /Type /Catalog /Pages 2 0 R /AcroForm 6 0 R >>"),
            (2, 0, "<< /Type /Pages /Count 1 /Kids [3 0 R] >>"),
            (3, 0, "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 300 300] /Resources << /Font << /F1 4 0 R >> >> /Annots [5 2 R] /Contents 7 0 R >>"),
            (4, 0, "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>"),
            (5, 2, "<< /Type /Annot /Subtype /Widget /FT /Tx /T (Name) /V (Ada) /Rect [50 50 180 70] /F 4 >>"),
            (6, 0, "<< /Fields [5 2 R] /SigFlags 2 >>"),
            (7, 0, BuildStream(Encoding.ASCII.GetBytes("BT /F1 12 Tf 72 720 Td (Generated form) Tj ET")))
        };

        var builder = new StringBuilder();
        builder.AppendLine("%PDF-1.7");
        foreach ((int objectNumber, int generation, string body) in entries) {
            builder.Append(objectNumber.ToString(CultureInfo.InvariantCulture)).Append(' ')
                .Append(generation.ToString(CultureInfo.InvariantCulture)).AppendLine(" obj");
            builder.AppendLine(body);
            builder.AppendLine("endobj");
        }

        builder.AppendLine("trailer");
        builder.AppendLine("<< /Root 1 0 R /Size 8 >>");
        builder.AppendLine("startxref");
        builder.AppendLine("123");
        builder.AppendLine("%%EOF");
        return Encoding.ASCII.GetBytes(builder.ToString());
    }

    private static byte[] BuildGeneratedFormPdfWithTrailerGenerations() {
        var entries = new List<(int ObjectNumber, int Generation, string Body)> {
            (1, 2, "<< /Type /Catalog /Pages 2 0 R /AcroForm 6 0 R >>"),
            (2, 0, "<< /Type /Pages /Count 1 /Kids [3 0 R] >>"),
            (3, 0, "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 300 300] /Resources << /Font << /F1 4 0 R >> >> /Annots [5 0 R] /Contents 7 0 R >>"),
            (4, 0, "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>"),
            (5, 0, "<< /Type /Annot /Subtype /Widget /FT /Tx /T (Name) /V (Ada) /Rect [50 50 180 70] /F 4 >>"),
            (6, 0, "<< /Fields [5 0 R] /SigFlags 2 >>"),
            (7, 0, BuildStream(Encoding.ASCII.GetBytes("BT /F1 12 Tf 72 720 Td (Generated form) Tj ET"))),
            (8, 3, "<< /Producer (OfficeIMO) >>")
        };

        var builder = new StringBuilder();
        builder.AppendLine("%PDF-1.7");
        foreach ((int objectNumber, int generation, string body) in entries) {
            builder.Append(objectNumber.ToString(CultureInfo.InvariantCulture)).Append(' ')
                .Append(generation.ToString(CultureInfo.InvariantCulture)).AppendLine(" obj");
            builder.AppendLine(body);
            builder.AppendLine("endobj");
        }

        builder.AppendLine("trailer");
        builder.AppendLine("<< /Root 1 2 R /Info 8 3 R /Size 9 >>");
        builder.AppendLine("startxref");
        builder.AppendLine("123");
        builder.AppendLine("%%EOF");
        return Encoding.ASCII.GetBytes(builder.ToString());
    }

    private static byte[] BuildVersionedPageBoxPdf(string version) {
        var objects = new List<string> {
            "<< /Type /Catalog /Pages 2 0 R >>",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 300 300] /Resources << /Font << /F1 4 0 R >> >> /Contents 5 0 R /UserUnit 2 >>",
            "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>",
            BuildStream(Encoding.ASCII.GetBytes("BT /F1 12 Tf 72 720 Td (Versioned page box) Tj ET"))
        };

        return Encoding.ASCII.GetBytes(BuildPdf(objects).Replace("%PDF-1.7", "%PDF-" + version));
    }

    private static byte[] BuildResizableAnnotatedPagePdf() {
        var objects = new List<string> {
            "<< /Type /Catalog /Pages 2 0 R /OpenAction [3 0 R /XYZ 20 80 1] /Dests << /Target [3 0 R /XYZ 20 80 1] >> /Outlines 7 0 R >>",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 300 300] /CropBox [10 10 110 110] /BleedBox [5 5 295 295] /TrimBox [10 10 290 290] /ArtBox [20 20 280 280] /UserUnit 2 /Rotate 90 /Resources << /Font << /F1 4 0 R >> >> /Annots [6 0 R] /Contents 5 0 R >>",
            "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>",
            BuildStream(Encoding.ASCII.GetBytes("BT /F1 12 Tf 20 20 Td (Resize source) Tj ET")),
            "<< /Type /Annot /Subtype /Link /Rect [20 30 40 50] /QuadPoints [20 80 40 80 20 70 40 70] /L [20 30 40 50] /Vertices [20 30 40 50] /InkList [[20 30 40 50]] /Dest [3 0 R /XYZ 20 80 1] >>",
            "<< /Type /Outlines /First 8 0 R /Last 8 0 R /Count 1 >>",
            "<< /Title (Target) /Parent 7 0 R /Dest [3 0 R /XYZ 20 80 1] >>"
        };

        return Encoding.ASCII.GetBytes(BuildPdf(objects));
    }

    private static byte[] BuildResizableTwoPageLinkPdf() {
        var objects = new List<string> {
            "<< /Type /Catalog /Pages 2 0 R >>",
            "<< /Type /Pages /Count 2 /Kids [3 0 R 7 0 R] >>",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 300 300] /CropBox [10 10 110 110] /Resources << /Font << /F1 4 0 R >> >> /Contents 5 0 R >>",
            "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>",
            BuildStream(Encoding.ASCII.GetBytes("BT /F1 12 Tf 20 20 Td (Resize target) Tj ET")),
            BuildStream(Encoding.ASCII.GetBytes("BT /F1 12 Tf 20 20 Td (Link source) Tj ET")),
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 300 300] /Resources << /Font << /F1 4 0 R >> >> /Annots [8 0 R] /Contents 6 0 R >>",
            "<< /Type /Annot /Subtype /Link /Rect [20 30 40 50] /Dest [3 0 R /XYZ 20 80 1] >>"
        };

        return Encoding.ASCII.GetBytes(BuildPdf(objects));
    }

    private static byte[] BuildResizableRotatedFitDestinationPdf() {
        var objects = new List<string> {
            "<< /Type /Catalog /Pages 2 0 R /OpenAction [3 0 R /FitH 80] >>",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 300 300] /CropBox [10 10 110 110] /Rotate 90 /Resources << /Font << /F1 4 0 R >> >> /Contents 5 0 R >>",
            "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>",
            BuildStream(Encoding.ASCII.GetBytes("BT /F1 12 Tf 20 20 Td (Fit target) Tj ET"))
        };

        return Encoding.ASCII.GetBytes(BuildPdf(objects));
    }

    private static byte[] BuildResizableRotatedPartialXyzDestinationPdf() {
        var objects = new List<string> {
            "<< /Type /Catalog /Pages 2 0 R /OpenAction [3 0 R /XYZ null 80 1] >>",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 300 300] /CropBox [10 10 110 110] /Rotate 90 /Resources << /Font << /F1 4 0 R >> >> /Contents 5 0 R >>",
            "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>",
            BuildStream(Encoding.ASCII.GetBytes("BT /F1 12 Tf 20 20 Td (Partial XYZ target) Tj ET"))
        };

        return Encoding.ASCII.GetBytes(BuildPdf(objects));
    }

    private static byte[] BuildResizableTallPagePdf() {
        var objects = new List<string> {
            "<< /Type /Catalog /Pages 2 0 R >>",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 100 300] /Resources << /Font << /F1 4 0 R >> >> /Contents 5 0 R >>",
            "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>",
            BuildStream(Encoding.ASCII.GetBytes("BT /F1 12 Tf 10 280 Td (Fill should crop to margins) Tj ET"))
        };

        return Encoding.ASCII.GetBytes(BuildPdf(objects));
    }

    private static byte[] BuildResizableTallPageWithCroppedAnnotationsPdf() {
        var objects = new List<string> {
            "<< /Type /Catalog /Pages 2 0 R >>",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 100 300] /Resources << /Font << /F1 4 0 R >> >> /Annots [6 0 R 7 0 R] /Contents 5 0 R >>",
            "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>",
            BuildStream(Encoding.ASCII.GetBytes("BT /F1 12 Tf 10 280 Td (Fill annotations should crop to margins) Tj ET")),
            "<< /Type /Annot /Subtype /Link /Rect [10 90 30 110] /A << /S /URI /URI (https://example.com/kept) >> >>",
            "<< /Type /Annot /Subtype /Link /Rect [10 10 30 20] /A << /S /URI /URI (https://example.com/clipped-away) >> >>"
        };

        return Encoding.ASCII.GetBytes(BuildPdf(objects));
    }

    private static byte[] BuildResizableSharedDestinationPdf() {
        var objects = new List<string> {
            "<< /Type /Catalog /Pages 2 0 R /OpenAction 8 0 R >>",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 300 300] /CropBox [10 10 110 110] /Resources << /Font << /F1 4 0 R >> >> /Annots [6 0 R 7 0 R] /Contents 5 0 R >>",
            "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>",
            BuildStream(Encoding.ASCII.GetBytes("BT /F1 12 Tf 20 20 Td (Shared target) Tj ET")),
            "<< /Type /Annot /Subtype /Link /Rect [20 30 40 50] /Dest 8 0 R >>",
            "<< /Type /Annot /Subtype /Link /Rect [60 30 80 50] /Dest 8 0 R >>",
            "[3 0 R /XYZ 20 80 1]"
        };

        return Encoding.ASCII.GetBytes(BuildPdf(objects));
    }

    private static byte[] BuildResizableIndirectAnnotationGeometryPdf() {
        var objects = new List<string> {
            "<< /Type /Catalog /Pages 2 0 R >>",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 300 300] /CropBox [10 10 110 110] /Resources << /Font << /F1 4 0 R >> >> /Annots [6 0 R] /Contents 5 0 R >>",
            "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>",
            BuildStream(Encoding.ASCII.GetBytes("BT /F1 12 Tf 20 20 Td (Indirect geometry) Tj ET")),
            "<< /Type /Annot /Subtype /Link /Rect 7 0 R /QuadPoints 8 0 R /InkList 9 0 R /Dest [3 0 R /XYZ 20 80 1] >>",
            "[20 30 40 50]",
            "[20 80 40 80 20 70 40 70]",
            "[[20 80 40 70]]"
        };

        return Encoding.ASCII.GetBytes(BuildPdf(objects));
    }

    private static byte[] BuildResizablePopupAnnotationPdf() {
        var objects = new List<string> {
            "<< /Type /Catalog /Pages 2 0 R >>",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 300 300] /CropBox [10 10 110 110] /Resources << /Font << /F1 4 0 R >> >> /Annots [6 0 R 7 0 R] /Contents 5 0 R >>",
            "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>",
            BuildStream(Encoding.ASCII.GetBytes("BT /F1 12 Tf 20 20 Td (Popup geometry) Tj ET")),
            "<< /Type /Annot /Subtype /Text /Rect [20 30 40 50] /Popup 7 0 R /Contents (Note) >>",
            "<< /Type /Annot /Subtype /Popup /Rect [50 50 150 120] /Parent 6 0 R >>"
        };

        return Encoding.ASCII.GetBytes(BuildPdf(objects));
    }

    private static byte[] BuildResizableAnnotationActionPdf() {
        var objects = new List<string> {
            "<< /Type /Catalog /Pages 2 0 R >>",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 300 300] /CropBox [10 10 110 110] /Resources << /Font << /F1 4 0 R >> >> /Annots [6 0 R] /Contents 5 0 R >>",
            "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>",
            BuildStream(Encoding.ASCII.GetBytes("BT /F1 12 Tf 20 20 Td (Action dependency) Tj ET")),
            "<< /Type /Annot /Subtype /Link /Rect [20 30 40 50] /A 7 0 R >>",
            "<< /S /URI /URI (https://example.com/review) >>"
        };

        return Encoding.ASCII.GetBytes(BuildPdf(objects));
    }

    private static string BuildStream(byte[] data) =>
        "<< /Length " + data.Length.ToString(CultureInfo.InvariantCulture) + " >>\nstream\n" +
        Encoding.ASCII.GetString(data) +
        "\nendstream";

    private static string BuildStream(byte[] data, string dictionaryEntries) =>
        "<< " + dictionaryEntries + " /Length " + data.Length.ToString(CultureInfo.InvariantCulture) + " >>\nstream\n" +
        Encoding.ASCII.GetString(data) +
        "\nendstream";

    private static string BuildPdf(IReadOnlyList<string> objects) {
        var builder = new StringBuilder();
        builder.AppendLine("%PDF-1.7");
        for (int i = 0; i < objects.Count; i++) {
            builder.Append((i + 1).ToString(CultureInfo.InvariantCulture)).AppendLine(" 0 obj");
            builder.AppendLine(objects[i]);
            builder.AppendLine("endobj");
        }

        builder.AppendLine("trailer");
        builder.Append("<< /Root 1 0 R /Size ").Append(objects.Count + 1).AppendLine(" >>");
        builder.AppendLine("startxref");
        builder.AppendLine("123");
        builder.AppendLine("%%EOF");
        return builder.ToString();
    }
}

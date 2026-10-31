using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfAllSeverityBatch21Tests {
    [Fact]
    public void UnprofiledExternalValidationDoesNotSatisfyProfiledProof() {
        PdfCreateOptions options = new PdfCreateOptions().ConfigurePdfAGroundwork(PdfComplianceProfile.PdfA3B);
        byte[] artifact = PdfDocument.Create(options).ToBytes();
        PdfExternalValidationResult validation = PdfExternalValidationResult.PassedForArtifact(
            PdfExternalValidatorKind.VeraPdf,
            "veraPDF",
            "1.30.2",
            "Generic validation passed.",
            artifact);

        PdfComplianceProofReport proof = PdfComplianceAnalyzer.AssessProof(
            PdfComplianceProfile.PdfA3B,
            options,
            artifact,
            new[] { validation },
            generatedStandardFonts: Array.Empty<PdfStandardFont>());

        Assert.False(proof.HasRequiredExternalValidation);
        Assert.False(proof.CanClaimConformance);
        Assert.Contains(PdfExternalValidatorKind.VeraPdf, proof.MissingExternalValidators);
    }

    [Fact]
    public void LegacyTwoArgumentFormCallsRemainSourceCompatibleWithoutMakingNullAmbiguous() {
        PdfDocumentForms forms = PdfDocument.Load(PdfDocument.Create().ToBytes()).Forms;
        var strings = new Dictionary<string, string>();
        var typed = new Dictionary<string, PdfFormFieldValue>();
        var formOptions = new PdfFormFillerOptions();

        _ = forms.FillResult(strings, formOptions);
        _ = forms.FillResult(typed, formOptions);
        _ = forms.FlattenResult(formOptions);
        _ = forms.FillAndFlattenResult(strings, formOptions);
        _ = forms.FillAndFlattenResult(typed, formOptions);

        _ = forms.FillResult(strings, null);
        _ = forms.FillResult(typed, null);
        _ = forms.FlattenResult(null);
        _ = forms.FillAndFlattenResult(strings, null);
        _ = forms.FillAndFlattenResult(typed, null);
    }

    [Fact]
    public void TabAlignmentPreservesPersistedRightValue() {
        Assert.Equal(0, (int)PdfTabAlignment.Left);
        Assert.Equal(1, (int)PdfTabAlignment.Right);
        Assert.Equal(2, (int)PdfTabAlignment.Center);
        Assert.Equal(3, (int)PdfTabAlignment.DecimalSeparator);
    }

    [Fact]
    public void TopOfPageKeepTogetherIgnoresSuppressedSpacingBefore() {
        PdfCreateOptions options = new() {
            PageWidth = global::OfficeIMO.OfficeLength.Points(320),
            PageHeight = global::OfficeIMO.OfficeLength.Points(170),
            MarginLeft = global::OfficeIMO.OfficeLength.Points(30),
            MarginRight = global::OfficeIMO.OfficeLength.Points(30),
            MarginTop = global::OfficeIMO.OfficeLength.Points(30),
            MarginBottom = global::OfficeIMO.OfficeLength.Points(30),
            DefaultFontSize = global::OfficeIMO.OfficeLength.Points(10)
        };

        byte[] bytes = PdfDocument.Create(options)
            .Paragraph(
                paragraph => paragraph.Text("This visible paragraph fits."),
                style: new PdfParagraphStyle {
                    KeepTogether = true,
                    SpacingBefore = global::OfficeIMO.OfficeLength.Points(500)
                })
            .ToBytes();

        Assert.NotEmpty(bytes);
    }

    [Fact]
    public void KeepTogetherCollapsesSpacingBeforeAfterMovingToFreshPage() {
        PdfCreateOptions options = new() {
            PageWidth = global::OfficeIMO.OfficeLength.Points(320),
            PageHeight = global::OfficeIMO.OfficeLength.Points(170),
            MarginLeft = global::OfficeIMO.OfficeLength.Points(30),
            MarginRight = global::OfficeIMO.OfficeLength.Points(30),
            MarginTop = global::OfficeIMO.OfficeLength.Points(30),
            MarginBottom = global::OfficeIMO.OfficeLength.Points(30),
            DefaultFontSize = global::OfficeIMO.OfficeLength.Points(10)
        };

        byte[] bytes = PdfDocument.Create(options)
            .Paragraph(paragraph => paragraph.Text("Preceding content."))
            .Paragraph(
                paragraph => paragraph.Text("This paragraph fits after its spacing collapses."),
                style: new PdfParagraphStyle {
                    KeepTogether = true,
                    SpacingBefore = global::OfficeIMO.OfficeLength.Points(500)
                })
            .ToBytes();

        Assert.NotEmpty(bytes);
    }


}

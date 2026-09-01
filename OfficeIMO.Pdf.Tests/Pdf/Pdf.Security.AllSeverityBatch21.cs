using System.Reflection;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfAllSeverityBatch21Tests {
    [Fact]
    public void UnprofiledExternalValidationDoesNotSatisfyProfiledProof() {
        PdfOptions options = new PdfOptions().ConfigurePdfAGroundwork(PdfComplianceProfile.PdfA3B);
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
    public void LegacyPublicTextAndFormSignaturesRemainAvailable() {
        Type stringValues = typeof(IReadOnlyDictionary<string, string>);

        AssertMethod(typeof(PdfEmbeddedFontFallbackSet), nameof(PdfEmbeddedFontFallbackSet.PlanText), typeof(string), typeof(string));
        AssertMethod(typeof(PdfEmbeddedFontFallbackSet), nameof(PdfEmbeddedFontFallbackSet.PlanTextRuns), typeof(string), typeof(string), typeof(PdfTextRun));
        AssertMethod(
            typeof(PdfEmbeddedFontFallbackSet),
            nameof(PdfEmbeddedFontFallbackSet.TryPlanTextRuns),
            typeof(string),
            typeof(IReadOnlyList<PdfTextRun>).MakeByRefType(),
            typeof(string),
            typeof(PdfTextRun),
            typeof(PdfConversionReport),
            typeof(string));

        Type diagnostics = typeof(PdfEmbeddedFontFallbackSet).Assembly.GetType("OfficeIMO.Pdf.PdfTextDiagnostics", throwOnError: true)!;
        AssertMethod(
            diagnostics,
            "PlanEmbeddedFontFallbackText",
            typeof(string),
            typeof(IEnumerable<PdfEmbeddedFontFallbackCandidate>),
            typeof(string));

        AssertMethod(
            typeof(PdfPageCanvas),
            nameof(PdfPageCanvas.Image),
            typeof(byte[]),
            typeof(double),
            typeof(double),
            typeof(double),
            typeof(double),
            typeof(PdfImageStyle),
            typeof(string),
            typeof(string),
            typeof(string),
            typeof(double));

        AssertMethod(typeof(PdfParagraphBuilder), nameof(PdfParagraphBuilder.Tab), typeof(PdfTabLeaderStyle));
        AssertMethod(typeof(PdfTextRun), nameof(PdfTextRun.Tab), typeof(PdfTabLeaderStyle));
        Assert.NotNull(typeof(PdfTextRun).GetConstructor(new[] {
            typeof(string),
            typeof(bool),
            typeof(bool),
            typeof(PdfColor?),
            typeof(bool),
            typeof(bool),
            typeof(double?),
            typeof(PdfStandardFont?),
            typeof(string),
            typeof(string),
            typeof(PdfTextBaseline),
            typeof(string),
            typeof(PdfTabLeaderStyle),
            typeof(PdfColor?),
            typeof(string)
        }));
    }

    [Fact]
    public void LegacyTwoArgumentFormCallsRemainSourceCompatibleWithoutMakingNullAmbiguous() {
        PdfDocumentForms forms = PdfDocument.Load(PdfDocument.Create().ToBytes()).Forms;
        var strings = new Dictionary<string, string>();
        var typed = new Dictionary<string, PdfFormFieldValue>();
        var formOptions = new PdfFormFillerOptions();

        _ = forms.TryFill(strings, formOptions);
        _ = forms.TryFill(typed, formOptions);
        _ = forms.TryFlatten(formOptions);
        _ = forms.TryFillAndFlatten(strings, formOptions);
        _ = forms.TryFillAndFlatten(typed, formOptions);

        _ = forms.TryFill(strings, null);
        _ = forms.TryFill(typed, null);
        _ = forms.TryFlatten(null);
        _ = forms.TryFillAndFlatten(strings, null);
        _ = forms.TryFillAndFlatten(typed, null);
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
        PdfOptions options = new() {
            PageWidth = 320,
            PageHeight = 170,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFontSize = 10
        };

        byte[] bytes = PdfDocument.Create(options)
            .Paragraph(
                paragraph => paragraph.Text("This visible paragraph fits."),
                style: new PdfParagraphStyle {
                    KeepTogether = true,
                    SpacingBefore = 500
                })
            .ToBytes();

        Assert.NotEmpty(bytes);
    }

    [Fact]
    public void KeepTogetherCollapsesSpacingBeforeAfterMovingToFreshPage() {
        PdfOptions options = new() {
            PageWidth = 320,
            PageHeight = 170,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFontSize = 10
        };

        byte[] bytes = PdfDocument.Create(options)
            .Paragraph(paragraph => paragraph.Text("Preceding content."))
            .Paragraph(
                paragraph => paragraph.Text("This paragraph fits after its spacing collapses."),
                style: new PdfParagraphStyle {
                    KeepTogether = true,
                    SpacingBefore = 500
                })
            .ToBytes();

        Assert.NotEmpty(bytes);
    }

    private static void AssertMethod(Type type, string name, params Type[] parameters) {
        Assert.NotNull(type.GetMethod(name, BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Static | BindingFlags.Instance, binder: null, parameters, modifiers: null));
    }
}

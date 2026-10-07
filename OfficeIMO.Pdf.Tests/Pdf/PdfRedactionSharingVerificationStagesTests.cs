using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfRedactionSharingVerificationStagesTests {
    [Fact]
    public void SharingChecksGlobalMarkersAndExternalValidatorAfterMetadataCleanup() {
        PdfDocument source = CreateSource();
        byte[] original = source.ToBytes();
        PdfRedactionPlan plan = source.Redactions.Search(new PdfRedactionSearchOptions().AddLiteral("PRIVATE-MARKER"));
        var validator = new FinalArtifactValidator();
        PdfRedactionVerificationOptions verification = CreateVerification(validator);

        PdfRedactionSharingResult result = source.Redactions.ApplyForSharing(plan,
            new PdfSanitizationOptions { ContentKindsToRemove = PdfSanitizationContentKind.UserMetadata },
            verificationOptions: verification);

        Assert.True(result.Summary.IsVerified);
        Assert.Null(PdfDocument.Load(result.ToBytes()).Inspect().Metadata.Author);
        Assert.Equal(1, validator.Calls);
        Assert.Equal(result.ToBytes(), validator.ValidatedBytes);
        Assert.Equal(original, source.ToBytes());
        Assert.Equal(new[] { "PRIVATE-MARKER" }, verification.RemovedTextMarkers);
        Assert.Equal(new[] { "PUBLIC-MARKER" }, verification.RetainedTextMarkers);
        Assert.Same(validator, Assert.Single(verification.ExternalValidators));
    }

    [Fact]
    public void SharingStillRejectsGlobalMarkerWhenPolicyRetainsMetadata() {
        PdfDocument source = CreateSource();
        byte[] original = source.ToBytes();
        PdfRedactionPlan plan = source.Redactions.Search(new PdfRedactionSearchOptions().AddLiteral("PRIVATE-MARKER"));
        PdfRedactionVerificationOptions verification = CreateVerification();

        InvalidOperationException error = Assert.Throws<InvalidOperationException>(() =>
            source.Redactions.ApplyForSharing(plan, new PdfSanitizationOptions {
                ContentKindsToRemove = PdfSanitizationContentKind.EmbeddedFiles
            }, verificationOptions: verification));

        Assert.Contains("PRIVATE-MARKER", error.Message);
        Assert.Equal(original, source.ToBytes());
    }

    [Fact]
    public void SharingStillRejectsFinalArtifactWhenExternalValidatorFails() {
        PdfDocument source = CreateSource();
        PdfRedactionPlan plan = source.Redactions.Search(new PdfRedactionSearchOptions().AddLiteral("PRIVATE-MARKER"));
        var validator = new FinalArtifactValidator { Accept = false };

        Assert.Throws<InvalidOperationException>(() => source.Redactions.ApplyForSharing(plan,
            new PdfSanitizationOptions { ContentKindsToRemove = PdfSanitizationContentKind.UserMetadata },
            verificationOptions: CreateVerification(validator)));

        Assert.Equal(1, validator.Calls);
        Assert.NotNull(validator.ValidatedBytes);
        Assert.Null(PdfDocument.Load(validator.ValidatedBytes!).Inspect().Metadata.Author);
    }

    private static PdfDocument CreateSource() => PdfDocument.Create(document => document.Content(content => content
        .Text("PRIVATE-MARKER").PageBreak().Text("PUBLIC-MARKER"))).Meta(author: "PRIVATE-MARKER");

    private static PdfRedactionVerificationOptions CreateVerification(IPdfRedactionExternalValidator? validator = null) {
        var result = new PdfRedactionVerificationOptions {
            RequireCompleteStreamInspection = true, CheckManagedRendering = true
        }.RequireRemovedText("PRIVATE-MARKER").RequireRetainedText("PUBLIC-MARKER");
        if (validator is not null) result.ExternalValidators.Add(validator);
        return result;
    }

    private sealed class FinalArtifactValidator : IPdfRedactionExternalValidator {
        internal int Calls { get; private set; }
        internal byte[]? ValidatedBytes { get; private set; }
        internal bool Accept { get; init; } = true;

        public PdfRedactionExternalValidationResult Validate(byte[] redactedPdf) {
            Calls++;
            ValidatedBytes = redactedPdf.ToArray();
            bool sanitized = PdfDocument.Load(redactedPdf).Inspect().Metadata.Author is null;
            return new PdfRedactionExternalValidationResult("final-artifact", Accept && sanitized);
        }
    }
}

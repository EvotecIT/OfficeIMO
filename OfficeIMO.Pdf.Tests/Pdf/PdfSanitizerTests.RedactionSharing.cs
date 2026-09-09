using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfSanitizerTests {
    [Fact]
    public void RedactionForSharingRechecksFinalSanitizedBytesAndPreservesPublicText() {
        PdfDocument source = PdfDocument.Load(BuildBeforeSharingPdf()).Sanitize(new PdfSanitizationOptions {
            ContentKindsToRemove = PdfSanitizationContentKind.Actions | PdfSanitizationContentKind.EmbeddedFiles |
                PdfSanitizationContentKind.OptionalContent,
            ActionKindsToRemove = PdfSanitizationActionKind.All
        }).ToDocument();
        PdfRedactionPlan plan = source.Redactions.Search(new PdfRedactionSearchOptions().AddLiteral("VISIBLE-PAGE-CONTENT"));
        var verification = new PdfRedactionVerificationOptions {
            RequireCompleteStreamInspection = true, CheckManagedRendering = true
        }.RequireRemovedText("VISIBLE-PAGE-CONTENT").RequireRetainedText("VISIBLE-AFTER-LAYER-FLATTEN");
        PdfRedactionSharingResult result = source.Redactions.ApplyForSharing(plan,
            new PdfSanitizationOptions {
                ContentKindsToRemove = PdfSanitizationContentKind.All,
                ActionKindsToRemove = PdfSanitizationActionKind.All
            }, verificationOptions: verification);

        Assert.True(result.Summary.IsVerified);
        Assert.True(result.Summary.SanitizedItemCount > 0);
        PdfDocument final = PdfDocument.Load(result.ToBytes());
        Assert.Null(final.Inspect().Metadata.Author);
        Assert.Empty(final.Inspect().Attachments);
        Assert.DoesNotContain("VISIBLE-PAGE-CONTENT", final.Read().Text);
        Assert.Contains("VISIBLE-AFTER-LAYER-FLATTEN", final.Read().Text);
        using var sha256 = System.Security.Cryptography.SHA256.Create();
        string expected = string.Concat(sha256.ComputeHash(result.ToBytes()).Select(value => value.ToString("X2")));
        Assert.Equal(expected, result.Summary.OutputSha256);
        byte[] mutableCopy = result.ToBytes();
        mutableCopy[0] = 0;
        Assert.Equal((byte)'%', result.ToBytes()[0]);
    }

    [Fact]
    public void RedactionForSharingDoesNotBypassActiveContentMutationGate() {
        PdfDocument source = PdfDocument.Load(BuildBeforeSharingPdf());
        PdfRedactionPlan plan = source.Redactions.Search(new PdfRedactionSearchOptions().AddLiteral("VISIBLE-PAGE-CONTENT"));
        Assert.Throws<PdfMutationBlockedException>(() => source.Redactions.ApplyForSharing(plan, new PdfSanitizationOptions()));
    }
}

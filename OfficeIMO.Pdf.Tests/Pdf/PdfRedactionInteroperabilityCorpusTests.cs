using OfficeIMO.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfRedactionInteroperabilityCorpusTests {
    [Theory]
    [InlineData("openpreserve-text-subset.pdf", CorpusOutcome.MutationBlocked)]
    [InlineData("openpreserve-pdfa1b-text.pdf", CorpusOutcome.Verified)]
    [InlineData("verapdf-tounicode-pass-a.pdf", CorpusOutcome.VerificationBlocked)]
    [InlineData("verapdf-optional-content.pdf", CorpusOutcome.MutationBlocked)]
    public void PinnedIndependentProducerCorpusCapturesVerifiedAndFailClosedOutcomes(string fileName, CorpusOutcome expectedOutcome) {
        byte[] source = File.ReadAllBytes(Path.Combine(FixtureRoot, fileName));
        PdfReadDocument readable = PdfReadDocument.Open(source);
        var target = readable.Pages
            .SelectMany(static (page, index) => page.GetTextSpans().Select(span => new {
                PageNumber = index + 1,
                Span = span,
                Bounds = PdfTextSpanGeometry.GetAxisAlignedBounds(span)
            }))
            .First(static item => !string.IsNullOrWhiteSpace(item.Span.Text) && item.Bounds.Width > 0D && item.Bounds.Height > 0D);
        PdfTextSpanBounds bounds = target.Bounds;
        PdfRedactionRegion region = PdfRedactionRegion.Quadrilateral(target.PageNumber, new[] {
            new PdfRedactionPoint(bounds.Left - 0.5D, bounds.Bottom - 0.5D),
            new PdfRedactionPoint(bounds.Right + 0.5D, bounds.Bottom - 0.5D),
            new PdfRedactionPoint(bounds.Right + 0.5D, bounds.Top + 0.5D),
            new PdfRedactionPoint(bounds.Left - 0.5D, bounds.Top + 0.5D)
        }, "independent-producer-corpus");
        PdfDocument document = PdfDocument.Load(source);
        PdfRedactionPlan plan = document.Redactions.Plan(new[] { region });
        var verificationOptions = new PdfRedactionVerificationOptions {
            CheckManagedRendering = true,
            RequireCompleteStreamInspection = true
        };
        verificationOptions.ExternalValidators.Add(new PdfPigTextAbsenceValidator(target.Span.Text));

        if (expectedOutcome == CorpusOutcome.MutationBlocked) {
            PdfMutationBlockedException exception = Assert.Throws<PdfMutationBlockedException>(() => document.Redactions.ApplyWithEvidence(
                plan,
                new PdfRedactionApplyOptions(),
                verificationOptions));
            Assert.Contains(exception.Plan.BlockerCodes, static code => code.StartsWith("FullRewrite.", StringComparison.Ordinal));
            return;
        }

        PdfRedactionApplyResult result = document.Redactions.ApplyWithEvidence(
            plan,
            new PdfRedactionApplyOptions(),
            verificationOptions);

        Assert.DoesNotContain(target.Span.Text, PdfDocument.Load(result.Pdf).Reader.Text(), StringComparison.Ordinal);
        Assert.True(Assert.Single(result.Evidence.Verification.ExternalValidationResults).IsValid);
        if (expectedOutcome == CorpusOutcome.Verified) {
            Assert.True(result.Evidence.Verification.IsVerified, string.Join("; ", result.Evidence.Verification.Issues.Select(static issue => issue.Message)));
        } else {
            Assert.False(result.Evidence.Verification.IsVerified);
            Assert.Contains(result.Evidence.Verification.Issues, static issue =>
                issue.Feature == "RedactionPlanResidual" && issue.Marker.StartsWith("VectorPath@", StringComparison.Ordinal));
        }
    }

    private sealed class PdfPigTextAbsenceValidator : IPdfRedactionExternalValidator {
        private readonly string _target;
        internal PdfPigTextAbsenceValidator(string target) { _target = target; }

        public PdfRedactionExternalValidationResult Validate(byte[] redactedPdf) {
            using var document = PdfPigDocument.Open(new MemoryStream(redactedPdf));
            string extracted = string.Join("\n", document.GetPages().Select(static page => page.Text));
            bool valid = !extracted.Contains(_target, StringComparison.Ordinal);
            return new PdfRedactionExternalValidationResult(
                "PdfPig independent extraction",
                valid,
                valid ? null : "The independently extracted target text remains present.");
        }
    }

    private static string FixtureRoot => Path.Combine(AppContext.BaseDirectory, "Pdf", "Fixtures", "Interoperability");

    public enum CorpusOutcome {
        Verified,
        MutationBlocked,
        VerificationBlocked
    }
}

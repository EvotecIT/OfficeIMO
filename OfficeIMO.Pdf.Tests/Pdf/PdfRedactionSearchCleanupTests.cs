using System.Text;
using OfficeIMO.Pdf;
using Xunit;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;

namespace OfficeIMO.Tests.Pdf;

public class PdfRedactionSearchCleanupTests {
    [Fact]
    public void SearchAndApply_RemovesLiteralRegexFieldMetadataAndAttachmentResidue() {
        var pdfOptions = new PdfOptions().AddEmbeddedFile("secret.txt", Encoding.UTF8.GetBytes("ATTACHMENT-SECRET"), "text/plain", PdfAssociatedFileRelationship.Data);
        byte[] source = PdfDocument.Create(pdfOptions)
            .Meta(title: "SECRET-METADATA", author: "Sensitive Author")
            .Paragraph(p => p.Text("Payroll token SECRET-2026 and SSN 123-45-6789"))
            .TextField("Customer.Secret", value: "FIELD-SECRET")
            .ToBytes();
        var search = new PdfRedactionSearchOptions()
            .AddLiteral("SECRET-2026")
            .AddRegex(@"\d{3}-\d{2}-\d{4}")
            .AddFormField("Customer.Secret");

        PdfRedactionPlan plan = PdfRedactionPlanner.Search(source, search);
        byte[] redacted = PdfRedactionApplier.Apply(source, plan, new PdfRedactionApplyOptions { CleanupScope = PdfRedactionCleanupScope.Metadata | PdfRedactionCleanupScope.Attachments });
        PdfDocumentInfo info = PdfInspector.Inspect(redacted);
        var verificationOptions = new PdfRedactionVerificationOptions { CheckManagedRendering = true }
            .RequireRemovedText("SECRET-2026", "123-45-6789", "FIELD-SECRET", "ATTACHMENT-SECRET", "SECRET-METADATA");
        verificationOptions.ExternalValidators.Add(new RawMarkerValidator("ATTACHMENT-SECRET"));
        verificationOptions.ExternalValidators.Add(new PdfPigTextValidator("SECRET-2026", "123-45-6789", "FIELD-SECRET"));
        PdfRedactionVerificationReport verification = PdfRedactionVerification.AssertVerified(redacted, verificationOptions);

        Assert.True(plan.IsSearchDriven);
        Assert.Equal(3, plan.SearchCriteria.Count);
        Assert.True(plan.Areas.Count >= 2);
        Assert.Empty(info.FormFields);
        Assert.Empty(info.Attachments);
        Assert.Null(info.Metadata.Title);
        Assert.True(verification.IsVerified);
        Assert.True(verification.ManagedRenderingChecked);
        Assert.Equal(2, verification.ExternalValidationResults.Count);
        Assert.All(verification.ExternalValidationResults, result => Assert.True(result.IsValid));
    }

    [Fact]
    public void Search_LogicalKindProducesReviewablePlanWithoutApplyingIt() {
        byte[] source = PdfDocument.Create().H1("Confidential heading").Paragraph(p => p.Text("Retained paragraph")).ToBytes();

        PdfRedactionPlan plan = PdfDocument.Open(source).SearchRedactions(new PdfRedactionSearchOptions().AddLogicalKind(PdfLogicalElementKind.Heading));

        Assert.True(plan.IsSearchDriven);
        Assert.Single(plan.SearchCriteria);
        Assert.Contains(plan.Areas, static area => area.Label == "logical-kind:Heading");
        Assert.Contains("Confidential heading", PdfTextExtractor.ExtractAllText(source), StringComparison.Ordinal);
    }

    [Fact]
    public void Apply_RemovesIntersectingPaintedPathsAndKeepsUnrelatedPaths() {
        const string content = "0 0 0 rg 10 10 40 40 re f 0 0 1 rg 120 120 30 30 re f";
        byte[] source = Encoding.ASCII.GetBytes(string.Join("\n", new[] {
            "%PDF-1.7",
            "1 0 obj", "<< /Type /Catalog /Pages 2 0 R >>", "endobj",
            "2 0 obj", "<< /Type /Pages /Count 1 /Kids [3 0 R] >>", "endobj",
            "3 0 obj", "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Contents 4 0 R >>", "endobj",
            "4 0 obj", "<< /Length " + content.Length.ToString(System.Globalization.CultureInfo.InvariantCulture) + " >>", "stream", content, "endstream", "endobj",
            "trailer", "<< /Root 1 0 R /Size 5 >>", "%%EOF"
        }));

        byte[] redacted = PdfRedactionApplier.Apply(source, new[] { new PdfRedactionArea(1, 15, 15, 10, 10) });
        string raw = PdfEncoding.Latin1GetString(redacted);

        Assert.DoesNotContain("10 10 40 40 re f", raw, StringComparison.Ordinal);
        Assert.Contains("120 120 30 30 re f", raw, StringComparison.Ordinal);
    }

    [Fact]
    public void Apply_AttachmentCleanupPrunesPageAssociatedFilePayloads() {
        byte[] source = PdfAssociatedFileTestSupport.BuildPageAssociatedFilePdf();

        byte[] redacted = PdfRedactionApplier.Apply(
            source,
            new[] { new PdfRedactionArea(1, 0, 0, 1, 1) },
            new PdfRedactionApplyOptions { CleanupScope = PdfRedactionCleanupScope.Attachments });

        Assert.Empty(PdfAttachmentExtractor.ExtractAttachments(redacted));
        string raw = Encoding.ASCII.GetString(redacted);
        Assert.DoesNotContain("/AF", raw, StringComparison.Ordinal);
        Assert.DoesNotContain(PdfAssociatedFileTestSupport.Payload, raw, StringComparison.Ordinal);
    }

    [Fact]
    public void Apply_PathScrubbingHonorsCallerDecodedStreamLimit() {
        const string content = "10 10 40 40 re f 120 120 30 30 re f";
        string encoded = string.Concat(Encoding.ASCII.GetBytes(content).Select(static value => value.ToString("X2", System.Globalization.CultureInfo.InvariantCulture))) + ">";
        byte[] source = Encoding.ASCII.GetBytes(string.Join("\n", new[] {
            "%PDF-1.7",
            "1 0 obj", "<< /Type /Catalog /Pages 2 0 R >>", "endobj",
            "2 0 obj", "<< /Type /Pages /Count 1 /Kids [3 0 R] >>", "endobj",
            "3 0 obj", "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Contents 4 0 R >>", "endobj",
            "4 0 obj", "<< /Filter /ASCIIHexDecode /Length " + encoded.Length.ToString(System.Globalization.CultureInfo.InvariantCulture) + " >>", "stream", encoded, "endstream", "endobj",
            "trailer", "<< /Root 1 0 R /Size 5 >>", "%%EOF"
        }));
        PdfRedactionPlan plan = PdfRedactionPlanner.Plan(source, new[] { new PdfRedactionArea(1, 15, 15, 10, 10) });
        var readOptions = new PdfReadOptions { Limits = new PdfReadLimits { MaxDecodedStreamBytes = 24 } };

        PdfMutationBlockedException exception = Assert.Throws<PdfMutationBlockedException>(() =>
            PdfRedactionApplier.Apply(source, plan, readOptions: readOptions));

        Assert.Contains("Read.ParserUnsupported", exception.Plan.BlockerCodes);
        Assert.Contains("DecodedStreamBytes", exception.Message, StringComparison.Ordinal);
        Assert.Contains("maximum 24", exception.Message, StringComparison.Ordinal);
    }

    private sealed class RawMarkerValidator : IPdfRedactionExternalValidator {
        private readonly string _marker;
        internal RawMarkerValidator(string marker) { _marker = marker; }
        public PdfRedactionExternalValidationResult Validate(byte[] redactedPdf) { bool valid = !PdfEncoding.Latin1GetString(redactedPdf).Contains(_marker); return new PdfRedactionExternalValidationResult("test-forensic-validator", valid, valid ? null : "Marker remains in raw bytes."); }
    }

    private sealed class PdfPigTextValidator : IPdfRedactionExternalValidator {
        private readonly string[] _markers;
        internal PdfPigTextValidator(params string[] markers) { _markers = markers; }
        public PdfRedactionExternalValidationResult Validate(byte[] redactedPdf) {
            using var document = PdfPigDocument.Open(new MemoryStream(redactedPdf));
            string text = string.Join("\n", document.GetPages().Select(page => page.Text));
            string? remaining = _markers.FirstOrDefault(marker => text.Contains(marker, StringComparison.Ordinal));
            return new PdfRedactionExternalValidationResult("PdfPig", remaining is null, remaining is null ? null : "Marker remains extractable: " + remaining);
        }
    }
}

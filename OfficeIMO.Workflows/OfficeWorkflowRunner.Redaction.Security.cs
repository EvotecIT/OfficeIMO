using OfficeIMO.Pdf;

namespace OfficeIMO.Workflows;

public sealed partial class OfficeWorkflowRunner {
    private static (byte[] Pdf, RedactionSignatureEvidence Evidence) ApplyDerivativeSignature(
        byte[] pdf,
        PdfRedactionWorkflowRequest request,
        PdfLoadOptions readOptions,
        int sourceSignatureCount,
        CancellationToken cancellationToken) {
        byte[] output = pdf;
        string? signerName = null;
        if (request.Recipe.SignaturePolicy == PdfRedactionSignaturePolicy.CreateAndSignDerivative) {
            if (request.OutputSigner is null) throw new RedactionWorkflowException("CreateAndSignDerivative requires a runtime output signer.");
            PdfExternalSignatureOptions options = request.OutputSignatureOptions ?? new PdfExternalSignatureOptions();
            options.CancellationToken = cancellationToken;
            PdfExternalSignatureCompletion completion = PdfDocument.Load(output, readOptions).Security.SignExternal(request.OutputSigner, options);
            output = completion.Pdf;
            signerName = completion.SignerName;
        }

        PdfSignatureValidationReport report = request.OutputSignatureValidator is null
            ? PdfDocument.Load(output, readOptions).Security.ValidateSignatures()
            : PdfDocument.Load(output, readOptions).Security.ValidateSignatures(request.OutputSignatureValidator);
        bool shouldBeSigned = request.Recipe.SignaturePolicy == PdfRedactionSignaturePolicy.CreateAndSignDerivative;
        if (shouldBeSigned) {
            if (report.SignatureCount != 1 || !report.IsStructurallyValid) {
                throw new RedactionWorkflowException("The signed derivative did not contain exactly one structurally valid output signature.");
            }
            if (request.OutputSignatureValidator is not null && (!report.MathematicalSignaturesVerified || !report.DigestVerified)) {
                throw new RedactionWorkflowException("The caller-provided validator rejected output signature math or its signed-content digest.");
            }
        } else if (report.HasSignatures) {
            throw new RedactionWorkflowException("The unsigned derivative retained an invalidated signature definition.");
        }

        return (output, new RedactionSignatureEvidence(
            sourceSignatureCount,
            report.SignatureCount,
            signerName,
            request.OutputSignatureValidator is not null && report.MathematicalSignaturesVerified && report.DigestVerified));
    }

    private static (byte[] Pdf, PdfLoadOptions ReadOptions) CreateSignatureFreeVerificationArtifact(
        byte[] output,
        PdfLoadOptions outputReadOptions,
        PdfRedactionWorkflowRequest request,
        CancellationToken cancellationToken) {
        if (request.Recipe.SignaturePolicy != PdfRedactionSignaturePolicy.CreateAndSignDerivative) return (output, outputReadOptions);
        PdfUnsignedDerivativeResult derivative = PdfDocument.Load(output, outputReadOptions).Security.CreateUnsignedDerivative(cancellationToken);
        return (derivative.Pdf, PdfLoadOptions.Default);
    }

    private static RedactionSignatureEvidence InspectDerivativeSignature(
        byte[] output,
        PdfLoadOptions readOptions,
        PdfRedactionWorkflowRequest request,
        int sourceSignatureCount) {
        PdfSignatureValidationReport report = request.OutputSignatureValidator is null
            ? PdfDocument.Load(output, readOptions).Security.ValidateSignatures()
            : PdfDocument.Load(output, readOptions).Security.ValidateSignatures(request.OutputSignatureValidator);
        bool shouldBeSigned = request.Recipe.SignaturePolicy == PdfRedactionSignaturePolicy.CreateAndSignDerivative;
        if (shouldBeSigned && (report.SignatureCount != 1 || !report.IsStructurallyValid)) {
            throw new RedactionWorkflowException("The existing derivative does not contain exactly one structurally valid output signature.");
        }
        if (!shouldBeSigned && report.HasSignatures) throw new RedactionWorkflowException("The existing output contains a signature not allowed by the selected derivative policy.");
        if (request.OutputSignatureValidator is not null && shouldBeSigned && (!report.MathematicalSignaturesVerified || !report.DigestVerified)) {
            throw new RedactionWorkflowException("The caller-provided validator rejected output signature math or its signed-content digest.");
        }
        return new RedactionSignatureEvidence(
            sourceSignatureCount,
            report.SignatureCount,
            null,
            request.OutputSignatureValidator is not null && report.MathematicalSignaturesVerified && report.DigestVerified);
    }

    private static IReadOnlyList<string> ValidateExternalArtifact(
        byte[] output,
        PdfRedactionWorkflowRequest request,
        PdfLoadOptions readOptions,
        CancellationToken cancellationToken) {
        if (request.ExternalValidators.Count == 0) return Array.Empty<string>();
        var options = new PdfRedactionVerificationOptions {
            RequireCompleteStreamInspection = true,
            CheckManagedRendering = true,
            CancellationToken = cancellationToken
        };
        foreach (IPdfRedactionCancellationAwareExternalValidator validator in request.ExternalValidators) options.ExternalValidators.Add(validator);
        PdfRedactionVerificationReport report = PdfDocument.Load(output, readOptions).Redactions.Verify(options);
        if (!report.IsVerified) throw new RedactionWorkflowException("One or more independent validators rejected the final redacted artifact.");
        return report.ExternalValidationResults.Select(static result => result.ValidatorName).Distinct(StringComparer.Ordinal).ToArray();
    }

    private readonly struct RedactionSignatureEvidence {
        internal RedactionSignatureEvidence(int sourceCount, int outputCount, string? signerName, bool cryptographicallyVerified) {
            SourceCount = sourceCount;
            OutputCount = outputCount;
            SignerName = signerName;
            CryptographicallyVerified = cryptographicallyVerified;
        }
        internal int SourceCount { get; }
        internal int OutputCount { get; }
        internal string? SignerName { get; }
        internal bool CryptographicallyVerified { get; }
    }
}

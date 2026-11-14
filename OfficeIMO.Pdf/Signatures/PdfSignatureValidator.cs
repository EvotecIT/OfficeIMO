namespace OfficeIMO.Pdf;

/// <summary>Dependency-free PDF signature structure validator.</summary>
internal static class PdfSignatureValidator {
    /// <summary>Validates signature structure, byte ranges, and preservation markers in a PDF byte array.</summary>
    public static PdfSignatureValidationReport Validate(byte[] pdf, PdfLoadOptions? options = null) {
        return ValidateCore(pdf, cryptographyProvider: null, options, security: null);
    }

    /// <summary>Validates signature structure using explicit read settings and cooperative cancellation.</summary>
    internal static PdfSignatureValidationReport Validate(byte[] pdf, PdfLoadOptions? options, CancellationToken cancellationToken) =>
        ValidateCore(pdf, cryptographyProvider: null, options, security: null, cancellationToken);

    /// <summary>Validates PDF signature structure and delegates CMS, trust, timestamp, and revocation policy to an optional provider.</summary>
    public static PdfSignatureValidationReport Validate(
        byte[] pdf,
        IPdfSignatureCryptographyProvider cryptographyProvider,
        PdfLoadOptions? options = null) {
        Guard.NotNull(cryptographyProvider, nameof(cryptographyProvider));
        return ValidateCore(pdf, cryptographyProvider, options, security: null);
    }

    internal static PdfSignatureValidationReport Validate(byte[] pdf, PdfDocumentSecurityInfo security) {
        Guard.NotNull(security, nameof(security));
        return ValidateCore(pdf, cryptographyProvider: null, options: null, security);
    }

    internal static PdfSignatureValidationReport Validate(
        byte[] pdf,
        PdfDocumentSecurityInfo security,
        CancellationToken cancellationToken) {
        Guard.NotNull(security, nameof(security));
        return ValidateCore(pdf, cryptographyProvider: null, options: null, security, cancellationToken);
    }

    private static PdfSignatureValidationReport ValidateCore(
        byte[] pdf,
        IPdfSignatureCryptographyProvider? cryptographyProvider,
        PdfLoadOptions? options,
        PdfDocumentSecurityInfo? security,
        CancellationToken cancellationToken = default) {
        Guard.NotNull(pdf, nameof(pdf));
        cancellationToken.ThrowIfCancellationRequested();

        bool objectGraphParsed = true;
        string? objectGraphError = null;
        PdfReadDocument? document = null;
        if (security == null) {
            try {
                document = PdfReadDocument.Open(pdf, options, cancellationToken);
                security = PdfInspector.Inspect(pdf, document, cancellationToken).Security;
            } catch (Exception ex) when (ex is not OperationCanceledException && ex is not OutOfMemoryException && ex is not StackOverflowException) {
                objectGraphParsed = false;
                objectGraphError = ex.Message;
                security = PdfInspector.Probe(pdf, options, cancellationToken).Security;
            }
        }

        var findings = new List<PdfSignatureValidationFinding>();
        if (!objectGraphParsed) {
            findings.Add(new PdfSignatureValidationFinding(
                PdfDiagnosticSeverity.Error,
                "signature-object-graph-parse-failed",
                "PDF objects could not be parsed for signature validation: " + objectGraphError));
        }

        if (!security.HasSignatures) {
            findings.Add(new PdfSignatureValidationFinding(
                PdfDiagnosticSeverity.Info,
                "no-signatures",
                "No PDF signature fields, signature values, or /ByteRange markers were detected."));
        } else if (security.Signatures.Count == 0) {
            findings.Add(new PdfSignatureValidationFinding(
                PdfDiagnosticSeverity.Error,
                "unreadable-signature",
                "PDF signature markers were detected, but no complete signature dictionary could be validated."));
        }

        var context = new CryptographicContext();
        if (cryptographyProvider is not null) {
            context = CreateCryptographicContext(pdf, security, document, cryptographyProvider, options, findings, cancellationToken);
        }

        var signatureResults = new List<PdfSignatureValidationResult>(security.Signatures.Count);
        foreach (PdfSignatureInfo signature in security.Signatures) {
            cancellationToken.ThrowIfCancellationRequested();
            signatureResults.Add(ValidateSignature(
                signature,
                pdf,
                security,
                cryptographyProvider,
                context,
                findings,
                cancellationToken));
        }

        cancellationToken.ThrowIfCancellationRequested();
        AddDocumentLevelFindings(security, findings, cryptographyProvider is not null);

        return new PdfSignatureValidationReport(
            security,
            pdf.LongLength,
            signatureResults.AsReadOnly(),
            findings.AsReadOnly(),
            objectGraphParsed,
            objectGraphError);
    }

    /// <summary>Validates signature structure, byte ranges, and preservation markers in a PDF file.</summary>
    public static PdfSignatureValidationReport Validate(string path, PdfLoadOptions? options = null) {
        PdfDocumentSource source = PdfDocumentSource.FromPath(path, options);
        return Validate(source.Bytes, source.Options);
    }

    /// <summary>Validates PDF signatures in a file through an optional cryptography provider.</summary>
    public static PdfSignatureValidationReport Validate(
        string path,
        IPdfSignatureCryptographyProvider cryptographyProvider,
        PdfLoadOptions? options = null) {
        Guard.NotNullOrWhiteSpace(path, nameof(path));
        Guard.NotNull(cryptographyProvider, nameof(cryptographyProvider));
        PdfDocumentSource source = PdfDocumentSource.FromPath(path, options);
        return Validate(source.Bytes, cryptographyProvider, source.Options);
    }

    /// <summary>Validates signature structure, byte ranges, and preservation markers in a readable PDF stream.</summary>
    public static PdfSignatureValidationReport Validate(Stream stream, PdfLoadOptions? options = null) {
        PdfDocumentSource source = PdfDocumentSource.FromRemainingStream(stream, options);
        return Validate(source.Bytes, source.Options);
    }

    /// <summary>Validates PDF signatures from a readable stream through an optional cryptography provider.</summary>
    public static PdfSignatureValidationReport Validate(
        Stream stream,
        IPdfSignatureCryptographyProvider cryptographyProvider,
        PdfLoadOptions? options = null) {
        Guard.NotNull(stream, nameof(stream));
        Guard.NotNull(cryptographyProvider, nameof(cryptographyProvider));
        PdfDocumentSource source = PdfDocumentSource.FromRemainingStream(stream, options);
        return Validate(source.Bytes, cryptographyProvider, source.Options);
    }

    private static PdfSignatureValidationResult ValidateSignature(
        PdfSignatureInfo signature,
        byte[] pdf,
        PdfDocumentSecurityInfo security,
        IPdfSignatureCryptographyProvider? cryptographyProvider,
        CryptographicContext context,
        List<PdfSignatureValidationFinding> aggregateFindings,
        CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        long fileLength = pdf.LongLength;
        var findings = new List<PdfSignatureValidationFinding>();
        IReadOnlyList<long> values = signature.ByteRangeValues;
        bool completeShape = signature.HasByteRange && values.Count == 4;
        bool ordered = false;
        bool coversEndOfFile = false;
        long? coveredBytes = null;
        long? gapStart = null;
        long? gapLength = null;
        long? unsignedByteCount = null;
        double? byteRangeCoverageRatio = null;

        if (!signature.HasByteRange) {
            AddSignatureFinding(findings, signature, PdfDiagnosticSeverity.Error, "signature-missing-byte-range", "Signature value does not contain a readable /ByteRange array.");
        } else if (!completeShape) {
            AddSignatureFinding(findings, signature, PdfDiagnosticSeverity.Error, "signature-unsupported-byte-range-shape", "Signature /ByteRange contains " + values.Count.ToString(System.Globalization.CultureInfo.InvariantCulture) + " numeric values; OfficeIMO.Pdf validates the common four-value detached signature shape.");
        } else {
            long firstStart = values[0];
            long firstLength = values[1];
            long secondStart = values[2];
            long secondLength = values[3];
            long firstEnd = firstStart + firstLength;
            long secondEnd = secondStart + secondLength;
            bool nonNegative = firstStart >= 0 && firstLength >= 0 && secondStart >= 0 && secondLength >= 0;
            bool inBounds = nonNegative && firstEnd >= firstStart && secondEnd >= secondStart && secondEnd <= fileLength;
            ordered = inBounds && firstStart == 0 && firstEnd <= secondStart;
            coversEndOfFile = inBounds && secondEnd == fileLength;
            coveredBytes = inBounds ? firstLength + secondLength : null;
            gapStart = inBounds ? firstEnd : null;
            gapLength = inBounds ? secondStart - firstEnd : null;
            unsignedByteCount = coveredBytes.HasValue ? Math.Max(0, fileLength - coveredBytes.Value) : null;
            byteRangeCoverageRatio = coveredBytes.HasValue && fileLength > 0 ? coveredBytes.Value / (double)fileLength : null;

            if (!nonNegative || !inBounds) {
                AddSignatureFinding(findings, signature, PdfDiagnosticSeverity.Error, "signature-byte-range-out-of-bounds", "Signature /ByteRange contains negative values, overflowing ranges, or ranges beyond the input file length.");
            } else if (!ordered) {
                AddSignatureFinding(findings, signature, PdfDiagnosticSeverity.Error, "signature-byte-range-not-ordered", "Signature /ByteRange segments are not ordered from the start of the file without overlap.");
            }

            if (gapLength.HasValue && gapLength.Value <= 0) {
                AddSignatureFinding(findings, signature, PdfDiagnosticSeverity.Error, "signature-byte-range-missing-contents-gap", "Signature /ByteRange does not leave an unsigned gap for the signature /Contents value.");
            }

            if (inBounds && !coversEndOfFile) {
                AddSignatureFinding(findings, signature, PdfDiagnosticSeverity.Warning, "signature-byte-range-does-not-cover-eof", "Signature /ByteRange does not end at the input file length; bytes after the signed ranges may be unsigned.");
            }

            if (byteRangeCoverageRatio.HasValue) {
                AddSignatureFinding(findings, signature, PdfDiagnosticSeverity.Info, "signature-byte-range-coverage", "Signature /ByteRange covers " + byteRangeCoverageRatio.Value.ToString("P2", System.Globalization.CultureInfo.InvariantCulture) + " of the input bytes.");
            }
        }

        if (!signature.HasContents) {
            AddSignatureFinding(findings, signature, PdfDiagnosticSeverity.Error, "signature-missing-contents", "Signature value does not contain a /Contents value.");
        } else if (signature.ContentsSizeBytes.GetValueOrDefault() == 0) {
            AddSignatureFinding(findings, signature, PdfDiagnosticSeverity.Warning, "signature-empty-contents", "Signature /Contents is present but empty.");
        }

        if (gapLength.HasValue &&
            signature.ContentsEncodedSizeBytes.HasValue &&
            gapLength.Value != signature.ContentsEncodedSizeBytes.Value) {
            AddSignatureFinding(findings, signature, PdfDiagnosticSeverity.Error, "signature-byte-range-contents-gap-mismatch", "Signature /ByteRange gap does not match the full /Contents token span.");
        }

        if (string.IsNullOrEmpty(signature.Filter)) {
            AddSignatureFinding(findings, signature, PdfDiagnosticSeverity.Warning, "signature-missing-filter", "Signature value does not identify a signature /Filter handler.");
        }

        if (string.IsNullOrEmpty(signature.SubFilter)) {
            AddSignatureFinding(findings, signature, PdfDiagnosticSeverity.Warning, "signature-missing-sub-filter", "Signature value does not identify a signature /SubFilter.");
        } else if (!signature.HasRecognizedSubFilter) {
            AddSignatureFinding(findings, signature, PdfDiagnosticSeverity.Warning, "signature-unknown-sub-filter", "Signature /SubFilter /" + signature.SubFilter + " is not one of the common CMS, CAdES, or RFC 3161 subfilters recognized by OfficeIMO.Pdf.");
        } else if (signature.IsDocumentTimestamp) {
            AddSignatureFinding(findings, signature, PdfDiagnosticSeverity.Info, "signature-document-timestamp-sub-filter", "Signature declares the ETSI.RFC3161 document timestamp subfilter.");
        } else if (signature.UsesCadesSubFilter) {
            AddSignatureFinding(findings, signature, PdfDiagnosticSeverity.Info, "signature-cades-sub-filter", "Signature declares the ETSI.CAdES.detached subfilter.");
        } else if (signature.UsesDetachedCmsSubFilter) {
            AddSignatureFinding(findings, signature, PdfDiagnosticSeverity.Info, "signature-detached-cms-sub-filter", "Signature declares a detached CMS/PKCS#7 subfilter.");
        }

        if (!signature.FieldObjectNumber.HasValue) {
            AddSignatureFinding(findings, signature, PdfDiagnosticSeverity.Warning, "signature-missing-owning-field", "Signature value is not linked from a readable AcroForm signature field.");
        }

        PdfSignatureCryptographicResult? cryptographicResult = null;
        PdfPadesLevel padesLevel = PdfPadesLevel.None;
        if (cryptographyProvider is not null) {
            if (context.DocumentTimestamps.TryGetValue(signature, out PrecomputedCryptography? precomputed)) {
                cryptographicResult = precomputed.Result;
                findings.AddRange(precomputed.Findings);
            } else {
                cryptographicResult = ValidateCryptographically(
                    pdf,
                    security,
                    signature,
                    completeShape,
                    ordered,
                    findings,
                    cryptographyProvider,
                    context.ValidationData,
                    cancellationToken);
            }
            if (cryptographicResult is not null && !signature.IsDocumentTimestamp) {
                padesLevel = DeterminePadesLevel(signature, cryptographicResult, context, findings);
            }
        }

        cancellationToken.ThrowIfCancellationRequested();
        aggregateFindings.AddRange(findings);
        return new PdfSignatureValidationResult(
            signature,
            completeShape,
            ordered,
            coversEndOfFile,
            coveredBytes,
            gapStart,
            gapLength,
            unsignedByteCount,
            byteRangeCoverageRatio,
            cryptographicResult,
            padesLevel,
            findings.AsReadOnly());
    }

    /// <summary>
    /// Validates document time-stamps first, so the DSS values present in each time-stamped revision gain a
    /// proof-of-existence time before the signatures that rely on them are validated.
    /// </summary>
    private static CryptographicContext CreateCryptographicContext(
        byte[] pdf,
        PdfDocumentSecurityInfo security,
        PdfReadDocument? document,
        IPdfSignatureCryptographyProvider provider,
        PdfLoadOptions? options,
        List<PdfSignatureValidationFinding> findings,
        CancellationToken cancellationToken) {
        var context = new CryptographicContext();
        int maxDecodedBytes = (options?.Limits ?? new PdfReadLimits()).MaxDecodedStreamBytes;
        if (document is not null && security.DocumentSecurityStore.HasValidationEvidence) {
            try {
                context.ValidationData = PdfDocumentValidationData.Read(
                    document.Objects, security.DocumentSecurityStore, maxDecodedBytes, cancellationToken);
            } catch (Exception ex) when (ex is not OperationCanceledException && ex is not OutOfMemoryException && ex is not StackOverflowException) {
                findings.Add(new PdfSignatureValidationFinding(
                    PdfDiagnosticSeverity.Warning,
                    "dss-validation-data-unreadable",
                    "Document Security Store values could not be decoded and were not used for validation: " + ex.Message,
                    isCryptographic: true));
            }
        }

        foreach (PdfSignatureInfo signature in security.Signatures) {
            cancellationToken.ThrowIfCancellationRequested();
            if (!signature.IsDocumentTimestamp) continue;
            bool usable = TryGetRevisionEnd(signature, pdf.LongLength, out long revisionEnd);
            var timestampFindings = new List<PdfSignatureValidationFinding>();
            PdfSignatureCryptographicResult? result = ValidateCryptographically(
                pdf,
                security,
                signature,
                usable,
                usable,
                timestampFindings,
                provider,
                context.ValidationData,
                cancellationToken);
            context.DocumentTimestamps[signature] = new PrecomputedCryptography(result, timestampFindings);
            if (usable &&
                result is { IsMathematicallyValid: true, TimestampStatus: PdfCryptographicValidationStatus.Valid, TimestampTime: DateTimeOffset time }) {
                context.ValidTimestampCoverage.Add(new TimestampCoverage(revisionEnd, time));
            }
        }

        if (context.ValidTimestampCoverage.Count > 0 && context.ValidationData.Revocation.Count > 0) {
            var proof = new Dictionary<string, DateTimeOffset>(StringComparer.Ordinal);
            foreach (TimestampCoverage coverage in context.ValidTimestampCoverage.OrderBy(static item => item.Time)) {
                cancellationToken.ThrowIfCancellationRequested();
                foreach (string hash in ReadRevisionValidationHashes(pdf, coverage.RevisionEnd, options, maxDecodedBytes, cancellationToken)) {
                    if (!proof.ContainsKey(hash)) proof[hash] = coverage.Time;
                }
            }
            context.ValidationData = context.ValidationData.WithProofOfExistence(proof);
        }
        return context;
    }

    /// <summary>Reads the DSS revocation values present in the revision a document time-stamp covers.</summary>
    private static string[] ReadRevisionValidationHashes(
        byte[] pdf,
        long revisionEnd,
        PdfLoadOptions? options,
        int maxDecodedBytes,
        CancellationToken cancellationToken) {
        try {
            byte[] revision = new byte[revisionEnd];
            Buffer.BlockCopy(pdf, 0, revision, 0, revision.Length);
            PdfReadDocument document = PdfReadDocument.OpenOwned(revision, options, cancellationToken);
            PdfDocumentSecurityInfo security = PdfInspector.Inspect(revision, document, cancellationToken).Security;
            return PdfDocumentValidationData.Read(document.Objects, security.DocumentSecurityStore, maxDecodedBytes, cancellationToken)
                .RevocationHashes.ToArray();
        } catch (Exception ex) when (ex is not OperationCanceledException && ex is not OutOfMemoryException && ex is not StackOverflowException) {
            return Array.Empty<string>();
        }
    }

    /// <summary>Determines the PAdES baseline level (ETSI EN 319 142-1 section 6) and records why the next level is missing.</summary>
    private static PdfPadesLevel DeterminePadesLevel(
        PdfSignatureInfo signature,
        PdfSignatureCryptographicResult result,
        CryptographicContext context,
        List<PdfSignatureValidationFinding> findings) {
        if (!signature.UsesCadesSubFilter) {
            AddLevelFinding(findings, signature, "pades-baseline-b-not-reached",
                "PAdES baseline signatures use the ETSI.CAdES.detached subfilter; this signature uses /" + signature.SubFilter + ".");
            return PdfPadesLevel.None;
        }
        if (!result.IsMathematicallyValid) {
            AddLevelFinding(findings, signature, "pades-baseline-b-not-reached",
                "PAdES B-B needs valid signature math and a matching signed-content digest.");
            return PdfPadesLevel.None;
        }
        if (!result.ReportsPadesEvidence) {
            AddLevelFinding(findings, signature, "pades-baseline-b-not-reached",
                result.ProviderName + " does not report the CMS attributes and revocation sources needed to establish a PAdES level.");
            return PdfPadesLevel.None;
        }
        if (!result.HasSigningCertificateAttribute) {
            AddLevelFinding(findings, signature, "pades-baseline-b-not-reached",
                "PAdES B-B needs the ESS signing-certificate-v2 signed attribute, which this signature does not carry.");
            return PdfPadesLevel.None;
        }

        bool coveredByDocumentTimestamp = TryGetRevisionEnd(signature, long.MaxValue, out long signatureEnd) &&
            context.ValidTimestampCoverage.Any(coverage => coverage.RevisionEnd >= signatureEnd);
        if (result.TimestampStatus != PdfCryptographicValidationStatus.Valid && !coveredByDocumentTimestamp) {
            AddLevelFinding(findings, signature, "pades-baseline-t-not-reached",
                result.TimestampStatus == PdfCryptographicValidationStatus.NotPerformed
                    ? "PAdES B-T needs a signature time-stamp or a later document time-stamp; neither is present and valid."
                    : "The signature time-stamp is " + result.TimestampStatus + " and no valid document time-stamp covers the signature.");
            return PdfPadesLevel.BaselineB;
        }

        if (result.RevocationStatus != PdfCryptographicValidationStatus.Valid) {
            AddLevelFinding(findings, signature, "pades-baseline-lt-not-reached",
                "PAdES B-LT needs valid revocation status for the signer and time-stamp authority paths; it is " +
                result.RevocationStatus + ". See the revocation findings for each certificate.");
            return PdfPadesLevel.BaselineT;
        }
        if (!result.RevocationSettledByDocument) {
            AddLevelFinding(findings, signature, "pades-baseline-lt-not-reached",
                "PAdES B-LT needs the revocation data in the Document Security Store; at least one certificate was settled only by data inside the signature or fetched online.");
            return PdfPadesLevel.BaselineT;
        }

        if (!result.RevocationEvidenceHasProofOfExistence) {
            AddLevelFinding(findings, signature, "pades-baseline-lta-not-reached",
                "PAdES B-LTA needs a valid document time-stamp covering the Document Security Store data used for validation.");
            return PdfPadesLevel.BaselineLT;
        }
        return PdfPadesLevel.BaselineLTA;
    }

    private static void AddLevelFinding(
        List<PdfSignatureValidationFinding> findings,
        PdfSignatureInfo signature,
        string code,
        string message) =>
        AddSignatureFinding(findings, signature, PdfDiagnosticSeverity.Info, code, message, isCryptographic: true);

    /// <summary>A usable four-value /ByteRange ends where the signed revision ends.</summary>
    private static bool TryGetRevisionEnd(PdfSignatureInfo signature, long fileLength, out long revisionEnd) {
        revisionEnd = 0;
        IReadOnlyList<long> values = signature.ByteRangeValues;
        if (!signature.HasByteRange || values.Count != 4) return false;
        long firstEnd = values[0] + values[1];
        long secondEnd = values[2] + values[3];
        bool usable = values[0] == 0 && values[1] >= 0 && values[2] >= firstEnd && values[3] >= 0 &&
            secondEnd >= values[2] && secondEnd <= fileLength;
        if (usable) revisionEnd = secondEnd;
        return usable;
    }

    private static PdfSignatureCryptographicResult? ValidateCryptographically(
        byte[] pdf,
        PdfDocumentSecurityInfo security,
        PdfSignatureInfo signature,
        bool completeShape,
        bool ordered,
        List<PdfSignatureValidationFinding> findings,
        IPdfSignatureCryptographyProvider provider,
        PdfDocumentValidationData validationData,
        CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        if (!completeShape || !ordered || signature.ContentsBytes is null || signature.ContentsBytes.Length == 0) {
            AddSignatureFinding(
                findings,
                signature,
                PdfDiagnosticSeverity.Error,
                "cryptographic-validation-input-unavailable",
                "Cryptographic validation requires a valid four-value /ByteRange and non-empty decoded /Contents bytes.",
                isCryptographic: true);
            return null;
        }

        try {
            byte[] signedContent = ReadSignedContent(pdf, signature.ByteRangeValues, cancellationToken);
            var input = new PdfSignatureCryptographyInput(
                signature,
                signedContent,
                (byte[])signature.ContentsBytes.Clone(),
                pdf.LongLength,
                security.DocumentSecurityStore,
                validationData);
            PdfSignatureCryptographicResult result = provider.Verify(input) ??
                throw new InvalidOperationException("Signature cryptography provider returned no result.");
            for (int i = 0; i < result.Findings.Count; i++) {
                cancellationToken.ThrowIfCancellationRequested();
                PdfSignatureCryptographicFinding finding = result.Findings[i];
                AddSignatureFinding(findings, signature, finding.Severity, finding.Code, finding.Message, isCryptographic: true);
            }

            return result;
        } catch (Exception ex) when (ex is not OperationCanceledException && ex is not OutOfMemoryException && ex is not StackOverflowException) {
            AddSignatureFinding(
                findings,
                signature,
                PdfDiagnosticSeverity.Error,
                "cryptographic-provider-failed",
                provider.Name + " could not validate the signature: " + ex.Message,
                isCryptographic: true);
            return new PdfSignatureCryptographicResult(
                provider.Name,
                PdfCryptographicValidationStatus.Error,
                PdfCryptographicValidationStatus.Error,
                PdfCryptographicValidationStatus.Error,
                PdfCryptographicValidationStatus.Error,
                PdfCryptographicValidationStatus.Error);
        }
    }

    private static byte[] ReadSignedContent(
        byte[] pdf,
        IReadOnlyList<long> byteRangeValues,
        CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        long totalLength = byteRangeValues[1] + byteRangeValues[3];
        if (totalLength > int.MaxValue) {
            throw new InvalidOperationException("Signed byte ranges exceed the current in-memory validation limit.");
        }

        var content = new byte[(int)totalLength];
        int destinationOffset = 0;
        for (int i = 0; i < byteRangeValues.Count; i += 2) {
            cancellationToken.ThrowIfCancellationRequested();
            int sourceOffset = checked((int)byteRangeValues[i]);
            int count = checked((int)byteRangeValues[i + 1]);
            while (count > 0) {
                cancellationToken.ThrowIfCancellationRequested();
                int chunkLength = Math.Min(count, 1024 * 1024);
                Buffer.BlockCopy(pdf, sourceOffset, content, destinationOffset, chunkLength);
                sourceOffset += chunkLength;
                destinationOffset += chunkLength;
                count -= chunkLength;
            }
        }

        return content;
    }

    private static void AddDocumentLevelFindings(
        PdfDocumentSecurityInfo security,
        List<PdfSignatureValidationFinding> findings,
        bool cryptographicValidationRequested) {
        if (security.HasSignatures && !cryptographicValidationRequested) {
            findings.Add(new PdfSignatureValidationFinding(
                PdfDiagnosticSeverity.Info,
                "cryptographic-trust-not-verified",
                "Signature structure was inspected without an optional cryptography provider; CMS math, trust, revocation, digest, and timestamp policy were not evaluated."));
        }

        if (security.AcroFormAppendOnly) {
            findings.Add(new PdfSignatureValidationFinding(
                PdfDiagnosticSeverity.Info,
                "acro-form-append-only",
                "AcroForm /SigFlags indicates append-only updates are expected."));
        }

        if (security.HasDocMDPPermissions) {
            string suffix = security.DocMDPPermissionLevel.HasValue
                ? " Permission level /P=" + security.DocMDPPermissionLevel.Value.ToString(System.Globalization.CultureInfo.InvariantCulture) + " was detected."
                : string.Empty;
            findings.Add(new PdfSignatureValidationFinding(
                PdfDiagnosticSeverity.Info,
                "doc-mdp-detected",
                "Catalog /Perms contains DocMDP certification permissions." + suffix,
                security.DocMDPSignatureObjectNumber));
        }

        if (security.HasUsageRights) {
            findings.Add(new PdfSignatureValidationFinding(
                PdfDiagnosticSeverity.Warning,
                "usage-rights-detected",
                "Catalog /Perms contains usage-rights entries that must be preserved by append-only mutation."));
        }

        if (security.HasLongTermValidationEvidence) {
            findings.Add(new PdfSignatureValidationFinding(
                PdfDiagnosticSeverity.Info,
                "long-term-validation-evidence-detected",
                "Document Security Store (/DSS) validation evidence was detected."));
        }

        if (security.HasIncrementalUpdates) {
            findings.Add(new PdfSignatureValidationFinding(
                PdfDiagnosticSeverity.Info,
                "incremental-updates-detected",
                "Incremental update markers were detected; signature-preserving changes should append a new revision."));
        }
    }

    private static void AddSignatureFinding(
        List<PdfSignatureValidationFinding> findings,
        PdfSignatureInfo signature,
        PdfDiagnosticSeverity severity,
        string code,
        string message,
        bool isCryptographic = false) {
        findings.Add(new PdfSignatureValidationFinding(
            severity,
            code,
            message,
            signature.ObjectNumber,
            signature.FieldObjectNumber,
            isCryptographic));
    }

    private sealed class CryptographicContext {
        internal PdfDocumentValidationData ValidationData { get; set; } = PdfDocumentValidationData.Empty;

        internal Dictionary<PdfSignatureInfo, PrecomputedCryptography> DocumentTimestamps { get; } = new();

        internal List<TimestampCoverage> ValidTimestampCoverage { get; } = new();
    }

    private sealed class PrecomputedCryptography {
        internal PrecomputedCryptography(PdfSignatureCryptographicResult? result, List<PdfSignatureValidationFinding> findings) {
            Result = result;
            Findings = findings;
        }

        internal PdfSignatureCryptographicResult? Result { get; }
        internal List<PdfSignatureValidationFinding> Findings { get; }
    }

    private readonly struct TimestampCoverage {
        internal TimestampCoverage(long revisionEnd, DateTimeOffset time) {
            RevisionEnd = revisionEnd;
            Time = time;
        }

        internal long RevisionEnd { get; }
        internal DateTimeOffset Time { get; }
    }
}

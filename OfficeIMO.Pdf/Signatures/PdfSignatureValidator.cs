namespace OfficeIMO.Pdf;

/// <summary>Dependency-free PDF signature structure validator.</summary>
public static class PdfSignatureValidator {
    /// <summary>Validates signature structure, byte ranges, and preservation markers in a PDF byte array.</summary>
    public static PdfSignatureValidationReport Validate(byte[] pdf, PdfReadOptions? options = null) {
        Guard.NotNull(pdf, nameof(pdf));

        PdfDocumentSecurityInfo security;
        bool objectGraphParsed = true;
        string? objectGraphError = null;
        try {
            security = PdfInspector.Inspect(pdf, options).Security;
        } catch (Exception ex) when (ex is not OutOfMemoryException && ex is not StackOverflowException) {
            objectGraphParsed = false;
            objectGraphError = ex.Message;
            security = PdfInspector.Probe(pdf).Security;
        }

        var findings = new List<PdfSignatureValidationFinding>();
        if (!objectGraphParsed) {
            findings.Add(new PdfSignatureValidationFinding(
                PdfDiagnosticSeverity.Error,
                "SignatureObjectGraphParseFailed",
                "PDF objects could not be parsed for signature validation: " + objectGraphError));
        }

        if (!security.HasSignatures) {
            findings.Add(new PdfSignatureValidationFinding(
                PdfDiagnosticSeverity.Info,
                "NoSignatures",
                "No PDF signature fields, signature values, or /ByteRange markers were detected."));
        }

        var signatureResults = new List<PdfSignatureValidationResult>(security.Signatures.Count);
        foreach (PdfSignatureInfo signature in security.Signatures) {
            signatureResults.Add(ValidateSignature(signature, pdf.LongLength, findings));
        }

        AddDocumentLevelFindings(security, findings);

        return new PdfSignatureValidationReport(
            security,
            pdf.LongLength,
            signatureResults.AsReadOnly(),
            findings.AsReadOnly(),
            objectGraphParsed,
            objectGraphError);
    }

    /// <summary>Validates signature structure, byte ranges, and preservation markers in a PDF file.</summary>
    public static PdfSignatureValidationReport Validate(string path, PdfReadOptions? options = null) {
        Guard.NotNullOrWhiteSpace(path, nameof(path));
        return Validate(File.ReadAllBytes(path), options);
    }

    /// <summary>Validates signature structure, byte ranges, and preservation markers in a readable PDF stream.</summary>
    public static PdfSignatureValidationReport Validate(Stream stream, PdfReadOptions? options = null) {
        Guard.NotNull(stream, nameof(stream));
        if (!stream.CanRead) {
            throw new ArgumentException("Stream must be readable.", nameof(stream));
        }

        using var buffer = new MemoryStream();
        stream.CopyTo(buffer);
        return Validate(buffer.ToArray(), options);
    }

    private static PdfSignatureValidationResult ValidateSignature(
        PdfSignatureInfo signature,
        long fileLength,
        List<PdfSignatureValidationFinding> aggregateFindings) {
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
            AddSignatureFinding(findings, signature, PdfDiagnosticSeverity.Error, "SignatureMissingByteRange", "Signature value does not contain a readable /ByteRange array.");
        } else if (!completeShape) {
            AddSignatureFinding(findings, signature, PdfDiagnosticSeverity.Error, "SignatureUnsupportedByteRangeShape", "Signature /ByteRange contains " + values.Count.ToString(System.Globalization.CultureInfo.InvariantCulture) + " numeric values; OfficeIMO.Pdf validates the common four-value detached signature shape.");
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
                AddSignatureFinding(findings, signature, PdfDiagnosticSeverity.Error, "SignatureByteRangeOutOfBounds", "Signature /ByteRange contains negative values, overflowing ranges, or ranges beyond the input file length.");
            } else if (!ordered) {
                AddSignatureFinding(findings, signature, PdfDiagnosticSeverity.Error, "SignatureByteRangeNotOrdered", "Signature /ByteRange segments are not ordered from the start of the file without overlap.");
            }

            if (gapLength.HasValue && gapLength.Value <= 0) {
                AddSignatureFinding(findings, signature, PdfDiagnosticSeverity.Error, "SignatureByteRangeMissingContentsGap", "Signature /ByteRange does not leave an unsigned gap for the signature /Contents value.");
            }

            if (inBounds && !coversEndOfFile) {
                AddSignatureFinding(findings, signature, PdfDiagnosticSeverity.Warning, "SignatureByteRangeDoesNotCoverEof", "Signature /ByteRange does not end at the input file length; bytes after the signed ranges may be unsigned.");
            }

            if (byteRangeCoverageRatio.HasValue) {
                AddSignatureFinding(findings, signature, PdfDiagnosticSeverity.Info, "SignatureByteRangeCoverage", "Signature /ByteRange covers " + byteRangeCoverageRatio.Value.ToString("P2", System.Globalization.CultureInfo.InvariantCulture) + " of the input bytes.");
            }
        }

        if (!signature.HasContents) {
            AddSignatureFinding(findings, signature, PdfDiagnosticSeverity.Error, "SignatureMissingContents", "Signature value does not contain a /Contents value.");
        } else if (signature.ContentsSizeBytes.GetValueOrDefault() == 0) {
            AddSignatureFinding(findings, signature, PdfDiagnosticSeverity.Warning, "SignatureEmptyContents", "Signature /Contents is present but empty.");
        }

        if (gapLength.HasValue &&
            signature.ContentsSizeBytes.HasValue &&
            gapLength.Value != (signature.ContentsSizeBytes.Value * 2L) + 2L) {
            AddSignatureFinding(findings, signature, PdfDiagnosticSeverity.Error, "SignatureByteRangeContentsGapMismatch", "Signature /ByteRange gap does not match the full /Contents hex literal length.");
        }

        if (string.IsNullOrEmpty(signature.Filter)) {
            AddSignatureFinding(findings, signature, PdfDiagnosticSeverity.Warning, "SignatureMissingFilter", "Signature value does not identify a signature /Filter handler.");
        }

        if (string.IsNullOrEmpty(signature.SubFilter)) {
            AddSignatureFinding(findings, signature, PdfDiagnosticSeverity.Warning, "SignatureMissingSubFilter", "Signature value does not identify a signature /SubFilter.");
        } else if (!signature.HasRecognizedSubFilter) {
            AddSignatureFinding(findings, signature, PdfDiagnosticSeverity.Warning, "SignatureUnknownSubFilter", "Signature /SubFilter /" + signature.SubFilter + " is not one of the common CMS, CAdES, or RFC 3161 subfilters recognized by OfficeIMO.Pdf.");
        } else if (signature.IsDocumentTimestamp) {
            AddSignatureFinding(findings, signature, PdfDiagnosticSeverity.Info, "SignatureDocumentTimestampSubFilter", "Signature declares the ETSI.RFC3161 document timestamp subfilter.");
        } else if (signature.UsesCadesSubFilter) {
            AddSignatureFinding(findings, signature, PdfDiagnosticSeverity.Info, "SignatureCadesSubFilter", "Signature declares the ETSI.CAdES.detached subfilter.");
        } else if (signature.UsesDetachedCmsSubFilter) {
            AddSignatureFinding(findings, signature, PdfDiagnosticSeverity.Info, "SignatureDetachedCmsSubFilter", "Signature declares a detached CMS/PKCS#7 subfilter.");
        }

        if (!signature.FieldObjectNumber.HasValue) {
            AddSignatureFinding(findings, signature, PdfDiagnosticSeverity.Warning, "SignatureMissingOwningField", "Signature value is not linked from a readable AcroForm signature field.");
        }

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
            findings.AsReadOnly());
    }

    private static void AddDocumentLevelFindings(PdfDocumentSecurityInfo security, List<PdfSignatureValidationFinding> findings) {
        if (security.HasSignatures) {
            findings.Add(new PdfSignatureValidationFinding(
                PdfDiagnosticSeverity.Info,
                "CryptographicTrustNotVerified",
                "Signature structure was inspected, but certificate-chain trust, revocation, digest, and CMS cryptographic verification are not performed by OfficeIMO.Pdf."));
        }

        if (security.AcroFormAppendOnly) {
            findings.Add(new PdfSignatureValidationFinding(
                PdfDiagnosticSeverity.Info,
                "AcroFormAppendOnly",
                "AcroForm /SigFlags indicates append-only updates are expected."));
        }

        if (security.HasDocMDPPermissions) {
            string suffix = security.DocMDPPermissionLevel.HasValue
                ? " Permission level /P=" + security.DocMDPPermissionLevel.Value.ToString(System.Globalization.CultureInfo.InvariantCulture) + " was detected."
                : string.Empty;
            findings.Add(new PdfSignatureValidationFinding(
                PdfDiagnosticSeverity.Info,
                "DocMDPDetected",
                "Catalog /Perms contains DocMDP certification permissions." + suffix,
                security.DocMDPSignatureObjectNumber));
        }

        if (security.HasUsageRights) {
            findings.Add(new PdfSignatureValidationFinding(
                PdfDiagnosticSeverity.Warning,
                "UsageRightsDetected",
                "Catalog /Perms contains usage-rights entries that must be preserved by append-only mutation."));
        }

        if (security.HasLongTermValidationEvidence) {
            findings.Add(new PdfSignatureValidationFinding(
                PdfDiagnosticSeverity.Info,
                "LongTermValidationEvidenceDetected",
                "Document Security Store (/DSS) validation evidence was detected."));
        }

        if (security.HasIncrementalUpdates) {
            findings.Add(new PdfSignatureValidationFinding(
                PdfDiagnosticSeverity.Info,
                "IncrementalUpdatesDetected",
                "Incremental update markers were detected; signature-preserving changes should append a new revision."));
        }
    }

    private static void AddSignatureFinding(
        List<PdfSignatureValidationFinding> findings,
        PdfSignatureInfo signature,
        PdfDiagnosticSeverity severity,
        string code,
        string message) {
        findings.Add(new PdfSignatureValidationFinding(
            severity,
            code,
            message,
            signature.ObjectNumber,
            signature.FieldObjectNumber));
    }
}

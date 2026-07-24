namespace OfficeIMO.Pdf;

/// <summary>Dependency-free PDF signature structure validator.</summary>
internal static class PdfSignatureValidator {
    /// <summary>Validates signature structure, byte ranges, and preservation markers in a PDF byte array.</summary>
    public static PdfSignatureValidationReport Validate(byte[] pdf, PdfReadOptions? options = null) {
        return ValidateCore(pdf, cryptographyProvider: null, options, security: null);
    }

    /// <summary>Validates PDF signature structure and delegates CMS, trust, timestamp, and revocation policy to an optional provider.</summary>
    public static PdfSignatureValidationReport Validate(
        byte[] pdf,
        IPdfSignatureCryptographyProvider cryptographyProvider,
        PdfReadOptions? options = null) {
        Guard.NotNull(cryptographyProvider, nameof(cryptographyProvider));
        return ValidateCore(pdf, cryptographyProvider, options, security: null);
    }

    internal static PdfSignatureValidationReport Validate(byte[] pdf, PdfDocumentSecurityInfo security) {
        Guard.NotNull(security, nameof(security));
        return ValidateCore(pdf, cryptographyProvider: null, options: null, security);
    }

    private static PdfSignatureValidationReport ValidateCore(
        byte[] pdf,
        IPdfSignatureCryptographyProvider? cryptographyProvider,
        PdfReadOptions? options,
        PdfDocumentSecurityInfo? security) {
        Guard.NotNull(pdf, nameof(pdf));

        bool objectGraphParsed = true;
        string? objectGraphError = null;
        if (security == null) {
            try {
                security = PdfInspector.Inspect(pdf, options).Security;
            } catch (Exception ex) when (ex is not OutOfMemoryException && ex is not StackOverflowException) {
                objectGraphParsed = false;
                objectGraphError = ex.Message;
                security = PdfInspector.Probe(pdf).Security;
            }
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
        } else if (security.Signatures.Count == 0) {
            findings.Add(new PdfSignatureValidationFinding(
                PdfDiagnosticSeverity.Error,
                "UnreadableSignature",
                "PDF signature markers were detected, but no complete signature dictionary could be validated."));
        }

        var signatureResults = new List<PdfSignatureValidationResult>(security.Signatures.Count);
        foreach (PdfSignatureInfo signature in security.Signatures) {
            signatureResults.Add(ValidateSignature(signature, pdf, security, cryptographyProvider, findings));
        }

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
    public static PdfSignatureValidationReport Validate(string path, PdfReadOptions? options = null) {
        Guard.NotNullOrWhiteSpace(path, nameof(path));
        return Validate(File.ReadAllBytes(path), options);
    }

    /// <summary>Validates PDF signatures in a file through an optional cryptography provider.</summary>
    public static PdfSignatureValidationReport Validate(
        string path,
        IPdfSignatureCryptographyProvider cryptographyProvider,
        PdfReadOptions? options = null) {
        Guard.NotNullOrWhiteSpace(path, nameof(path));
        Guard.NotNull(cryptographyProvider, nameof(cryptographyProvider));
        return Validate(File.ReadAllBytes(path), cryptographyProvider, options);
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

    /// <summary>Validates PDF signatures from a readable stream through an optional cryptography provider.</summary>
    public static PdfSignatureValidationReport Validate(
        Stream stream,
        IPdfSignatureCryptographyProvider cryptographyProvider,
        PdfReadOptions? options = null) {
        Guard.NotNull(stream, nameof(stream));
        Guard.NotNull(cryptographyProvider, nameof(cryptographyProvider));
        if (!stream.CanRead) {
            throw new ArgumentException("Stream must be readable.", nameof(stream));
        }

        using var buffer = new MemoryStream();
        stream.CopyTo(buffer);
        return Validate(buffer.ToArray(), cryptographyProvider, options);
    }

    private static PdfSignatureValidationResult ValidateSignature(
        PdfSignatureInfo signature,
        byte[] pdf,
        PdfDocumentSecurityInfo security,
        IPdfSignatureCryptographyProvider? cryptographyProvider,
        List<PdfSignatureValidationFinding> aggregateFindings) {
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
            signature.ContentsEncodedSizeBytes.HasValue &&
            gapLength.Value != signature.ContentsEncodedSizeBytes.Value) {
            AddSignatureFinding(findings, signature, PdfDiagnosticSeverity.Error, "SignatureByteRangeContentsGapMismatch", "Signature /ByteRange gap does not match the full /Contents token span.");
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

        PdfSignatureCryptographicResult? cryptographicResult = null;
        if (cryptographyProvider is not null) {
            cryptographicResult = ValidateCryptographically(
                pdf,
                security,
                signature,
                completeShape,
                ordered,
                findings,
                cryptographyProvider);
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
            cryptographicResult,
            findings.AsReadOnly());
    }

    private static PdfSignatureCryptographicResult? ValidateCryptographically(
        byte[] pdf,
        PdfDocumentSecurityInfo security,
        PdfSignatureInfo signature,
        bool completeShape,
        bool ordered,
        List<PdfSignatureValidationFinding> findings,
        IPdfSignatureCryptographyProvider provider) {
        if (!completeShape || !ordered || signature.ContentsBytes is null || signature.ContentsBytes.Length == 0) {
            AddSignatureFinding(
                findings,
                signature,
                PdfDiagnosticSeverity.Error,
                "CryptographicValidationInputUnavailable",
                "Cryptographic validation requires a valid four-value /ByteRange and non-empty decoded /Contents bytes.",
                isCryptographic: true);
            return null;
        }

        try {
            byte[] signedContent = ReadSignedContent(pdf, signature.ByteRangeValues);
            var input = new PdfSignatureCryptographyInput(
                signature,
                signedContent,
                (byte[])signature.ContentsBytes.Clone(),
                pdf.LongLength,
                security.DocumentSecurityStore);
            PdfSignatureCryptographicResult result = provider.Verify(input) ??
                throw new InvalidOperationException("Signature cryptography provider returned no result.");
            for (int i = 0; i < result.Findings.Count; i++) {
                PdfSignatureCryptographicFinding finding = result.Findings[i];
                AddSignatureFinding(findings, signature, finding.Severity, finding.Code, finding.Message, isCryptographic: true);
            }

            return result;
        } catch (Exception ex) when (ex is not OutOfMemoryException && ex is not StackOverflowException) {
            AddSignatureFinding(
                findings,
                signature,
                PdfDiagnosticSeverity.Error,
                "CryptographicProviderFailed",
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

    private static byte[] ReadSignedContent(byte[] pdf, IReadOnlyList<long> byteRangeValues) {
        long totalLength = byteRangeValues[1] + byteRangeValues[3];
        if (totalLength > int.MaxValue) {
            throw new InvalidOperationException("Signed byte ranges exceed the current in-memory validation limit.");
        }

        var content = new byte[(int)totalLength];
        int destinationOffset = 0;
        for (int i = 0; i < byteRangeValues.Count; i += 2) {
            int sourceOffset = checked((int)byteRangeValues[i]);
            int count = checked((int)byteRangeValues[i + 1]);
            Buffer.BlockCopy(pdf, sourceOffset, content, destinationOffset, count);
            destinationOffset += count;
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
                "CryptographicTrustNotVerified",
                "Signature structure was inspected without an optional cryptography provider; CMS math, trust, revocation, digest, and timestamp policy were not evaluated."));
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
}

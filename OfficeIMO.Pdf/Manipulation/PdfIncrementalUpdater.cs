using OfficeIMO.Drawing.Internal;
namespace OfficeIMO.Pdf;

/// <summary>
/// Append-only PDF update helpers for changes that can be represented as a new incremental revision.
/// </summary>
internal static partial class PdfIncrementalUpdater {
    /// <summary>
    /// Analyzes append-only mutation support for a PDF byte array.
    /// </summary>
    public static PdfAppendOnlyMutationReport AnalyzeAppendOnlyMutation(byte[] pdf) {
        return AnalyzeAppendOnlyMutation(pdf, null);
    }

    /// <summary>Analyzes append-only mutation support using optional password and parsing settings.</summary>
    public static PdfAppendOnlyMutationReport AnalyzeAppendOnlyMutation(byte[] pdf, PdfReadOptions? readOptions) {
        Guard.NotNull(pdf, nameof(pdf));
        return AnalyzeAppendOnlyMutation(PdfSyntax.ReadDocumentSecurityInfo(pdf, readOptions));
    }

    /// <summary>
    /// Analyzes append-only mutation support from already-read PDF security and revision markers.
    /// </summary>
    public static PdfAppendOnlyMutationReport AnalyzeAppendOnlyMutation(PdfDocumentSecurityInfo security) {
        Guard.NotNull(security, nameof(security));
        return BuildAppendOnlyMutationReport(security, fieldNames: null);
    }

    internal static PdfAppendOnlyMutationReport AnalyzeAppendOnlyMutation(PdfDocumentSecurityInfo security, IEnumerable<string>? fieldNames) {
        Guard.NotNull(security, nameof(security));
        return BuildAppendOnlyMutationReport(security, fieldNames);
    }

    /// <summary>Analyzes append-only mutation support for a readable PDF stream.</summary>
    public static PdfAppendOnlyMutationReport AnalyzeAppendOnlyMutation(Stream input) {
        Guard.NotNull(input, nameof(input));
        if (!input.CanRead) {
            throw new ArgumentException("Stream must be readable.", nameof(input));
        }

        using var buffer = new MemoryStream();
        input.CopyTo(buffer);
        return AnalyzeAppendOnlyMutation(buffer.ToArray());
    }

    /// <summary>Analyzes append-only mutation support for a PDF file.</summary>
    public static PdfAppendOnlyMutationReport AnalyzeAppendOnlyMutation(string inputPath) {
        Guard.NotNullOrWhiteSpace(inputPath, nameof(inputPath));
        return AnalyzeAppendOnlyMutation(File.ReadAllBytes(inputPath));
    }

    /// <summary>
    /// Appends a metadata-only revision to a PDF byte array without rewriting the existing bytes.
    /// </summary>
    public static byte[] UpdateMetadata(
        byte[] pdf,
        string? title = null,
        string? author = null,
        string? subject = null,
        string? keywords = null,
        PdfReadOptions? readOptions = null,
        bool createXmpMetadata = false) {
        Guard.NotNull(pdf, nameof(pdf));
        _ = PdfMutationPlanner.RequireAppendOnly(pdf, PdfMutationOperation.UpdateMetadata, readOptions);

        PdfDocumentSecurityInfo security = PdfSyntax.ReadDocumentSecurityInfo(pdf, readOptions);

        var (objects, trailerRaw) = PdfSyntax.ParseObjects(pdf, readOptions);
        if (!security.RootObjectNumber.HasValue) {
            throw new InvalidOperationException("PDF root catalog reference is required for an incremental metadata update.");
        }

        if (!security.LastStartXrefOffset.HasValue) {
            throw new InvalidOperationException("PDF startxref offset is required for an incremental metadata update.");
        }

        PdfDocumentInfo documentInfo = PdfInspector.Inspect(pdf, readOptions);
        PdfMetadata existing = documentInfo.Metadata;
        PdfXmpMetadataInfo? existingXmp = documentInfo.XmpMetadata;
        var updated = new PdfMetadata {
            Title = title ?? existing.Title ?? existingXmp?.Title,
            Author = author ?? existing.Author ?? existingXmp?.Creator,
            Subject = subject ?? existing.Subject ?? existingXmp?.Description,
            Keywords = keywords ?? existing.Keywords ?? existingXmp?.Keywords
        };

        int newInfoObjectNumber = objects.Count == 0 ? 1 : objects.Keys.Max() + 1;
        objects[newInfoObjectNumber] = new PdfIndirectObject(newInfoObjectNumber, 0, PdfInfoDictionaryBuilder.BuildDictionary(updated));
        var changedObjectNumbers = new List<int> { newInfoObjectNumber };
        SynchronizeXmpMetadata(objects, security, existingXmp, updated, createXmpMetadata, changedObjectNumbers);
        PdfStandardSecurityHandler? encryptionHandler = null;
        if (security.HasEncryption &&
            !PdfSyntax.TryCreateDecryptor(objects, trailerRaw, readOptions, out encryptionHandler)) {
            throw new PdfUnsupportedEncryptionException("PDF encryption context could not be created for the incremental update.");
        }

        return PdfIncrementalObjectWriter.Append(
            pdf,
            objects,
            security,
            trailerRaw,
            changedObjectNumbers: changedObjectNumbers,
            infoObjectNumberOverride: newInfoObjectNumber,
            encryptionHandler: encryptionHandler);
    }

    /// <summary>Appends a metadata-only revision to a PDF stream.</summary>
    public static byte[] UpdateMetadata(
        Stream input,
        string? title = null,
        string? author = null,
        string? subject = null,
        string? keywords = null,
        PdfReadOptions? readOptions = null,
        bool createXmpMetadata = false) {
        Guard.NotNull(input, nameof(input));
        if (!input.CanRead) {
            throw new ArgumentException("Stream must be readable.", nameof(input));
        }

        using var buffer = new MemoryStream();
        input.CopyTo(buffer);
        return UpdateMetadata(buffer.ToArray(), title, author, subject, keywords, readOptions, createXmpMetadata);
    }

    /// <summary>Appends a metadata-only revision to a PDF file and writes the result to <paramref name="outputPath"/>.</summary>
    public static void UpdateMetadata(
        string inputPath,
        string outputPath,
        string? title = null,
        string? author = null,
        string? subject = null,
        string? keywords = null,
        PdfReadOptions? readOptions = null,
        bool createXmpMetadata = false) {
        Guard.NotNullOrWhiteSpace(inputPath, nameof(inputPath));
        Guard.NotNullOrWhiteSpace(outputPath, nameof(outputPath));
        OfficeFileCommit.WriteAllBytes(outputPath, UpdateMetadata(File.ReadAllBytes(inputPath), title, author, subject, keywords, readOptions, createXmpMetadata));
    }

    private static void SynchronizeXmpMetadata(
        Dictionary<int, PdfIndirectObject> objects,
        PdfDocumentSecurityInfo security,
        PdfXmpMetadataInfo? existingXmp,
        PdfMetadata updated,
        bool createXmpMetadata,
        List<int> changedObjectNumbers) {
        if (!security.RootObjectNumber.HasValue ||
            !objects.TryGetValue(security.RootObjectNumber.Value, out PdfIndirectObject? catalogObject) ||
            catalogObject.Value is not PdfDictionary catalog) {
            throw new InvalidOperationException("PDF root catalog dictionary is required for XMP metadata synchronization.");
        }

        PdfReference? metadataReference = catalog.Get<PdfReference>("Metadata");
        if (metadataReference is null && !createXmpMetadata) {
            return;
        }

        byte[] xml;
        PdfDictionary streamDictionary;
        int metadataObjectNumber;
        int metadataGeneration;
        if (metadataReference is not null &&
            PdfObjectLookup.TryGet(objects, metadataReference, out PdfIndirectObject? metadataObject) &&
            metadataObject.Value is PdfStream metadataStream) {
            if (existingXmp is null || existingXmp.RawXml is null || !existingXmp.IsWellFormedXml || existingXmp.HasUnsupportedFilters) {
                throw new InvalidOperationException("The existing XMP metadata stream cannot be decoded and preserved safely.");
            }

            xml = PdfXmpMetadataSynchronizer.Synchronize(existingXmp.RawXml, updated);
            streamDictionary = PdfXmpMetadataSynchronizer.CloneUnfilteredMetadataDictionary(metadataStream.Dictionary);
            metadataObjectNumber = metadataObject.ObjectNumber;
            metadataGeneration = metadataObject.Generation;
        } else {
            xml = PdfXmpMetadataBuilder.Build(updated.Title, updated.Author, updated.Subject, updated.Keywords);
            streamDictionary = new PdfDictionary();
            streamDictionary.Items["Type"] = new PdfName("Metadata");
            streamDictionary.Items["Subtype"] = new PdfName("XML");
            metadataObjectNumber = objects.Count == 0 ? 1 : objects.Keys.Max() + 1;
            metadataGeneration = 0;
            catalog.Items["Metadata"] = new PdfReference(metadataObjectNumber, metadataGeneration);
            changedObjectNumbers.Add(catalogObject.ObjectNumber);
        }

        objects[metadataObjectNumber] = new PdfIndirectObject(
            metadataObjectNumber,
            metadataGeneration,
            new PdfStream(streamDictionary, xml));
        changedObjectNumbers.Add(metadataObjectNumber);
    }

    private static PdfAppendOnlyMutationReport BuildAppendOnlyMutationReport(PdfDocumentSecurityInfo security, IEnumerable<string>? fieldNames) {
        var commonBlockers = new List<string>();
        var metadataBlockers = new List<string>();
        var formBlockers = new List<string>();
        var signaturePreparationBlockers = new List<string>();
        var longTermValidationBlockers = new List<string>();
        var annotationBlockers = new List<string>();
        var warnings = new List<string>();
        if (security.HasEncryption && !security.HasOwnerAuthorization) {
            commonBlockers.Add("Encrypted");
        }

        bool hasSignatureContent = security.SignatureFieldCount > 0 || security.SignatureCount > 0 || security.HasByteRange;
        if (security.HasUsageRights) {
            commonBlockers.Add("UsageRights");
        }

        if (!security.RootObjectNumber.HasValue) {
            commonBlockers.Add("MissingRoot");
        }

        if (!security.LastStartXrefOffset.HasValue) {
            commonBlockers.Add("MissingStartXref");
        }

        metadataBlockers.AddRange(commonBlockers);
        formBlockers.AddRange(commonBlockers);
        signaturePreparationBlockers.AddRange(commonBlockers);
        longTermValidationBlockers.AddRange(commonBlockers);
        annotationBlockers.AddRange(commonBlockers);
        if (security.HasEncryption) {
            signaturePreparationBlockers.Add("EncryptedRawSignatureObject");
        }
        bool blockedBySignatureFieldLock = HasBlockingSignatureFieldLock(security, fieldNames);

        if (hasSignatureContent) {
            metadataBlockers.Add("Signed");
            if (!CanAppendFormFieldsWithDocMDP(security, null)) {
                formBlockers.Add("Signed");
            } else {
                warnings.Add("SignedDocMDPFormFill");
            }

            if (!security.HasDocMDPPermissions) {
                warnings.Add("SignedApprovalAnnotationChange");
            }
        }

        if (!hasSignatureContent) {
            longTermValidationBlockers.Add("Unsigned");
        }

        if (security.HasDocMDPPermissions) {
            metadataBlockers.Add("DocMDP");
            if (!CanAppendFormFieldsWithDocMDP(security, null)) {
                formBlockers.Add("DocMDP");
            } else {
                warnings.Add("DocMDPAllowsFormFill");
            }

            warnings.Add("DocMDPDssMaintenance");
            if (security.DocMDPPermissionLevel == 3) {
                warnings.Add("DocMDPAllowsAnnotations");
            } else {
                annotationBlockers.Add("DocMDP");
            }
        }

        if (blockedBySignatureFieldLock) {
            formBlockers.Add("SignatureFieldLock");
        }

        if (security.HasIncrementalUpdates) {
            warnings.Add("ExistingIncrementalRevisions");
        }

        if (security.AcroFormAppendOnly) {
            warnings.Add("AcroFormAppendOnly");
        }

        var supported = new List<string>();
        if (metadataBlockers.Count == 0) {
            supported.Add("Metadata");
        }

        if (formBlockers.Count == 0) {
            supported.Add("FormFill");
        }

        if (longTermValidationBlockers.Count == 0) {
            supported.Add("LongTermValidation");
        }

        if (annotationBlockers.Count == 0) {
            supported.Add("Annotations");
        }

        bool canPrepareSignature =
            signaturePreparationBlockers.Count == 0 &&
            !hasSignatureContent &&
            !security.HasDocMDPPermissions;
        if (canPrepareSignature) {
            supported.Add("SignaturePrepare");
        }

        var blocked = new List<string>();
        if (metadataBlockers.Count > 0) {
            blocked.Add("Metadata");
        }

        if (formBlockers.Count > 0) {
            blocked.Add("FormFill");
        }

        if (!canPrepareSignature) {
            blocked.Add("SignaturePrepare");
        }

        if (longTermValidationBlockers.Count > 0) {
            blocked.Add("LongTermValidation");
        }

        if (annotationBlockers.Count > 0) {
            blocked.Add("Annotations");
        }

        blocked.Add("PageTree");
        blocked.Add("Attachments");

        var blockers = metadataBlockers
            .Concat(formBlockers)
            .Distinct(StringComparer.Ordinal)
            .ToArray();

        return new PdfAppendOnlyMutationReport(
            security,
            supported.AsReadOnly(),
            blocked.AsReadOnly(),
            blockers,
            warnings.Distinct(StringComparer.Ordinal).ToArray());
    }

    private static bool CanAppendFormFieldsWithDocMDP(PdfDocumentSecurityInfo security, IEnumerable<string>? fieldNames) {
        if (!security.HasDocMDPPermissions ||
            !security.DocMDPPermissionLevel.HasValue ||
            security.DocMDPPermissionLevel.Value < 2 ||
            security.DocMDPPermissionLevel.Value > 3) {
            return false;
        }

        if (fieldNames is null) {
            return !security.Signatures.Any(static signature => LocksEveryField(signature.FieldLock));
        }

        return GetFirstLockedFormFieldName(security, fieldNames) is null;
    }

    private static bool HasBlockingSignatureFieldLock(PdfDocumentSecurityInfo security, IEnumerable<string>? fieldNames) {
        HashSet<string>? requestedFields = fieldNames is null
            ? null
            : new HashSet<string>(fieldNames.Where(static field => !string.IsNullOrWhiteSpace(field)), StringComparer.Ordinal);
        foreach (PdfSignatureInfo signature in security.Signatures) {
            PdfSignatureFieldLockInfo? fieldLock = signature.FieldLock;
            if (fieldLock is null) {
                continue;
            }

            if (requestedFields is null) {
                return true;
            }

            if (fieldLock.LocksAllFields) {
                return true;
            }

            if (fieldLock.LocksIncludedFields &&
                fieldLock.Fields.Any(lockedField => requestedFields.Any(requestedField => FieldNamesOverlap(lockedField, requestedField)))) {
                return true;
            }

            if (fieldLock.LocksAllExceptListedFields &&
                requestedFields.Any(requestedField => !fieldLock.Fields.Any(excludedField => FieldNamesOverlap(excludedField, requestedField)))) {
                return true;
            }
        }

        return false;
    }

    private static bool FieldNamesOverlap(string lockFieldName, string requestedFieldName) {
        if (string.IsNullOrWhiteSpace(lockFieldName) ||
            string.IsNullOrWhiteSpace(requestedFieldName)) {
            return false;
        }

        return string.Equals(lockFieldName, requestedFieldName, StringComparison.Ordinal) ||
            requestedFieldName.StartsWith(lockFieldName + ".", StringComparison.Ordinal) ||
            lockFieldName.StartsWith(requestedFieldName + ".", StringComparison.Ordinal);
    }

    private static string? GetFirstLockedFormFieldName(PdfDocumentSecurityInfo security, IEnumerable<string> fieldNames) {
        var requested = new HashSet<string>(fieldNames, StringComparer.Ordinal);
        if (requested.Count == 0) {
            return null;
        }

        foreach (PdfSignatureInfo signature in security.Signatures) {
            PdfSignatureFieldLockInfo? fieldLock = signature.FieldLock;
            if (fieldLock is null) {
                continue;
            }

            foreach (string fieldName in requested) {
                if (IsFieldLocked(fieldLock, fieldName)) {
                    return fieldName;
                }
            }
        }

        return null;
    }

    private static bool IsFieldLocked(PdfSignatureFieldLockInfo fieldLock, string fieldName) {
        if (fieldLock.LocksAllFields) {
            return true;
        }

        bool listed = fieldLock.Fields.Any(lockedField => FieldNamesOverlap(lockedField, fieldName));
        if (fieldLock.LocksIncludedFields) {
            return listed;
        }

        if (fieldLock.LocksAllExceptListedFields) {
            return !listed;
        }

        return false;
    }

    private static bool LocksEveryField(PdfSignatureFieldLockInfo? fieldLock) {
        return fieldLock is not null &&
            (fieldLock.LocksAllFields ||
            (fieldLock.LocksAllExceptListedFields && fieldLock.Fields.Count == 0));
    }

}

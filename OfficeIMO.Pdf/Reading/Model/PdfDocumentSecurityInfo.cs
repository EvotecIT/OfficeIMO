namespace OfficeIMO.Pdf;

/// <summary>
/// Lightweight security, signature, and revision markers read from a PDF file.
/// </summary>
public sealed class PdfDocumentSecurityInfo {
    private const int PrintPermissionBit = 4;
    private const int ModifyPermissionBit = 8;
    private const int CopyPermissionBit = 16;
    private const int AnnotatePermissionBit = 32;
    private const int FillFormsPermissionBit = 256;
    private const int AccessibilityPermissionBit = 512;
    private const int AssemblePermissionBit = 1024;
    private const int HighQualityPrintPermissionBit = 2048;
    private const int AcroFormSignaturesExistFlag = 1;
    private const int AcroFormAppendOnlyFlag = 2;

    internal PdfDocumentSecurityInfo(
        bool hasEncryption,
        int? encryptObjectNumber,
        string? encryptionFilter,
        string? encryptionSubFilter,
        int? encryptionVersion,
        int? encryptionRevision,
        int? encryptionLengthBits,
        int? encryptionPermissions,
        bool? encryptMetadata,
        PdfPasswordAuthenticationRole passwordAuthenticationRole,
        bool hasSignatures,
        IReadOnlyList<int> signatureFieldObjectNumbers,
        IReadOnlyList<string> signatureFieldNames,
        IReadOnlyList<PdfSignatureInfo> signatures,
        int signatureValueCount,
        bool hasByteRange,
        int byteRangeValueCount,
        int? acroFormSignatureFlags,
        bool hasDocMDPPermissions,
        int? docMDPSignatureObjectNumber,
        string? docMDPTransformMethod,
        string? docMDPTransformVersion,
        int? docMDPPermissionLevel,
        bool hasUsageRights,
        IReadOnlyList<int> usageRightsObjectNumbers,
        PdfDocumentDssInfo documentSecurityStore,
        int? rootObjectNumber,
        int? rootObjectGeneration,
        int? infoObjectNumber,
        int? infoObjectGeneration,
        bool hasTrailerId,
        int startXrefCount,
        int? lastStartXrefOffset,
        IReadOnlyList<int> startXrefOffsets,
        IReadOnlyList<int> previousXrefOffsets,
        IReadOnlyList<PdfDocumentRevisionInfo> revisions,
        bool hasPreviousRevision,
        bool hasXrefStreams,
        bool hasObjectStreams) {
        HasEncryption = hasEncryption;
        EncryptObjectNumber = encryptObjectNumber;
        EncryptionFilter = encryptionFilter;
        EncryptionSubFilter = encryptionSubFilter;
        EncryptionVersion = encryptionVersion;
        EncryptionRevision = encryptionRevision;
        EncryptionLengthBits = encryptionLengthBits;
        EncryptionPermissions = encryptionPermissions;
        EncryptMetadata = encryptMetadata;
        PasswordAuthenticationRole = passwordAuthenticationRole;
        HasSignatures = hasSignatures;
        SignatureFieldObjectNumbers = signatureFieldObjectNumbers;
        SignatureFieldNames = signatureFieldNames;
        Signatures = signatures;
        SignatureValueCount = signatureValueCount;
        HasByteRange = hasByteRange;
        ByteRangeValueCount = byteRangeValueCount;
        AcroFormSignatureFlags = acroFormSignatureFlags;
        HasDocMDPPermissions = hasDocMDPPermissions;
        DocMDPSignatureObjectNumber = docMDPSignatureObjectNumber;
        DocMDPTransformMethod = docMDPTransformMethod;
        DocMDPTransformVersion = docMDPTransformVersion;
        DocMDPPermissionLevel = docMDPPermissionLevel;
        HasUsageRights = hasUsageRights;
        UsageRightsObjectNumbers = usageRightsObjectNumbers;
        DocumentSecurityStore = documentSecurityStore;
        RootObjectNumber = rootObjectNumber;
        RootObjectGeneration = rootObjectGeneration;
        InfoObjectNumber = infoObjectNumber;
        InfoObjectGeneration = infoObjectGeneration;
        HasTrailerId = hasTrailerId;
        StartXrefCount = startXrefCount;
        LastStartXrefOffset = lastStartXrefOffset;
        StartXrefOffsets = startXrefOffsets;
        PreviousXrefOffsets = previousXrefOffsets;
        Revisions = revisions;
        HasPreviousRevision = hasPreviousRevision;
        HasXrefStreams = hasXrefStreams;
        HasObjectStreams = hasObjectStreams;
    }

    /// <summary>True when the file contains an /Encrypt marker.</summary>
    public bool HasEncryption { get; }

    /// <summary>Encryption dictionary object number when the trailer points to an indirect dictionary.</summary>
    public int? EncryptObjectNumber { get; }

    /// <summary>Encryption /Filter name, for example Standard, when readable.</summary>
    public string? EncryptionFilter { get; }

    /// <summary>Encryption /SubFilter name, when readable.</summary>
    public string? EncryptionSubFilter { get; }

    /// <summary>Raw encryption algorithm version from /V, when readable.</summary>
    public int? EncryptionVersion { get; }

    /// <summary>Raw standard-security-handler revision from /R, when readable.</summary>
    public int? EncryptionRevision { get; }

    /// <summary>Encryption key length in bits from /Length, when readable.</summary>
    public int? EncryptionLengthBits { get; }

    /// <summary>Raw standard-security-handler permission bits from /P, when readable.</summary>
    public int? EncryptionPermissions { get; }

    /// <summary>Typed Standard security permissions, when a raw `/P` mask is present.</summary>
    public PdfStandardPermissions? AllowedStandardPermissions =>
        EncryptionPermissions.HasValue
            ? PdfStandardEncryptionOptions.FromRawPermissions(EncryptionPermissions.Value)
            : (PdfStandardPermissions?)null;

    /// <summary>Encryption /EncryptMetadata flag, when readable.</summary>
    public bool? EncryptMetadata { get; }

    /// <summary>Role established by the supplied Standard-security password, or <see cref="PdfPasswordAuthenticationRole.None"/> when no authentication was performed.</summary>
    public PdfPasswordAuthenticationRole PasswordAuthenticationRole { get; }

    /// <summary>True when the supplied password authenticated as the Standard-security owner password.</summary>
    public bool HasOwnerAuthorization => PasswordAuthenticationRole == PdfPasswordAuthenticationRole.Owner;

    /// <summary>True when the encryption dictionary exposed at least one readable setting.</summary>
    public bool HasReadableEncryptionSettings =>
        EncryptObjectNumber.HasValue ||
        !string.IsNullOrEmpty(EncryptionFilter) ||
        !string.IsNullOrEmpty(EncryptionSubFilter) ||
        EncryptionVersion.HasValue ||
        EncryptionRevision.HasValue ||
        EncryptionLengthBits.HasValue ||
        EncryptionPermissions.HasValue ||
        EncryptMetadata.HasValue;

    /// <summary>True when the raw permissions allow low-resolution printing.</summary>
    public bool? AllowsPrinting => ReadPermission(PrintPermissionBit);

    /// <summary>True when the raw permissions allow document modification.</summary>
    public bool? AllowsModification => ReadPermission(ModifyPermissionBit);

    /// <summary>True when the raw permissions allow content copying or extraction.</summary>
    public bool? AllowsCopying => ReadPermission(CopyPermissionBit);

    /// <summary>True when the raw permissions allow comments or form-field annotation changes.</summary>
    public bool? AllowsAnnotationChanges => ReadPermission(AnnotatePermissionBit);

    /// <summary>True when the raw permissions allow filling existing form fields.</summary>
    public bool? AllowsFormFilling => ReadPermission(FillFormsPermissionBit);

    /// <summary>True when the raw permissions allow accessibility extraction.</summary>
    public bool? AllowsAccessibilityExtraction => ReadPermission(AccessibilityPermissionBit);

    /// <summary>True when the raw permissions allow document assembly.</summary>
    public bool? AllowsDocumentAssembly => ReadPermission(AssemblePermissionBit);

    /// <summary>True when the raw permissions allow high-quality printing.</summary>
    public bool? AllowsHighQualityPrinting => ReadPermission(HighQualityPrintPermissionBit);

    /// <summary>True when signature markers, signature fields, or signature values were found.</summary>
    public bool HasSignatures { get; }

    /// <summary>Object numbers for AcroForm fields whose /FT is /Sig.</summary>
    public IReadOnlyList<int> SignatureFieldObjectNumbers { get; }

    /// <summary>Readable names for AcroForm signature fields.</summary>
    public IReadOnlyList<string> SignatureFieldNames { get; }

    /// <summary>Readable signature value dictionaries and their owning AcroForm fields.</summary>
    public IReadOnlyList<PdfSignatureInfo> Signatures { get; }

    /// <summary>Number of AcroForm fields whose /FT is /Sig.</summary>
    public int SignatureFieldCount => SignatureFieldObjectNumbers.Count;

    /// <summary>Number of readable signature value dictionaries.</summary>
    public int SignatureCount => Signatures.Count;

    /// <summary>Number of signature value dictionaries discovered from /Type /Sig or /ByteRange.</summary>
    public int SignatureValueCount { get; }

    /// <summary>True when a signature /ByteRange array was found.</summary>
    public bool HasByteRange { get; }

    /// <summary>Number of numeric values found in signature /ByteRange arrays.</summary>
    public int ByteRangeValueCount { get; }

    /// <summary>Number of byte ranges represented by the numeric /ByteRange values.</summary>
    public int ByteRangeSegmentCount => ByteRangeValueCount / 2;

    /// <summary>Raw AcroForm /SigFlags value, when readable.</summary>
    public int? AcroFormSignatureFlags { get; }

    /// <summary>True when AcroForm /SigFlags says signatures exist.</summary>
    public bool AcroFormSignaturesExist => HasAcroFormSignatureFlag(AcroFormSignaturesExistFlag);

    /// <summary>True when AcroForm /SigFlags says changes should be appended.</summary>
    public bool AcroFormAppendOnly => HasAcroFormSignatureFlag(AcroFormAppendOnlyFlag);

    /// <summary>True when catalog /Perms exposes DocMDP permissions.</summary>
    public bool HasDocMDPPermissions { get; }

    /// <summary>Object number referenced by catalog /Perms /DocMDP, when readable.</summary>
    public int? DocMDPSignatureObjectNumber { get; }

    /// <summary>DocMDP signature reference /TransformMethod name, when readable.</summary>
    public string? DocMDPTransformMethod { get; }

    /// <summary>DocMDP /TransformParams /V version name or string, when readable.</summary>
    public string? DocMDPTransformVersion { get; }

    /// <summary>DocMDP /TransformParams /P permission level, when readable.</summary>
    public int? DocMDPPermissionLevel { get; }

    /// <summary>True when catalog /Perms exposes usage-rights entries such as /UR or /UR3.</summary>
    public bool HasUsageRights { get; }

    /// <summary>Object numbers referenced by catalog usage-rights entries such as /UR or /UR3.</summary>
    public IReadOnlyList<int> UsageRightsObjectNumbers { get; }

    /// <summary>Document Security Store (/DSS) evidence used for long-term signature validation, when present.</summary>
    public PdfDocumentDssInfo DocumentSecurityStore { get; }

    /// <summary>True when the catalog exposes a /DSS dictionary.</summary>
    public bool HasDocumentSecurityStore => DocumentSecurityStore.HasDss;

    /// <summary>True when the /DSS dictionary exposes validation evidence references or VRI entries.</summary>
    public bool HasLongTermValidationEvidence => DocumentSecurityStore.HasValidationEvidence;

    /// <summary>Trailer root catalog object number, when readable.</summary>
    public int? RootObjectNumber { get; }

    /// <summary>Trailer root catalog generation, when readable.</summary>
    public int? RootObjectGeneration { get; }

    /// <summary>Trailer info dictionary object number, when readable.</summary>
    public int? InfoObjectNumber { get; }

    /// <summary>Trailer info dictionary generation, when readable.</summary>
    public int? InfoObjectGeneration { get; }

    /// <summary>True when a trailer /ID entry was found.</summary>
    public bool HasTrailerId { get; }

    /// <summary>Number of startxref sections found in the file.</summary>
    public int StartXrefCount { get; }

    /// <summary>Offset from the last startxref section, when readable.</summary>
    public int? LastStartXrefOffset { get; }

    /// <summary>All readable startxref offsets in file order.</summary>
    public IReadOnlyList<int> StartXrefOffsets { get; }

    /// <summary>All readable /Prev offsets in file order.</summary>
    public IReadOnlyList<int> PreviousXrefOffsets { get; }

    /// <summary>Readable cross-reference revision markers in file order.</summary>
    public IReadOnlyList<PdfDocumentRevisionInfo> Revisions { get; }

    /// <summary>Number of readable cross-reference revision markers.</summary>
    public int RevisionCount => Revisions.Count;

    /// <summary>True when multiple revisions or /Prev links were found.</summary>
    public bool HasIncrementalUpdates => StartXrefCount > 1 || HasPreviousRevision;

    /// <summary>True when a trailer or xref stream points to a previous revision.</summary>
    public bool HasPreviousRevision { get; }

    /// <summary>True when xref stream markers were found.</summary>
    public bool HasXrefStreams { get; }

    /// <summary>True when object stream markers were found.</summary>
    public bool HasObjectStreams { get; }

    /// <summary>True when mutation must preserve the existing file by appending a new revision instead of rewriting bytes in place.</summary>
    public bool RequiresAppendOnlyMutation =>
        HasSignatures ||
        AcroFormAppendOnly ||
        HasDocMDPPermissions ||
        HasUsageRights ||
        HasIncrementalUpdates;

    /// <summary>True when the current OfficeIMO.Pdf writer cannot safely attempt append-only mutation for this input yet.</summary>
    public bool BlocksOfficeIMOAppendOnlyMutation =>
        HasEncryption ||
        HasSignatures ||
        HasDocMDPPermissions ||
        HasUsageRights;

    /// <summary>True when OfficeIMO.Pdf should avoid safe full-rewrite mutation for this input.</summary>
    public bool BlocksOfficeIMOFullRewriteMutation =>
        HasEncryption ||
        HasSignatures ||
        HasDocMDPPermissions ||
        HasUsageRights ||
        HasXrefStreams ||
        HasObjectStreams;

    private bool? ReadPermission(int bit) {
        return EncryptionPermissions.HasValue ? (EncryptionPermissions.Value & bit) != 0 : (bool?)null;
    }

    private bool HasAcroFormSignatureFlag(int flag) {
        return AcroFormSignatureFlags.HasValue && (AcroFormSignatureFlags.Value & flag) == flag;
    }
}

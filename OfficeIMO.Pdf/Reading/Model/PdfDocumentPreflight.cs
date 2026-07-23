namespace OfficeIMO.Pdf;

/// <summary>
/// Wrapper-friendly PDF capability report for OfficeIMO.Pdf read and rewrite operations.
/// </summary>
public sealed partial class PdfDocumentPreflight {
    private readonly PdfDocumentInfo? _documentInfo;

    internal PdfDocumentPreflight(
        PdfDocumentProbe probe,
        PdfDocumentInfo? documentInfo,
        bool canRead,
        bool canRewrite,
        IReadOnlyList<string> diagnostics,
        IReadOnlyList<PdfReadBlocker> readBlockers,
        IReadOnlyList<PdfRewriteBlocker> rewriteBlockers,
        PdfPermissionPolicy permissionPolicy) {
        Probe = probe;
        _documentInfo = documentInfo;
        CanRead = canRead;
        CanRewrite = canRewrite;
        Diagnostics = diagnostics;
        ReadBlockers = readBlockers;
        RewriteBlockers = rewriteBlockers;
        PermissionPolicy = permissionPolicy;
    }

    /// <summary>Lightweight PDF markers read before full parsing.</summary>
    public PdfDocumentProbe Probe { get; }

    /// <summary>Parsed document information when content extraction is authorized; otherwise <see langword="null"/>.</summary>
    public PdfDocumentInfo? DocumentInfo => PdfPermissionAuthorization.CanExtractContent(Probe.Security, PermissionPolicy)
        ? _documentInfo
        : null;

    /// <summary>Parsed document information retained for internal capability and mutation planning.</summary>
    internal PdfDocumentInfo? UncheckedDocumentInfo => _documentInfo;

    /// <summary>True when OfficeIMO.Pdf can parse enough of the document for read-oriented operations.</summary>
    public bool CanRead { get; }

    /// <summary>True when OfficeIMO.Pdf can attempt rewrite-style manipulation without known security blockers.</summary>
    public bool CanRewrite { get; }

    /// <summary>True when OfficeIMO.Pdf can attempt text and structured text readback operations for this PDF.</summary>
    public bool CanExtractText => CanRead && PdfPermissionAuthorization.CanExtractText(Probe.Security, PermissionPolicy);

    /// <summary>True when OfficeIMO.Pdf can attempt image XObject extraction for this PDF.</summary>
    public bool CanExtractImages => _documentInfo is not null && !HasImageExtractionBlocker() && PdfPermissionAuthorization.CanExtractContent(Probe.Security, PermissionPolicy);

    /// <summary>True when OfficeIMO.Pdf can attempt embedded-file and associated-file attachment extraction for this PDF.</summary>
    public bool CanExtractAttachments => _documentInfo is not null && !HasAttachmentExtractionBlocker() && PdfPermissionAuthorization.CanExtractContent(Probe.Security, PermissionPolicy);

    /// <summary>True when OfficeIMO.Pdf can attempt logical object readback through PdfLogicalDocument for this PDF.</summary>
    public bool CanReadLogicalObjects => CanRead && PdfPermissionAuthorization.CanExtractContent(Probe.Security, PermissionPolicy);

    /// <summary>Permission policy used while evaluating extraction and mutation capabilities.</summary>
    public PdfPermissionPolicy PermissionPolicy { get; }

    /// <summary>True when authenticated user-password restrictions are being explicitly ignored.</summary>
    public bool PermissionRestrictionsIgnored => PdfPermissionAuthorization.RestrictionsIgnored(Probe.Security, PermissionPolicy);

    /// <summary>True when OfficeIMO.Pdf can attempt at least one page-level rewrite operation.</summary>
    public bool CanManipulatePages => CanRewrite || CanUseAuthenticatedEncryptedPageMutation();

    /// <summary>True when OfficeIMO.Pdf can attempt simple AcroForm value updates for named text, choice, or button fields.</summary>
    public bool CanFillSimpleFormFields => CanRead && !HasFormMutationBlocker(PdfMutationOperation.FillFormFields) && HasSimpleFillableFormFields();

    /// <summary>True when OfficeIMO.Pdf can attempt simple AcroForm flattening for text, choice, or button widgets with page-backed rectangles.</summary>
    public bool CanFlattenSimpleFormFields => CanRead && !HasFormMutationBlocker(PdfMutationOperation.FlattenFormFields) && HasSimpleFlattenableFormFields();

    /// <summary>True when OfficeIMO.Pdf can attempt simple AcroForm value updates followed by simple widget flattening.</summary>
    public bool CanFillAndFlattenSimpleFormFields => CanFillSimpleFormFields && CanFlattenSimpleFormFields;

    /// <summary>Human-readable diagnostics explaining blocked or risky operations.</summary>
    public IReadOnlyList<string> Diagnostics { get; }

    /// <summary>Structured reasons why read-oriented operations are blocked.</summary>
    public IReadOnlyList<PdfReadBlocker> ReadBlockers { get; }

    /// <summary>Structured reasons why rewrite-style manipulation is blocked.</summary>
    public IReadOnlyList<PdfRewriteBlocker> RewriteBlockers { get; }

    /// <summary>Returns true when a specific read blocker is present.</summary>
    public bool HasReadBlocker(PdfReadBlockerKind kind) {
        for (int i = 0; i < ReadBlockers.Count; i++) {
            if (ReadBlockers[i].Kind == kind) {
                return true;
            }
        }

        return false;
    }

    /// <summary>Returns true when a specific rewrite blocker is present.</summary>
    public bool HasRewriteBlocker(PdfRewriteBlockerKind kind) {
        for (int i = 0; i < RewriteBlockers.Count; i++) {
            if (RewriteBlockers[i].Kind == kind) {
                return true;
            }
        }

        return false;
    }

    private bool CanUseAuthenticatedEncryptedPageMutation() {
        if (!CanRead || !Probe.Security.HasEncryption) {
            return false;
        }

        for (int i = 0; i < RewriteBlockers.Count; i++) {
            if (RewriteBlockers[i].Kind != PdfRewriteBlockerKind.Encryption) {
                return false;
            }
        }

        return PdfPermissionAuthorization.CanMutate(Probe.Security, PermissionPolicy, PdfMutationOperation.ModifyPageContent) ||
            PdfPermissionAuthorization.CanMutate(Probe.Security, PermissionPolicy, PdfMutationOperation.ModifyPageTree) ||
            PdfPermissionAuthorization.CanMutate(Probe.Security, PermissionPolicy, PdfMutationOperation.MergeDocuments);
    }

    /// <summary>Returns true when the requested wrapper-facing capability can be attempted for this PDF.</summary>
    public bool Can(PdfPreflightCapability capability) {
        switch (capability) {
            case PdfPreflightCapability.ExtractText:
                return CanExtractText;
            case PdfPreflightCapability.ExtractImages:
                return CanExtractImages;
            case PdfPreflightCapability.ExtractAttachments:
                return CanExtractAttachments;
            case PdfPreflightCapability.ReadLogicalObjects:
                return CanReadLogicalObjects;
            case PdfPreflightCapability.ManipulatePages:
                return CanManipulatePages;
            case PdfPreflightCapability.AppendMetadataRevision:
                return CanAppendMetadataRevision;
            case PdfPreflightCapability.AppendFormFieldRevision:
                return CanAppendFormFieldRevision;
            case PdfPreflightCapability.PrepareExternalSignatureRevision:
                return CanPrepareExternalSignatureRevision;
            case PdfPreflightCapability.FillSimpleFormFields:
                return CanFillSimpleFormFields;
            case PdfPreflightCapability.FlattenSimpleFormFields:
                return CanFlattenSimpleFormFields;
            case PdfPreflightCapability.FillAndFlattenSimpleFormFields:
                return CanFillAndFlattenSimpleFormFields;
            default:
                throw new ArgumentOutOfRangeException(nameof(capability), capability, "Unsupported PDF preflight capability.");
        }
    }

    /// <summary>Returns operation-specific diagnostics explaining why a wrapper-facing capability is blocked, or an empty list when it can be attempted.</summary>
    public IReadOnlyList<string> GetCapabilityDiagnostics(PdfPreflightCapability capability) {
        if (Can(capability)) {
            return Array.Empty<string>();
        }

        switch (capability) {
            case PdfPreflightCapability.ExtractText:
                return GetTextExtractionDiagnostics();
            case PdfPreflightCapability.ExtractImages:
                return GetImageExtractionDiagnostics();
            case PdfPreflightCapability.ExtractAttachments:
                return GetAttachmentExtractionDiagnostics();
            case PdfPreflightCapability.ReadLogicalObjects:
                return GetReadCapabilityDiagnostics("PDF logical object extraction is not available because OfficeIMO.Pdf cannot read this PDF.");
            case PdfPreflightCapability.ManipulatePages:
                return GetPageManipulationDiagnostics();
            case PdfPreflightCapability.AppendMetadataRevision:
                return GetAppendOnlyCapabilityDiagnostics("Metadata", "PDF append-only metadata revision is not available for this PDF.");
            case PdfPreflightCapability.AppendFormFieldRevision:
                return GetAppendOnlyCapabilityDiagnostics("FormFill", "PDF append-only form-field revision is not available for this PDF.");
            case PdfPreflightCapability.PrepareExternalSignatureRevision:
                return GetAppendOnlyCapabilityDiagnostics("SignaturePrepare", "PDF append-only external-signature preparation is not available for this PDF.");
            case PdfPreflightCapability.FillSimpleFormFields:
                return GetSimpleFormCapabilityDiagnostics(requireFillableField: true, requireFlattenableWidget: false);
            case PdfPreflightCapability.FlattenSimpleFormFields:
                return GetSimpleFormCapabilityDiagnostics(requireFillableField: false, requireFlattenableWidget: true);
            case PdfPreflightCapability.FillAndFlattenSimpleFormFields:
                return GetSimpleFormCapabilityDiagnostics(requireFillableField: true, requireFlattenableWidget: true);
            default:
                throw new ArgumentOutOfRangeException(nameof(capability), capability, "Unsupported PDF preflight capability.");
        }
    }

    private bool HasFormMutationBlocker(PdfMutationOperation operation) {
        return Probe.HasSignatures ||
            Probe.HasActiveContent ||
            _documentInfo?.AcroFormSignaturesExist == true ||
            _documentInfo?.HasActiveContent == true ||
            !PdfPermissionAuthorization.CanRewriteFormFields(Probe.Security, PermissionPolicy, operation);
    }

    private bool HasImageExtractionBlocker() {
        return HasReadBlocker(PdfReadBlockerKind.MissingHeader) ||
            HasReadBlocker(PdfReadBlockerKind.Encryption) ||
            HasReadBlocker(PdfReadBlockerKind.NoPages) ||
            HasReadBlocker(PdfReadBlockerKind.ParserUnsupported);
    }

    private bool HasAttachmentExtractionBlocker() {
        return HasReadBlocker(PdfReadBlockerKind.MissingHeader) ||
            HasReadBlocker(PdfReadBlockerKind.Encryption) ||
            HasReadBlocker(PdfReadBlockerKind.ParserUnsupported);
    }

    private bool HasSimpleFillableFormFields() {
        if (_documentInfo is null || _documentInfo.FormFields.Count == 0) {
            return false;
        }

        for (int i = 0; i < _documentInfo.FormFields.Count; i++) {
            PdfFormField field = _documentInfo.FormFields[i];
            if (IsNamedSimpleFillField(field)) {
                return true;
            }
        }

        return false;
    }

    private bool HasSimpleFlattenableFormFields() {
        if (_documentInfo is null || _documentInfo.FormFields.Count == 0) {
            return false;
        }

        bool hasFlattenableWidget = false;
        for (int i = 0; i < _documentInfo.FormFields.Count; i++) {
            PdfFormField field = _documentInfo.FormFields[i];
            if (!IsNamedSimpleFlattenField(field) || field.Widgets.Count == 0) {
                return false;
            }

            for (int j = 0; j < field.Widgets.Count; j++) {
                PdfFormWidget widget = field.Widgets[j];
                if (!widget.ObjectNumber.HasValue ||
                    !widget.PageNumber.HasValue ||
                    widget.Width <= 0D ||
                    widget.Height <= 0D) {
                    return false;
                }

                hasFlattenableWidget = true;
            }
        }

        return hasFlattenableWidget;
    }

    private static bool IsNamedSimpleFillField(PdfFormField field) {
        return !string.IsNullOrEmpty(field.Name) &&
            (field.Kind == PdfFormFieldKind.Text ||
            field.Kind == PdfFormFieldKind.Choice ||
            field.Kind == PdfFormFieldKind.Button);
    }

    private static bool IsNamedSimpleFlattenField(PdfFormField field) {
        return !string.IsNullOrEmpty(field.Name) &&
            (field.Kind == PdfFormFieldKind.Text ||
            field.Kind == PdfFormFieldKind.Choice ||
            field.Kind == PdfFormFieldKind.Button);
    }

    private IReadOnlyList<string> GetReadCapabilityDiagnostics(string fallbackMessage) {
        if (ReadBlockers.Count == 0) {
            return new[] { fallbackMessage };
        }

        var messages = new List<string>(ReadBlockers.Count);
        for (int i = 0; i < ReadBlockers.Count; i++) {
            AddDistinct(messages, ReadBlockers[i].Message);
        }

        AddRange(messages, SecurityDiagnostics);
        return messages.AsReadOnly();
    }

    private IReadOnlyList<string> GetTextExtractionDiagnostics() {
        if (CanRead && !PdfPermissionAuthorization.CanExtractText(Probe.Security, PermissionPolicy)) {
            return new[] { "PDF text extraction is restricted by the authenticated user-password permissions. Supply owner authorization or explicitly ignore permission restrictions." };
        }

        return GetReadCapabilityDiagnostics("PDF text extraction is not available because OfficeIMO.Pdf cannot read this PDF.");
    }

    private IReadOnlyList<string> GetImageExtractionDiagnostics() {
        if (CanRead && !PdfPermissionAuthorization.CanExtractContent(Probe.Security, PermissionPolicy)) {
            return new[] { "PDF image extraction is restricted by the authenticated user-password permissions. Supply owner authorization or explicitly ignore permission restrictions." };
        }

        if (ReadBlockers.Count == 0) {
            return new[] { "PDF image extraction is not available because OfficeIMO.Pdf cannot inspect this PDF." };
        }

        var messages = new List<string>(ReadBlockers.Count);
        for (int i = 0; i < ReadBlockers.Count; i++) {
            if (ReadBlockers[i].Kind != PdfReadBlockerKind.UnsupportedContentStreamFilter) {
                AddDistinct(messages, ReadBlockers[i].Message);
            }
        }

        AddRange(messages, SecurityDiagnostics);
        if (messages.Count == 0) {
            AddDistinct(messages, "PDF image extraction is not available for this PDF.");
        }

        return messages.AsReadOnly();
    }

    private IReadOnlyList<string> GetAttachmentExtractionDiagnostics() {
        if (CanRead && !PdfPermissionAuthorization.CanExtractContent(Probe.Security, PermissionPolicy)) {
            return new[] { "PDF attachment extraction is restricted by the authenticated user-password permissions. Supply owner authorization or explicitly ignore permission restrictions." };
        }

        if (ReadBlockers.Count == 0) {
            return new[] { "PDF attachment extraction is not available because OfficeIMO.Pdf cannot inspect this PDF." };
        }

        var messages = new List<string>(ReadBlockers.Count);
        for (int i = 0; i < ReadBlockers.Count; i++) {
            if (ReadBlockers[i].Kind != PdfReadBlockerKind.NoPages &&
                ReadBlockers[i].Kind != PdfReadBlockerKind.UnsupportedContentStreamFilter) {
                AddDistinct(messages, ReadBlockers[i].Message);
            }
        }

        AddRange(messages, SecurityDiagnostics);
        if (messages.Count == 0) {
            AddDistinct(messages, "PDF attachment extraction is not available for this PDF.");
        }

        return messages.AsReadOnly();
    }

    private System.Collections.ObjectModel.ReadOnlyCollection<string> GetPageManipulationDiagnostics() {
        var messages = new List<string>(ReadBlockers.Count + RewriteBlockers.Count);
        for (int i = 0; i < ReadBlockers.Count; i++) {
            AddDistinct(messages, ReadBlockers[i].Message);
        }

        for (int i = 0; i < RewriteBlockers.Count; i++) {
            AddDistinct(messages, RewriteBlockers[i].Message);
        }

        AddRange(messages, SecurityDiagnostics);
        if (messages.Count == 0) {
            AddDistinct(messages, "PDF page manipulation is not available for this PDF.");
        }

        return messages.AsReadOnly();
    }

    private System.Collections.ObjectModel.ReadOnlyCollection<string> GetAppendOnlyCapabilityDiagnostics(string action, string fallbackMessage) {
        var messages = new List<string>();
        PdfAppendOnlyMutationReport report = AppendOnlyMutationReport;
        if (report.BlockedActions.Contains(action, StringComparer.Ordinal)) {
            AddDistinct(messages, fallbackMessage);
        }

        if (!CanRead) {
            AddRange(messages, GetReadCapabilityDiagnostics(fallbackMessage));
        } else {
            AddRange(messages, SecurityDiagnostics);
        }

        for (int i = 0; i < report.Blockers.Count; i++) {
            AddDistinct(messages, "Append-only blocker: " + report.Blockers[i] + ".");
        }

        for (int i = 0; i < report.Warnings.Count; i++) {
            AddDistinct(messages, "Append-only warning: " + report.Warnings[i] + ".");
        }

        if (messages.Count == 0) {
            AddDistinct(messages, fallbackMessage);
        }

        return messages.AsReadOnly();
    }

    private System.Collections.ObjectModel.ReadOnlyCollection<string> GetSimpleFormCapabilityDiagnostics(bool requireFillableField, bool requireFlattenableWidget) {
        var messages = new List<string>();
        if (!CanRead) {
            AddRange(messages, GetReadCapabilityDiagnostics("PDF form operations are not available because OfficeIMO.Pdf cannot read this PDF."));
            return messages.AsReadOnly();
        }

        if (Probe.HasEncryption &&
            !Probe.Security.HasOwnerAuthorization &&
            !PermissionRestrictionsIgnored) {
            AddDistinct(messages, "Encrypted PDF form filling and flattening require owner authorization or an explicit IgnoreRestrictions permission policy because the operation performs an unencrypted full rewrite.");
            AddRange(messages, SecurityDiagnostics);
        }

        if (Probe.HasSignatures || _documentInfo?.AcroFormSignaturesExist == true) {
            AddDistinct(messages, "Signed PDF files are not supported for form filling or flattening by OfficeIMO.Pdf yet.");
            AddRange(messages, SignatureMutationDiagnostics);
        }

        if (Probe.HasActiveContent || _documentInfo?.HasActiveContent == true) {
            AddDistinct(messages, "PDF active content is not supported for form filling or flattening by OfficeIMO.Pdf yet.");
        }

        if (HasRewriteBlocker(PdfRewriteBlockerKind.Encryption)) {
            AddRange(messages, GetPageManipulationDiagnostics());
        }

        if (requireFillableField && !HasSimpleFillableFormFields()) {
            AddDistinct(messages, "PDF does not contain named text, choice, or button AcroForm fields supported for simple form filling by OfficeIMO.Pdf.");
        }

        if (requireFlattenableWidget && !HasSimpleFlattenableFormFields()) {
            AddDistinct(messages, "PDF does not contain named text, choice, or button AcroForm widgets with readable page-backed rectangles supported for simple form flattening by OfficeIMO.Pdf.");
        }

        if (messages.Count == 0) {
            AddDistinct(messages, "PDF simple form operation is not available for this PDF.");
        }

        return messages.AsReadOnly();
    }

    private static void AddRange(List<string> messages, IReadOnlyList<string> values) {
        for (int i = 0; i < values.Count; i++) {
            AddDistinct(messages, values[i]);
        }
    }

    private static void AddDistinct(List<string> messages, string message) {
        if (string.IsNullOrWhiteSpace(message)) {
            return;
        }

        for (int i = 0; i < messages.Count; i++) {
            if (string.Equals(messages[i], message, StringComparison.Ordinal)) {
                return;
            }
        }

        messages.Add(message);
    }
}

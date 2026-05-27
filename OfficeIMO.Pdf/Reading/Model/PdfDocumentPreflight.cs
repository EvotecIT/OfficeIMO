namespace OfficeIMO.Pdf;

/// <summary>
/// Wrapper-friendly PDF capability report for OfficeIMO.Pdf read and rewrite operations.
/// </summary>
public sealed class PdfDocumentPreflight {
    internal PdfDocumentPreflight(
        PdfDocumentProbe probe,
        PdfDocumentInfo? documentInfo,
        bool canRead,
        bool canRewrite,
        IReadOnlyList<string> diagnostics,
        IReadOnlyList<PdfReadBlocker> readBlockers,
        IReadOnlyList<PdfRewriteBlocker> rewriteBlockers) {
        Probe = probe;
        DocumentInfo = documentInfo;
        CanRead = canRead;
        CanRewrite = canRewrite;
        Diagnostics = diagnostics;
        ReadBlockers = readBlockers;
        RewriteBlockers = rewriteBlockers;
    }

    /// <summary>Lightweight PDF markers read before full parsing.</summary>
    public PdfDocumentProbe Probe { get; }

    /// <summary>Parsed document information when the document can be inspected.</summary>
    public PdfDocumentInfo? DocumentInfo { get; }

    /// <summary>True when OfficeIMO.Pdf can parse enough of the document for read-oriented operations.</summary>
    public bool CanRead { get; }

    /// <summary>True when OfficeIMO.Pdf can attempt rewrite-style manipulation without known security blockers.</summary>
    public bool CanRewrite { get; }

    /// <summary>True when OfficeIMO.Pdf can attempt text and logical readback operations for this PDF.</summary>
    public bool CanExtractText => CanRead;

    /// <summary>True when OfficeIMO.Pdf can attempt page-level rewrite operations such as extract, split, merge, import, edit, stamp, and metadata updates.</summary>
    public bool CanManipulatePages => CanRewrite;

    /// <summary>True when OfficeIMO.Pdf can attempt simple AcroForm value updates for named text, choice, or button fields.</summary>
    public bool CanFillSimpleFormFields => CanRead && !HasFormMutationBlocker() && HasSimpleFillableFormFields();

    /// <summary>True when OfficeIMO.Pdf can attempt simple AcroForm flattening for text or button widgets with page-backed rectangles.</summary>
    public bool CanFlattenSimpleFormFields => CanRead && !HasFormMutationBlocker() && HasSimpleFlattenableFormFields();

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

    private bool HasFormMutationBlocker() {
        return Probe.HasSignatures ||
            Probe.HasActiveContent ||
            DocumentInfo?.AcroFormSignaturesExist == true ||
            DocumentInfo?.HasActiveContent == true;
    }

    private bool HasSimpleFillableFormFields() {
        if (DocumentInfo is null || DocumentInfo.FormFields.Count == 0) {
            return false;
        }

        for (int i = 0; i < DocumentInfo.FormFields.Count; i++) {
            PdfFormField field = DocumentInfo.FormFields[i];
            if (IsNamedSimpleFillField(field)) {
                return true;
            }
        }

        return false;
    }

    private bool HasSimpleFlattenableFormFields() {
        if (DocumentInfo is null || DocumentInfo.FormFields.Count == 0) {
            return false;
        }

        bool hasFlattenableWidget = false;
        for (int i = 0; i < DocumentInfo.FormFields.Count; i++) {
            PdfFormField field = DocumentInfo.FormFields[i];
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
            field.Kind == PdfFormFieldKind.Button);
    }
}

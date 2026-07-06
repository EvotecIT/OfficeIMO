namespace OfficeIMO.Pdf;

/// <summary>
/// Wrapper-facing OfficeIMO.Pdf operation categories covered by preflight checks.
/// </summary>
public enum PdfPreflightCapability {
    /// <summary>Text and structured text readback operations.</summary>
    ExtractText,

    /// <summary>Image XObject extraction operations.</summary>
    ExtractImages,

    /// <summary>Embedded-file and associated-file attachment extraction operations.</summary>
    ExtractAttachments,

    /// <summary>Page-level rewrite operations such as extract, split, merge, import, edit, stamp, and metadata updates.</summary>
    ManipulatePages,

    /// <summary>Append-only metadata revision updates that preserve the existing PDF bytes.</summary>
    AppendMetadataRevision,

    /// <summary>Append-only AcroForm field-value revision updates that preserve the existing PDF bytes.</summary>
    AppendFormFieldRevision,

    /// <summary>Append-only external-signature placeholder revision preparation.</summary>
    PrepareExternalSignatureRevision,

    /// <summary>Simple AcroForm value updates for named text, choice, or button fields.</summary>
    FillSimpleFormFields,

    /// <summary>Simple AcroForm flattening for named text, choice, or button widgets with page-backed rectangles.</summary>
    FlattenSimpleFormFields,

    /// <summary>Simple AcroForm value updates followed by simple widget flattening.</summary>
    FillAndFlattenSimpleFormFields,

    /// <summary>Logical object readback through <see cref="PdfLogicalDocument"/>.</summary>
    ReadLogicalObjects
}

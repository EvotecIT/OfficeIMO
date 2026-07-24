namespace OfficeIMO.Pdf;

/// <summary>
/// Wrapper-facing OfficeIMO.Pdf operation categories covered by preflight checks.
/// </summary>
public enum PdfPreflightCapability {
    /// <summary>Text and structured text readback operations.</summary>
    ExtractText = 0,

    /// <summary>Image XObject extraction operations.</summary>
    ExtractImages = 5,

    /// <summary>Embedded-file and associated-file attachment extraction operations.</summary>
    ExtractAttachments = 7,

    /// <summary>Page-level rewrite operations such as extract, split, merge, import, edit, stamp, and metadata updates.</summary>
    ManipulatePages = 1,

    /// <summary>Append-only metadata revision updates that preserve the existing PDF bytes.</summary>
    AppendMetadataRevision = 8,

    /// <summary>Append-only AcroForm field-value revision updates that preserve the existing PDF bytes.</summary>
    AppendFormFieldRevision = 9,

    /// <summary>Append-only external-signature placeholder revision preparation.</summary>
    PrepareExternalSignatureRevision = 10,

    /// <summary>Simple AcroForm value updates for named text, choice, or button fields.</summary>
    FillSimpleFormFields = 2,

    /// <summary>Simple AcroForm flattening for named text, choice, or button widgets with page-backed rectangles.</summary>
    FlattenSimpleFormFields = 3,

    /// <summary>Simple AcroForm value updates followed by simple widget flattening.</summary>
    FillAndFlattenSimpleFormFields = 4,

    /// <summary>Logical object readback through <see cref="PdfLogicalDocument"/>.</summary>
    ReadLogicalObjects = 6
}

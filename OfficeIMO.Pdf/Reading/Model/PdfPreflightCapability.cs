namespace OfficeIMO.Pdf;

/// <summary>
/// Wrapper-facing OfficeIMO.Pdf operation categories covered by preflight checks.
/// </summary>
public enum PdfPreflightCapability {
    /// <summary>Text and structured text readback operations.</summary>
    ExtractText,

    /// <summary>Image XObject extraction operations.</summary>
    ExtractImages,

    /// <summary>Page-level rewrite operations such as extract, split, merge, import, edit, stamp, and metadata updates.</summary>
    ManipulatePages,

    /// <summary>Simple AcroForm value updates for named text, choice, or button fields.</summary>
    FillSimpleFormFields,

    /// <summary>Simple AcroForm flattening for named text or button widgets with page-backed rectangles.</summary>
    FlattenSimpleFormFields,

    /// <summary>Simple AcroForm value updates followed by simple widget flattening.</summary>
    FillAndFlattenSimpleFormFields,

    /// <summary>Logical object readback through <see cref="PdfLogicalDocument"/>.</summary>
    ReadLogicalObjects
}

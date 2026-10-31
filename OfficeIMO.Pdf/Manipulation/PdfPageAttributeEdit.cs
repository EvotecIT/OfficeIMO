namespace OfficeIMO.Pdf;

/// <summary>An immutable page-attribute edit for an atomic, ordered batch.</summary>
/// <remarks>A null selection means every page. Edits never remove content from source streams.</remarks>
/// <remarks>SetBox uses native page user-space coordinates; one unit is /UserUnit physical points (1/72 inch), defaulting to 1. CropMargins uses physical-point distances from the displayed edges after crop, rotation and /UserUnit mapping.</remarks>
public sealed class PdfPageAttributeEdit {
    private PdfPageAttributeEdit(Kind kind, PdfPageSelection? selection, int rotation = 0,
        PdfPageBoundaryBox box = PdfPageBoundaryBox.CropBox,
        double left = 0, double bottom = 0, double right = 0, double top = 0) {
        Operation = kind; Selection = selection; Rotation = rotation; Box = box;
        Left = left; Bottom = bottom; Right = right; Top = top;
    }

    /// <summary>Selected one-based pages, or null for every page.</summary>
    public PdfPageSelection? Selection { get; }
    internal Kind Operation { get; }
    internal int Rotation { get; }
    internal PdfPageBoundaryBox Box { get; }
    internal double Left { get; }
    internal double Bottom { get; }
    internal double Right { get; }
    internal double Top { get; }

    /// <summary>Sets the absolute page rotation to a multiple of 90 degrees.</summary>
    public static PdfPageAttributeEdit SetRotation(int degrees, PdfPageSelection? selection = null) =>
        new PdfPageAttributeEdit(Kind.SetRotation, selection, rotation: degrees);

    /// <summary>Turns each selected page clockwise by a multiple of 90 degrees relative to its effective rotation.</summary>
    public static PdfPageAttributeEdit RotateBy(int degrees, PdfPageSelection? selection = null) =>
        new PdfPageAttributeEdit(Kind.RotateBy, selection, rotation: degrees);

    /// <summary>Sets a boundary box in unrotated PDF user-space coordinates.</summary>
    public static PdfPageAttributeEdit SetBox(PdfPageBoundaryBox box, double left, double bottom,
        double right, double top, PdfPageSelection? selection = null) =>
        new PdfPageAttributeEdit(Kind.SetBox, selection, box: box, left: left, bottom: bottom, right: right, top: top);

    /// <summary>Crops inward from the current displayed edges, using physical points (1/72 inch).</summary>
    /// <remarks>Resolves rotation, crop origins, and UserUnit after earlier edits in the same batch.</remarks>
    public static PdfPageAttributeEdit CropMargins(double left, double bottom, double right, double top,
        PdfPageSelection? selection = null) =>
        new PdfPageAttributeEdit(Kind.CropMargins, selection, left: left, bottom: bottom, right: right, top: top);

    /// <summary>Restores the visible crop to the current MediaBox without changing content.</summary>
    public static PdfPageAttributeEdit ResetCrop(PdfPageSelection? selection = null) =>
        new PdfPageAttributeEdit(Kind.ResetCrop, selection);

    internal enum Kind { SetRotation, RotateBy, SetBox, CropMargins, ResetCrop }
}

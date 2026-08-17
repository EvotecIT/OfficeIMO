using OfficeIMO.PowerPoint.Pdf;
using OfficeIMO.Web.Converter.Models;

namespace OfficeIMO.Web.Converter.Services;

public static class BrowserPowerPointImportProfileCatalog {
    public static BrowserPowerPointImportProfile Editable { get; } = new(
        PdfPowerPointImportMode.EditableContent,
        "editable",
        "Editable content",
        "Reconstruct text blocks, detected tables, safe vector primitives, and separate raster images as native PowerPoint objects. Complex PDF authoring intent cannot always be recovered.",
        "Reconstructed",
        "editable-content-slides");

    public static BrowserPowerPointImportProfile Visual { get; } = new(
        PdfPowerPointImportMode.VisualPages,
        "visual",
        "Visual pages",
        "Place each PDF page on a slide as one page-sized image. Appearance is prioritized, but text, shapes, charts, and tables are not editable.",
        "Visual",
        "visual-page-slides");

    public static BrowserPowerPointImportProfile Hybrid { get; } = new(
        PdfPowerPointImportMode.HybridVisualAndEditableTables,
        "hybrid",
        "Visual + editable tables",
        "Keep each PDF page as a visual background and overlay detected tables as native PowerPoint tables. Other page content remains part of the image.",
        "Hybrid",
        "hybrid-visual-table-slides");

    public static BrowserPowerPointImportProfile Tables { get; } = new(
        PdfPowerPointImportMode.EditableTables,
        "tables",
        "Tables only",
        "Create native PowerPoint tables from detected PDF tables. Text and graphics outside detected tables are omitted.",
        "Partial",
        "editable-table-slides");

    public static IReadOnlyList<BrowserPowerPointImportProfile> All { get; } = [
        Editable,
        Visual,
        Hybrid,
        Tables
    ];

    public static BrowserPowerPointImportProfile Find(string? id) =>
        All.FirstOrDefault(profile => string.Equals(profile.Id, id, StringComparison.OrdinalIgnoreCase))
        ?? Editable;

    public static BrowserPowerPointImportProfile Find(PdfPowerPointImportMode mode) =>
        All.FirstOrDefault(profile => profile.Mode == mode)
        ?? Editable;
}

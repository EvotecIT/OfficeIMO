using PptCore = OfficeIMO.PowerPoint;

namespace OfficeIMO.PowerPoint.Pdf;

/// <summary>Defines how PDF content is reconstructed in PowerPoint.</summary>
public enum PdfPowerPointImportMode {
    /// <summary>Creates one high-fidelity rendered image slide per PDF page.</summary>
    VisualPages,
    /// <summary>Reconstructs detected tables as editable PowerPoint tables.</summary>
    EditableTables,
    /// <summary>Keeps a rendered visual page and overlays detected tables as editable PowerPoint tables.</summary>
    HybridVisualAndEditableTables,
    /// <summary>Reconstructs text, detected tables, safe vector primitives, and supported images as editable slide objects.</summary>
    EditableContent
}

/// <summary>
/// Options for importing PDF content into a PowerPoint presentation.
/// </summary>
public sealed class PdfPowerPointImportOptions {
    /// <summary>Import strategy. Defaults to one visual slide per PDF page.</summary>
    public PdfPowerPointImportMode Mode { get; set; } = PdfPowerPointImportMode.VisualPages;

    /// <summary>Optional caller-ordered page selection used by all import modes.</summary>
    public OfficeIMO.Pdf.PdfPageSelection? PageSelection { get; set; }

    /// <summary>Raster resolution used by visual-page import.</summary>
    public double Dpi { get; set; } = 144D;

    /// <summary>Maximum pages processed by one import.</summary>
    public int MaxPages { get; set; } = 100;

    /// <summary>Maximum output pixels for one rendered PDF page.</summary>
    public long MaxPixelsPerPage { get; set; } = 64L * 1024L * 1024L;

    /// <summary>Maximum encoded PNG bytes retained for one rendered PDF page.</summary>
    public long MaxOutputBytesPerPage { get; set; } = 64L * 1024L * 1024L;

    /// <summary>Maximum aggregate encoded PNG bytes retained by visual-page import.</summary>
    public long MaxTotalOutputBytes { get; set; } = 256L * 1024L * 1024L;

    /// <summary>Creates the editable-table reconstruction profile.</summary>
    public static PdfPowerPointImportOptions CreateEditableTables() => new PdfPowerPointImportOptions {
        Mode = PdfPowerPointImportMode.EditableTables
    };

    /// <summary>Creates the hybrid visual-page plus editable-table overlay profile.</summary>
    public static PdfPowerPointImportOptions CreateHybrid() => new PdfPowerPointImportOptions {
        Mode = PdfPowerPointImportMode.HybridVisualAndEditableTables
    };

    /// <summary>Creates the semantic editable-content reconstruction profile.</summary>
    public static PdfPowerPointImportOptions CreateEditableContent() => new PdfPowerPointImportOptions {
        Mode = PdfPowerPointImportMode.EditableContent
    };

    /// <summary>
    /// Caller-supplied deterministic fonts used by visual and hybrid page rasterization when the
    /// source PDF does not contain a supported embedded font program.
    /// </summary>
    public OfficeIMO.Drawing.OfficeFontFaceCollection RenderFonts { get; set; } = new();

    /// <summary>Optional text shaper used by visual and hybrid page rasterization.</summary>
    public OfficeIMO.Drawing.IOfficeTextShapingProvider? TextShapingProvider { get; set; }

    /// <summary>Optional BCP 47 language hint passed to <see cref="TextShapingProvider"/>.</summary>
    public string? TextShapingLanguage { get; set; }

    /// <summary>Maximum editable text boxes, vector shapes, images, and tables created for one source page.</summary>
    public int MaxEditableObjectsPerPage { get; set; } = 5_000;

    /// <summary>
    /// Maximum body rows to import per detected table. Values less than or equal to zero import all rows.
    /// </summary>
    public int MaxRows { get; set; }

    /// <summary>
    /// Maximum body rows written to one PowerPoint slide. Values less than or equal to zero keep all imported rows on one slide.
    /// </summary>
    public int MaxRowsPerSlide { get; set; }

    /// <summary>
    /// Maximum columns written to one PowerPoint slide. Values less than or equal to zero keep all columns on one slide.
    /// </summary>
    public int MaxColumnsPerSlide { get; set; }

    /// <summary>
    /// PowerPoint table style applied to imported tables.
    /// </summary>
    public PptCore.PowerPointTableStylePreset TableStyle { get; set; } = PptCore.PowerPointTableStylePreset.Default;

    /// <summary>
    /// When true, a slide title describing the source PDF page and table is added above each imported table.
    /// </summary>
    public bool IncludeSourceTitles { get; set; } = true;

    /// <summary>
    /// When true, inferred column names are written as a PowerPoint table header row.
    /// </summary>
    public bool IncludeColumnHeaderRows { get; set; } = true;

    /// <summary>
    /// When true, banded row styling is enabled on imported tables.
    /// </summary>
    public bool BandedRows { get; set; } = true;

    /// <summary>
    /// When true, body cells in inferred numeric PDF columns are right-aligned in the generated PowerPoint tables.
    /// </summary>
    public bool AlignNumericColumns { get; set; } = true;

    /// <summary>
    /// Left position of each imported table in EMUs.
    /// </summary>
    public long TableLeft { get; set; } = 457200L;

    /// <summary>
    /// Top position of each imported table in EMUs.
    /// </summary>
    public long TableTop { get; set; } = 1371600L;

    /// <summary>
    /// Width of each imported table in EMUs.
    /// </summary>
    public long TableWidth { get; set; } = 11277600L;

    /// <summary>
    /// Height of each imported table in EMUs.
    /// </summary>
    public long TableHeight { get; set; } = 4876800L;

    /// <summary>
    /// Slide title text written when no tables are detected, keeping the produced presentation meaningful.
    /// </summary>
    public string EmptyPresentationTitle { get; set; } = "PDF Tables";

    /// <summary>
    /// Slide body text written when no tables are detected.
    /// </summary>
    public string EmptyPresentationMessage { get; set; } = "No PDF tables detected.";
}

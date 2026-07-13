using PptCore = OfficeIMO.PowerPoint;

namespace OfficeIMO.PowerPoint.Pdf;

/// <summary>
/// Options for extracting logical PDF tables into a PowerPoint presentation.
/// </summary>
public sealed class PdfPowerPointTableImportOptions {
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

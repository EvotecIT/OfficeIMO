using OfficeIMO.Html;

namespace OfficeIMO.PowerPoint.Html;

/// <summary>
/// Options for importing semantic OfficeIMO PowerPoint HTML into a presentation.
/// </summary>
public sealed class HtmlPowerPointExportOptions {
    private HtmlImportLimits _limits = new HtmlImportLimits();

    /// <summary>Shared native-artifact limits for this import operation.</summary>
    public HtmlImportLimits Limits {
        get => _limits;
        set => _limits = value ?? new HtmlImportLimits();
    }

    /// <summary>
    /// Maximum rectangular table size imported from one HTML table, including merged spans.
    /// </summary>
    public int MaxTableCells {
        get => Limits.MaxTableCells;
        set => Limits.MaxTableCells = value;
    }

    /// <summary>Controls semantic restoration versus ordinary HTML section import.</summary>
    public HtmlImportMode Mode { get; set; } = HtmlImportMode.Semantic;

    /// <summary>
    /// Imports semantic slide tables as native PowerPoint tables.
    /// </summary>
    public bool ImportTables { get; set; } = true;

    /// <summary>
    /// Imports embedded data URI pictures from the semantic picture inventory.
    /// </summary>
    public bool ImportPictures { get; set; } = true;

    /// <summary>
    /// Imports chart inventory items as native charts with reconstructed placeholder data.
    /// </summary>
    public bool ImportChartInventory { get; set; } = true;

    /// <summary>
    /// Imports presenter notes when they are present in the semantic extraction proof.
    /// </summary>
    public bool ImportNotes { get; set; } = true;

    /// <summary>Maps bounded positioned, floating, flex, and grid regions to editable slide geometry.</summary>
    public bool ImportEditableLayoutRegions { get; set; } = true;

    /// <summary>Policy applied to imported hyperlinks, including target-specific run metadata.</summary>
    public HtmlUrlPolicy HyperlinkUrlPolicy { get; set; } = HtmlUrlPolicy.CreateHyperlinkProfile();

    // Ordinary href values have already passed the document policy during HTML
    // normalization. Target-specific metadata still uses the complete policy.
    internal HtmlUrlPolicy? NormalizedHyperlinkUrlPolicy { get; set; }

    /// <summary>Throws <see cref="OfficeConversionException"/> when the conversion reports loss.</summary>
    public bool FailOnLoss { get; set; }

    /// <summary>Creates an independent copy of these settings.</summary>
    public HtmlPowerPointExportOptions Clone() => new HtmlPowerPointExportOptions {
        Limits = Limits.Clone(),
        Mode = Mode,
        ImportTables = ImportTables,
        ImportPictures = ImportPictures,
        ImportChartInventory = ImportChartInventory,
        ImportNotes = ImportNotes,
        ImportEditableLayoutRegions = ImportEditableLayoutRegions,
        HyperlinkUrlPolicy = HyperlinkUrlPolicy.Clone(),
        NormalizedHyperlinkUrlPolicy = NormalizedHyperlinkUrlPolicy?.Clone(),
        FailOnLoss = FailOnLoss
    };
}

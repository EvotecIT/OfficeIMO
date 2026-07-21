using OfficeIMO.Html;

namespace OfficeIMO.PowerPoint.Html;

/// <summary>
/// Options for importing semantic OfficeIMO PowerPoint HTML into a presentation.
/// </summary>
public sealed class HtmlToPowerPointOptions {
    private HtmlImportLimits _limits = HtmlImportLimits.CreateDefault();

    /// <summary>Shared native-artifact limits for this import operation.</summary>
    public HtmlImportLimits Limits {
        get => _limits;
        set => _limits = value ?? HtmlImportLimits.CreateDefault();
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

    internal HtmlToPowerPointOptions Clone() => new HtmlToPowerPointOptions {
        Limits = Limits.Clone(),
        Mode = Mode,
        ImportTables = ImportTables,
        ImportPictures = ImportPictures,
        ImportChartInventory = ImportChartInventory,
        ImportNotes = ImportNotes
    };
}

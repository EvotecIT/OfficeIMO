using OfficeIMO.Html;

namespace OfficeIMO.Excel.Html;

/// <summary>
/// Options for importing semantic OfficeIMO Excel HTML into a workbook.
/// </summary>
public sealed class HtmlToExcelOptions {
    private HtmlImportLimits _limits = HtmlImportLimits.CreateDefault();

    /// <summary>Shared native-artifact limits for this import operation.</summary>
    public HtmlImportLimits Limits {
        get => _limits;
        set => _limits = value ?? HtmlImportLimits.CreateDefault();
    }

    /// <summary>
    /// Maximum number of HTML table grid slots imported across each worksheet table, including merged spans.
    /// </summary>
    public int MaxTableCells {
        get => Limits.MaxTableCells;
        set => Limits.MaxTableCells = value;
    }

    /// <summary>Controls semantic restoration versus ordinary HTML table import.</summary>
    public HtmlImportMode Mode { get; set; } = HtmlImportMode.Semantic;

    /// <summary>
    /// Imports embedded data URI images from the semantic image inventory.
    /// </summary>
    public bool ImportImages { get; set; } = true;

    /// <summary>
    /// Imports chart inventory items as native charts when the worksheet table contains enough data.
    /// </summary>
    public bool ImportChartInventory { get; set; } = true;

    /// <summary>
    /// Imports comment inventory items as native cell comments.
    /// </summary>
    public bool ImportComments { get; set; } = true;

    /// <summary>
    /// Imports formula inventory items as native cell formulas.
    /// </summary>
    public bool ImportFormulas { get; set; } = true;

    /// <summary>
    /// Allows formula restoration from caller-untrusted semantic HTML when <see cref="ImportFormulas"/> is enabled.
    /// Keep disabled unless the caller has independently validated the formula source.
    /// </summary>
    public bool AllowUntrustedFormulas { get; set; }

    internal HtmlToExcelOptions Clone() => new HtmlToExcelOptions {
        Limits = Limits.Clone(),
        Mode = Mode,
        ImportImages = ImportImages,
        ImportChartInventory = ImportChartInventory,
        ImportComments = ImportComments,
        ImportFormulas = ImportFormulas,
        AllowUntrustedFormulas = AllowUntrustedFormulas
    };
}

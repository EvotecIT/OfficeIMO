namespace OfficeIMO.Html;

/// <summary>
/// Shared native-artifact limits for HTML import adapters such as Excel and PowerPoint.
/// </summary>
public sealed class HtmlImportLimits {
    /// <summary>Creates the bounded default profile used for untrusted semantic HTML.</summary>
    public static HtmlImportLimits CreateDefault() => new HtmlImportLimits();

    /// <summary>Maximum semantic sheets, slides, or equivalent top-level containers.</summary>
    public int MaxSemanticContainers { get; set; } = 1_000;

    /// <summary>Maximum imported native tables.</summary>
    public int MaxTables { get; set; } = 1_000;

    /// <summary>Maximum grid slots in one imported table, including spans.</summary>
    public int MaxTableCells { get; set; } = 50_000;

    /// <summary>Maximum imported embedded images.</summary>
    public int MaxImages { get; set; } = 256;

    /// <summary>Maximum decoded bytes in one embedded image.</summary>
    public long MaxImageBytes { get; set; } = 10L * 1024L * 1024L;

    /// <summary>Maximum decoded embedded-image bytes across one import operation.</summary>
    public long MaxTotalImageBytes { get; set; } = 50L * 1024L * 1024L;

    /// <summary>Maximum imported native charts.</summary>
    public int MaxCharts { get; set; } = 1_000;

    /// <summary>Maximum series accepted in one semantic chart.</summary>
    public int MaxChartSeries { get; set; } = 1_000;

    /// <summary>Maximum categories accepted in one semantic chart.</summary>
    public int MaxChartCategories { get; set; } = 10_000;

    /// <summary>Maximum chart data points across one import operation.</summary>
    public long MaxChartPoints { get; set; } = 1_000_000L;

    /// <summary>Maximum imported native shapes across one operation.</summary>
    public int MaxShapes { get; set; } = 10_000;

    /// <summary>Maximum comments, formulas, notes, and similar annotations.</summary>
    public int MaxAnnotations { get; set; } = 100_000;

    /// <summary>Maximum characters accepted by one target-native text or metadata field.</summary>
    public int MaxMetadataCharacters { get; set; } = 1024 * 1024;

    /// <summary>Largest absolute coordinate or size accepted from semantic geometry.</summary>
    public double MaxAbsoluteGeometry { get; set; } = 1_000_000D;

    /// <summary>Creates an independent limits snapshot.</summary>
    public HtmlImportLimits Clone() => new HtmlImportLimits {
        MaxSemanticContainers = MaxSemanticContainers,
        MaxTables = MaxTables,
        MaxTableCells = MaxTableCells,
        MaxImages = MaxImages,
        MaxImageBytes = MaxImageBytes,
        MaxTotalImageBytes = MaxTotalImageBytes,
        MaxCharts = MaxCharts,
        MaxChartSeries = MaxChartSeries,
        MaxChartCategories = MaxChartCategories,
        MaxChartPoints = MaxChartPoints,
        MaxShapes = MaxShapes,
        MaxAnnotations = MaxAnnotations,
        MaxMetadataCharacters = MaxMetadataCharacters,
        MaxAbsoluteGeometry = MaxAbsoluteGeometry
    };

    /// <summary>Validates that every configured budget is finite and positive.</summary>
    public void Validate() {
        Positive(MaxSemanticContainers, nameof(MaxSemanticContainers));
        Positive(MaxTables, nameof(MaxTables));
        Positive(MaxTableCells, nameof(MaxTableCells));
        Positive(MaxImages, nameof(MaxImages));
        Positive(MaxImageBytes, nameof(MaxImageBytes));
        Positive(MaxTotalImageBytes, nameof(MaxTotalImageBytes));
        Positive(MaxCharts, nameof(MaxCharts));
        Positive(MaxChartSeries, nameof(MaxChartSeries));
        Positive(MaxChartCategories, nameof(MaxChartCategories));
        Positive(MaxChartPoints, nameof(MaxChartPoints));
        Positive(MaxShapes, nameof(MaxShapes));
        Positive(MaxAnnotations, nameof(MaxAnnotations));
        Positive(MaxMetadataCharacters, nameof(MaxMetadataCharacters));
        if (double.IsNaN(MaxAbsoluteGeometry) || double.IsInfinity(MaxAbsoluteGeometry) || MaxAbsoluteGeometry <= 0D) {
            throw new ArgumentOutOfRangeException(nameof(MaxAbsoluteGeometry), "Maximum absolute geometry must be finite and positive.");
        }

        if (MaxTotalImageBytes < MaxImageBytes) {
            throw new ArgumentOutOfRangeException(nameof(MaxTotalImageBytes), "The total image byte limit cannot be smaller than the per-image limit.");
        }
    }

    private static void Positive(long value, string name) {
        if (value <= 0L) throw new ArgumentOutOfRangeException(name, "HTML import limits must be positive.");
    }
}

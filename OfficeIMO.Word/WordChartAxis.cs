namespace OfficeIMO.Word;

/// <summary>Category settings used when subsequent series are added to a Word chart.</summary>
/// <remarks>Changing these settings does not rewrite the cached categories of existing series.</remarks>
public sealed class WordChartAxis {
    internal WordChartAxis(WordChart chart) => Chart = chart;
    /// <summary>The owning chart.</summary>
    public WordChart Chart { get; }
    /// <summary>The category labels used by subsequent series.</summary>
    public IReadOnlyList<string> Categories {
        get => Chart.SeriesCategories.AsReadOnly();
        set => Chart.SeriesCategories = new List<string>(value ?? throw new ArgumentNullException(nameof(value)));
    }
}

namespace OfficeIMO;

/// <summary>Shared title and physical dimensions for an editable native chart.</summary>
public class OfficeChartOptions {
    /// <summary>Gets or sets the chart title. Null keeps the format's default title behavior.</summary>
    public string? Title { get; set; }

    /// <summary>Gets or sets the drawing width. Null keeps the format's default width.</summary>
    public OfficeLength? Width { get; set; }

    /// <summary>Gets or sets the drawing height. Null keeps the format's default height.</summary>
    public OfficeLength? Height { get; set; }
}

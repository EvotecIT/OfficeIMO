namespace OfficeIMO.IWork;

/// <summary>Selected character-marker placement and size from a native list style.</summary>
public sealed class IWorkListLayout {
    internal IWorkListLayout(double markerIndentPoints, double textIndentEm, double markerScale) {
        MarkerIndentPoints = markerIndentPoints;
        TextIndentEm = textIndentEm;
        MarkerScale = markerScale;
    }

    /// <summary>Gets the marker's additional indentation in points.</summary>
    public double MarkerIndentPoints { get; }
    /// <summary>Gets the text offset from the marker in multiples of the paragraph font size.</summary>
    public double TextIndentEm { get; }
    /// <summary>Gets the marker size relative to paragraph text; 1 means the same size.</summary>
    public double MarkerScale { get; }
}

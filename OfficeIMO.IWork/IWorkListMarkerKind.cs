namespace OfficeIMO.IWork;

/// <summary>Identifies the native list marker independently of its displayed label.</summary>
public enum IWorkListMarkerKind {
    /// <summary>The paragraph has no list marker.</summary>
    None = 0,
    /// <summary>The source selects an image marker whose appearance is not reconstructed.</summary>
    Image = 1,
    /// <summary>The label is literal text, even when it resembles a number.</summary>
    Text = 2,
    /// <summary>The source selects a numbering format; the label is its initial marker, not a qualified counter.</summary>
    Number = 3
}

namespace OfficeIMO.PowerPoint;

/// <summary>Native DrawingML strike-through variants for PowerPoint text.</summary>
public enum PowerPointStrikeStyle {
    /// <summary>No strike-through.</summary>
    None,
    /// <summary>A single strike-through line.</summary>
    Single,
    /// <summary>A double strike-through line.</summary>
    Double
}

/// <summary>Native DrawingML capitalization variants for PowerPoint text.</summary>
public enum PowerPointCapitalization {
    /// <summary>No capitalization effect.</summary>
    None,
    /// <summary>Small-capital text.</summary>
    SmallCaps,
    /// <summary>All-capital text.</summary>
    AllCaps
}

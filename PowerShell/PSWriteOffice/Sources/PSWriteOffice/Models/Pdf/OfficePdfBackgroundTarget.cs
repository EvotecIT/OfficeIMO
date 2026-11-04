namespace PSWriteOffice.Models.Pdf;

/// <summary>Background decoration cleared by Set-OfficePdfBackground.</summary>
public enum OfficePdfBackgroundTarget {
    /// <summary>All background decorations.</summary>
    All,
    /// <summary>The page color.</summary>
    Color,
    /// <summary>The background image.</summary>
    Image,
    /// <summary>The background shapes.</summary>
    Shapes
}

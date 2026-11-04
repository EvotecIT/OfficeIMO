namespace OfficeIMO.Publisher;

/// <summary>The registration of a master definition, independent of the document page to which it is assigned.</summary>
public enum PublisherMasterPageSide {
    /// <summary>A centered, independent single-page master.</summary>
    Single,
    /// <summary>The left definition of a two-page master, registered at its left edge.</summary>
    Left,
    /// <summary>The right definition of a two-page master, registered at its right edge.</summary>
    Right
}

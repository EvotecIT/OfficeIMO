namespace OfficeIMO.Entities;

/// <summary>Host-supplied coordinates for a text segment.</summary>
public sealed class EntitySegmentLocation {
    /// <summary>Creates a location from the source part, address or page.</summary>
    public EntitySegmentLocation(string? part = null, string? address = null, int? page = null) {
        Part = part;
        Address = address;
        Page = page;
    }

    /// <summary>Source part name, such as a worksheet name.</summary>
    public string? Part { get; }
    /// <summary>Address within the source part, such as a cell coordinate.</summary>
    public string? Address { get; }
    /// <summary>Host-supplied page number.</summary>
    public int? Page { get; }
}

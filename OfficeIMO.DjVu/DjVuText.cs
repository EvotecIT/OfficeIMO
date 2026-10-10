namespace OfficeIMO.DjVu;

/// <summary>The state of a page's stored text layer.</summary>
public enum DjVuTextStatus {
    /// <summary>No stored text chunk exists.</summary>
    Absent,
    /// <summary>A valid layer contains no text.</summary>
    Empty,
    /// <summary>The existing layer contains decoded text.</summary>
    Present,
    /// <summary>A stored layer exists but is malformed.</summary>
    Corrupt
}

/// <summary>DjVu's reading-order text-zone hierarchy.</summary>
public enum DjVuTextZoneKind {
    /// <summary>The page.</summary>
    Page = 1,
    /// <summary>A column.</summary>
    Column = 2,
    /// <summary>A region.</summary>
    Region = 3,
    /// <summary>A paragraph.</summary>
    Paragraph = 4,
    /// <summary>A line.</summary>
    Line = 5,
    /// <summary>A word.</summary>
    Word = 6,
    /// <summary>A character.</summary>
    Character = 7
}

/// <summary>A rectangle in native page pixels. Text and render regions use a bottom-left origin; display bounds use a top-left origin.</summary>
public readonly struct DjVuRectangle {
    /// <summary>Creates a rectangle. Width and height must be nonnegative.</summary>
    public DjVuRectangle(int x, int y, int width, int height) {
        if (width < 0) throw new ArgumentOutOfRangeException(nameof(width));
        if (height < 0) throw new ArgumentOutOfRangeException(nameof(height));
        X = x; Y = y; Width = width; Height = height;
    }
    /// <summary>Left edge in native pixels.</summary>
    public int X { get; }
    /// <summary>Vertical origin in native pixels, using the coordinate system of the owning API.</summary>
    public int Y { get; }
    /// <summary>Width in native pixels.</summary>
    public int Width { get; }
    /// <summary>Height in native pixels.</summary>
    public int Height { get; }
}

/// <summary>A validated reading-order zone in an existing stored text layer.</summary>
public sealed class DjVuTextZone {
    internal DjVuTextZone(DjVuTextZoneKind kind, DjVuRectangle bounds, int byteOffset, int byteLength,
        int characterOffset, int characterLength, List<DjVuTextZone> children) {
        Kind = kind; Bounds = bounds; ByteOffset = byteOffset; ByteLength = byteLength;
        CharacterOffset = characterOffset; CharacterLength = characterLength;
        Children = children.AsReadOnly();
    }
    /// <summary>Granularity of this zone.</summary>
    public DjVuTextZoneKind Kind { get; }
    /// <summary>Unrotated native bounds, using a bottom-left origin.</summary>
    public DjVuRectangle Bounds { get; }
    /// <summary>Offset in the source layer's UTF-8 text bytes.</summary>
    public int ByteOffset { get; }
    /// <summary>Length in the source layer's UTF-8 text bytes.</summary>
    public int ByteLength { get; }
    /// <summary>Offset in the decoded .NET string, in UTF-16 code units.</summary>
    public int CharacterOffset { get; }
    /// <summary>Length in the decoded .NET string, in UTF-16 code units.</summary>
    public int CharacterLength { get; }
    /// <summary>Contained zones in stored reading order.</summary>
    public IReadOnlyList<DjVuTextZone> Children { get; }
}

/// <summary>Stored text and geometry. Existing text may originate from historical OCR.</summary>
public sealed class DjVuTextResult {
    internal DjVuTextResult(DjVuTextStatus status, string text, List<DjVuTextZone>? zones = null, string? diagnostic = null) {
        Status = status; Text = text; Zones = (zones ?? new List<DjVuTextZone>()).AsReadOnly(); Diagnostic = diagnostic;
    }
    /// <summary>Distinguishes missing, empty, present and corrupt layers.</summary>
    public DjVuTextStatus Status { get; }
    /// <summary>Exact decoded stored text, including its original separators.</summary>
    public string Text { get; }
    /// <summary>Root zones in stored reading order.</summary>
    public IReadOnlyList<DjVuTextZone> Zones { get; }
    /// <summary>A decoding failure when the layer is corrupt; otherwise null.</summary>
    public string? Diagnostic { get; }
}

namespace OfficeIMO.Entities;

/// <summary>The portion of a match written in one source segment.</summary>
public readonly struct EntityMatchPart {
    internal EntityMatchPart(int segmentIndex, int start, int length) {
        SegmentIndex = segmentIndex;
        Start = start;
        Length = length;
    }

    /// <summary>Zero-based index in the input segments.</summary>
    public int SegmentIndex { get; }
    /// <summary>UTF-16 offset in that segment.</summary>
    public int Start { get; }
    /// <summary>UTF-16 length in that segment.</summary>
    public int Length { get; }
}

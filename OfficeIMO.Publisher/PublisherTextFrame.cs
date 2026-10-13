using OfficeIMO.Drawing;

namespace OfficeIMO.Publisher;

/// <summary>
/// A recovered native text frame. Frame links and geometry come from the source;
/// assigned text ranges use OfficeIMO's managed layout and can differ from Publisher.
/// </summary>
public sealed class PublisherTextFrame {
    internal PublisherTextFrame(uint id, uint storyId, uint pageId, uint? previous, uint? next, uint order,
        double x, double y, double width, double height, int columns, double columnSpacing,
        int? textStart, int? textLength, bool hasOverflow, IReadOnlyList<uint> wrapObjects, OfficeTransform pageTransform) {
        Id = id; StoryId = storyId; PageId = pageId; PreviousFrameId = previous; NextFrameId = next;
        Order = order; X = x; Y = y; Width = width; Height = height; ColumnCount = columns;
        ColumnSpacing = columnSpacing; TextStart = textStart; TextLength = textLength; HasOverflow = hasOverflow;
        WrapObjectIds = Array.AsReadOnly(wrapObjects.ToArray());
        PageTransform = pageTransform;
    }
    /// <summary>Native publication object identifier.</summary>
    public uint Id { get; }
    /// <summary>Identifier of the complete source story in <see cref="PublisherDocument.TextStories"/>.</summary>
    public uint StoryId { get; }
    /// <summary>Owning document or master page identifier.</summary>
    public uint PageId { get; }
    /// <summary>Previous native frame, or null for the first frame.</summary>
    public uint? PreviousFrameId { get; }
    /// <summary>Next native frame, or null for the last frame.</summary>
    public uint? NextFrameId { get; }
    /// <summary>Zero-based native order within the linked story.</summary>
    public uint Order { get; }
    /// <summary>Left edge in page-local points, before object and enclosing-group rotation or reflection.</summary>
    public double X { get; }
    /// <summary>Top edge in page-local points, before object and enclosing-group rotation or reflection.</summary>
    public double Y { get; }
    /// <summary>Frame width in points.</summary>
    public double Width { get; }
    /// <summary>Frame height in points.</summary>
    public double Height { get; }
    /// <summary>Maps frame-local points into page-local points, including object and enclosing-group rotation/reflection.</summary>
    public OfficeTransform PageTransform { get; }
    /// <summary>Native number of equal-width text columns.</summary>
    public int ColumnCount { get; }
    /// <summary>Gap between text columns in points.</summary>
    public double ColumnSpacing { get; }
    /// <summary>Native objects referenced by this frame's text-wrap exclusion list.</summary>
    public IReadOnlyList<uint> WrapObjectIds { get; }
    /// <summary>UTF-16 start in the normalized source story, or null when placement is unresolved.</summary>
    public int? TextStart { get; }
    /// <summary>Assigned UTF-16 length, including paragraph separators; null when placement is unresolved.</summary>
    public int? TextLength { get; }
    /// <summary>Whether story content remains after the final recovered frame.</summary>
    public bool HasOverflow { get; }
}

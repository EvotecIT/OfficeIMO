using OfficeIMO.Drawing;

namespace OfficeIMO.Publisher;

/// <summary>A recovered publication page. Coordinates and dimensions are points, with a local top-left origin.</summary>
public sealed class PublisherPage {
    internal PublisherPage(uint id, string name, uint? masterPageId, OfficeDrawing drawing,
        IReadOnlyList<PublisherTextFrame> textFrames, IReadOnlyList<PublisherTable> tables) {
        Id = id; Name = name; MasterPageId = masterPageId; Drawing = drawing;
        TextFrames = Array.AsReadOnly(textFrames.ToArray());
        Tables = Array.AsReadOnly(tables.ToArray());
    }
    /// <summary>Native page identifier.</summary>
    public uint Id { get; }
    /// <summary>Native page or master name; empty when unnamed.</summary>
    public string Name { get; }
    /// <summary>Applied master page identifier, when declared in the source.</summary>
    public uint? MasterPageId { get; }
    /// <summary>Recovered drawing scene, including the applied master. Changes affect exported graphics, not the original .pub bytes.</summary>
    public OfficeDrawing Drawing { get; }
    /// <summary>Native text frames owned by this page. Frames inherited from a master are listed on that master page.</summary>
    public IReadOnlyList<PublisherTextFrame> TextFrames { get; }
    /// <summary>Native tables owned by this page. Tables inherited from a master are listed on that master page.</summary>
    public IReadOnlyList<PublisherTable> Tables { get; }
    /// <summary>Page width in points.</summary>
    public double Width => Drawing.Width;
    /// <summary>Page height in points.</summary>
    public double Height => Drawing.Height;
}

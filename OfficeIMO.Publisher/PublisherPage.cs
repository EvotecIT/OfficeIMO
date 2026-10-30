using OfficeIMO.Drawing;

namespace OfficeIMO.Publisher;

/// <summary>A recovered publication page. Coordinates and dimensions are points, with a local top-left origin.</summary>
public sealed class PublisherPage {
    private readonly OfficeDrawing _drawing;
    internal PublisherPage(uint id, string name, uint? masterPageId, double width, double height, OfficeDrawing drawing,
        IReadOnlyList<PublisherTextFrame> textFrames, IReadOnlyList<PublisherTable> tables, IReadOnlyList<PublisherPicture> pictures,
        IReadOnlyList<PublisherShape> shapes) {
        Id = id; Name = name; MasterPageId = masterPageId; _drawing = drawing;
        Width = width; Height = height;
        TextFrames = Array.AsReadOnly(textFrames.ToArray());
        Tables = Array.AsReadOnly(tables.ToArray());
        Pictures = Array.AsReadOnly(pictures.ToArray());
        Shapes = Array.AsReadOnly(shapes.ToArray());
    }
    /// <summary>Native page identifier.</summary>
    public uint Id { get; }
    /// <summary>Native page or master name; empty when unnamed.</summary>
    public string Name { get; }
    /// <summary>Applied master page identifier, when declared in the source.</summary>
    public uint? MasterPageId { get; }
    /// <summary>Creates a detached drawing snapshot, including the applied master. Changes to this copy do not edit the publication.</summary>
    public OfficeDrawing CreateDrawing() => _drawing.Clone();
    /// <summary>Current text frames owned by this page, including authored continuation frames. Frames inherited from a master are listed on that master page.</summary>
    public IReadOnlyList<PublisherTextFrame> TextFrames { get; }
    /// <summary>Current tables owned by this page, including authored tables. Tables inherited from a master are listed on that master page.</summary>
    public IReadOnlyList<PublisherTable> Tables { get; }
    /// <summary>Picture placements owned by this page, in recovered drawing order. Pictures inherited from a master are listed on that master.</summary>
    public IReadOnlyList<PublisherPicture> Pictures { get; }
    /// <summary>Vector artwork owned by this page, in drawing order. Text-capable artwork shares its ID and placement with its text frame; inherited artwork remains listed on the master.</summary>
    public IReadOnlyList<PublisherShape> Shapes { get; }
    /// <summary>Page width in points.</summary>
    public double Width { get; }
    /// <summary>Page height in points.</summary>
    public double Height { get; }
}

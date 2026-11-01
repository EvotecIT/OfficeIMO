using OfficeIMO.Drawing;

namespace OfficeIMO.Publisher;

/// <summary>An immutable current document or master page. Coordinates and dimensions are points, with a local top-left origin.</summary>
public sealed class PublisherPage {
    private readonly OfficeDrawing _drawing;
    internal PublisherPage(uint id, string name, uint? masterPageId, double width, double height, OfficeDrawing drawing,
        IReadOnlyList<PublisherTextFrame> textFrames, IReadOnlyList<PublisherTable> tables, IReadOnlyList<PublisherPicture> pictures,
        IReadOnlyList<PublisherShape> shapes, IReadOnlyList<uint> objectIds, bool hasNativeOrigin = true, PublisherMasterPageSide? masterSide = null) {
        Id = id; Name = name; MasterPageId = masterPageId; _drawing = drawing;
        Width = width; Height = height;
        HasNativeOrigin = hasNativeOrigin;
        MasterSide = masterSide;
        TextFrames = Array.AsReadOnly(textFrames.ToArray());
        Tables = Array.AsReadOnly(tables.ToArray());
        Pictures = Array.AsReadOnly(pictures.ToArray());
        Shapes = Array.AsReadOnly(shapes.ToArray());
        ObjectIds = Array.AsReadOnly(objectIds.ToArray());
    }
    /// <summary>Stable publication page identifier, shared with the publication's object and story identity space.</summary>
    public uint Id { get; }
    /// <summary>Whether this page definition was recovered from the native source. Adding content or changing its applied master retains this provenance.</summary>
    public bool HasNativeOrigin { get; }
    /// <summary>Page or master name; empty when unnamed.</summary>
    public string Name { get; }
    /// <summary>Current applied master page identifier, or null when none applies, including native pages that suppress master formatting. A suppressed source can retain a latent native reference separately.</summary>
    public uint? MasterPageId { get; }
    /// <summary>Registration of this master definition; null for document pages or an unqualified native master layout. Applied positions can differ from definition coordinates for left/right masters.</summary>
    public PublisherMasterPageSide? MasterSide { get; }
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
    /// <summary>Owned object identifiers in back-to-front drawing order. Text-capable artwork appears once even when it also has a text-frame view. Applied master objects remain on the master page.</summary>
    public IReadOnlyList<uint> ObjectIds { get; }
    /// <summary>Page width in points.</summary>
    public double Width { get; }
    /// <summary>Page height in points.</summary>
    public double Height { get; }
}

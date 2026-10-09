using OfficeIMO.Drawing;
using OfficeIMO.Drawing.Binary;

namespace OfficeIMO.Publisher.Internal;

internal sealed partial class PublisherPageProjector {
    private readonly PublisherContents _source;
    private readonly PublisherQuillText _text;
    private readonly PublisherEscherData _escher;
    private readonly PublisherParseContext _context;
    private readonly HashSet<uint> _projected = new();
    private readonly Dictionary<uint, uint?> _owners = new();
    private readonly HashSet<uint> _usedStories = new();
    private readonly Dictionary<uint, PublisherSourcePage> _pages;
    internal PublisherPageProjector(PublisherContents source, PublisherQuillText text, PublisherEscherData escher, PublisherParseContext context) {
        _source = source; _text = text; _escher = escher; _context = context;
        _pages = source.Pages.ToDictionary(page => page.Id);
    }
    internal PublisherDocument Project() {
        var scenes = new Dictionary<uint, OfficeDrawing>();
        var pageShapes = new Dictionary<uint, List<PublisherEscherShape>>();
        foreach (PublisherEscherShape shape in _escher.Shapes) {
            _context.Token.ThrowIfCancellationRequested();
            uint? owner = Owner(shape.Id, 0);
            if (!owner.HasValue) continue;
            if (!pageShapes.TryGetValue(owner.Value, out List<PublisherEscherShape>? shapes))
                pageShapes.Add(owner.Value, shapes = new List<PublisherEscherShape>());
            shapes.Add(shape);
        }
        foreach (PublisherSourcePage page in _source.Pages) {
            _context.Token.ThrowIfCancellationRequested();
            var drawing = new OfficeDrawing(_source.Width, _source.Height);
            if (pageShapes.TryGetValue(page.Id, out List<PublisherEscherShape>? shapes))
                foreach (PublisherEscherShape shape in shapes) ProjectShape(shape, drawing);
            scenes.Add(page.Id, drawing);
        }
        var masters = new List<PublisherPage>();
        var pages = new List<PublisherPage>();
        foreach (PublisherSourcePage page in _source.Pages) {
            _context.Token.ThrowIfCancellationRequested();
            var drawing = new OfficeDrawing(_source.Width, _source.Height);
            if (!page.IsMaster) drawing.AddShape(new OfficeShape { Kind = OfficeShapeKind.Rectangle, Width = drawing.Width, Height = drawing.Height, FillColor = OfficeColor.White }, 0, 0);
            if (page.Master.HasValue && page.Master.Value != 0) {
                if (_pages.TryGetValue(page.Master.Value, out PublisherSourcePage? master) && master.IsMaster) {
                    _context.AccountProjection(scenes[master.Id]);
                    drawing.AddDrawingForClippedRendering(scenes[master.Id], 0, 0, null);
                }
                else _context.Add("PUB_MASTER_REFERENCE_UNRESOLVED", "A declared master page could not be applied.", OfficeConversionLossKind.Omission, "Contents/page/" + page.Id);
            }
            if (page.Background.GetValueOrDefault() != 0)
                _context.Add("PUB_PAGE_BACKGROUND_UNASSESSED", "A native page background definition was not reconstructed separately from its page artwork.",
                    OfficeConversionLossKind.Unassessed, "Contents/page/" + page.Id);
            if (page.IsMaster && page.Master.GetValueOrDefault() != 0)
                _context.Add("PUB_NESTED_MASTER_UNASSESSED", "Master-page inheritance beyond the directly referenced definition is not reconstructed on document pages.",
                    OfficeConversionLossKind.Unassessed, "Contents/page/" + page.Id);
            _context.AccountProjection(scenes[page.Id]);
            drawing.AddDrawingForClippedRendering(scenes[page.Id], 0, 0, null);
            var clipped = new OfficeDrawing(drawing.Width, drawing.Height);
            clipped.AddClippedDrawing(drawing, 0, 0, OfficeClipPath.Rectangle(drawing.Width, drawing.Height));
            var projectedPage = new PublisherPage(page.Id, page.Name, page.Master, clipped);
            if (page.IsMaster) masters.Add(projectedPage); else pages.Add(projectedPage);
        }
        if (pages.Count == 0) throw new InvalidDataException("Publisher publication has no printable document pages.");
        foreach (PublisherSourceShape shape in _source.Shapes.Values) {
            if (!_projected.Contains(shape.Chunk.Id)) _context.Add("PUB_OBJECT_NOT_PROJECTED",
                "A native publication object was not reconstructed on a document or master page.", OfficeConversionLossKind.Omission, PublisherEscherReader.ShapeLocation(shape.Chunk.Id));
        }
        foreach (PublisherTextStory story in _text.Stories.Values) {
            if (!_usedStories.Contains(story.Id)) _context.Add("PUB_STORY_NOT_PLACED", "Recovered story text is available in TextStories, but was not placed on a publication page.",
                OfficeConversionLossKind.Omission, "Quill/story/" + story.Id);
        }
        var report = new PublisherReadReport(_context.Diagnostics, _source.Shapes.Count, _projected.Count,
            _text.Stories.Count, _escher.Images.Count, _context.Records);
        return new PublisherDocument(pages, masters, _text.Stories.Values.ToArray(), _escher.Images.Values.ToArray(), report);
    }
    private uint? Owner(uint id, int depth) {
        _context.CheckDepth(depth);
        if (_owners.TryGetValue(id, out uint? owner)) return owner;
        if (_source.ObjectPages.TryGetValue(id, out uint page)) owner = page;
        else if (_source.Shapes.TryGetValue(id, out PublisherSourceShape? shape) && shape.Chunk.Parent.HasValue) {
            uint parent = shape.Chunk.Parent.Value;
            owner = _pages.ContainsKey(parent) ? parent : Owner(parent, depth + 1);
        }
        _owners[id] = owner;
        return owner;
    }
    private void ProjectShape(PublisherEscherShape shape, OfficeDrawing page) {
        _context.Token.ThrowIfCancellationRequested();
        if (!_source.Shapes.TryGetValue(shape.Id, out PublisherSourceShape? native)) {
            _context.Add("PUB_DRAWING_REFERENCE_UNRESOLVED", "An OfficeArt object has no publication descriptor.", OfficeConversionLossKind.Unassessed, PublisherEscherReader.ShapeLocation(shape.Id));
            return;
        }
        if (shape.IsGroup) { _projected.Add(shape.Id); return; }
        if (shape.Style.Hidden == true) { _projected.Add(shape.Id); return; }
        if (!shape.Bounds.HasValue) {
            _context.Add("PUB_OBJECT_ANCHOR_UNRESOLVED", "A publication object has no resolved native anchor.", OfficeConversionLossKind.Omission, PublisherEscherReader.ShapeLocation(shape.Id)); return;
        }
        PublisherNativeRectangle bounds = shape.Bounds.Value;
        double x = Math.Min(bounds.X1, bounds.X2) / 12700 + page.Width / 2;
        double y = Math.Min(bounds.Y1, bounds.Y2) / 12700 + page.Height / 2;
        double width = Math.Abs(bounds.X2 - bounds.X1) / 12700, height = Math.Abs(bounds.Y2 - bounds.Y1) / 12700;
        double rotation = shape.Transform.RotationDegrees.GetValueOrDefault();
        double angle = ((rotation % 360) + 360) % 360;
        if (angle is >= 45 and < 135 or >= 225 and < 315) { x += (width - height) / 2; y += (height - width) / 2; (width, height) = (height, width); }
        double geometryWidth = width, geometryHeight = height;
        if (shape.Type == 20 && (width > 0 || height > 0)) { width = Math.Max(width, 0.01); height = Math.Max(height, 0.01); }
        if (width <= 0 || height <= 0) {
            _context.Add("PUB_DEGENERATE_OBJECT_OMITTED", "A zero-sized drawing object was omitted.", OfficeConversionLossKind.Omission, PublisherEscherReader.ShapeLocation(shape.Id)); return;
        }
        if (!FiniteRectangle(x, y, width, height)) throw new InvalidDataException("Publisher group projection produced invalid geometry.");
        if (x < 0 || y < 0 || x + width > page.Width || y + height > page.Height)
            _context.Add("PUB_PAGE_BLEED_CLIPPED", "Artwork outside the publication page remains in the scene but is clipped to the printable page in exports.", OfficeConversionLossKind.Approximation, PublisherEscherReader.ShapeLocation(shape.Id));
        var local = new OfficeDrawing(width, height);
        ProjectGraphic(shape, local, geometryWidth, geometryHeight);
        uint? storyId = native.Value(0x27);
        if (storyId.HasValue) {
            if (_text.Stories.TryGetValue(storyId.Value, out PublisherTextStory? story)) {
                if (_usedStories.Add(storyId.Value)) {
                    uint vertical = native.Value(0x35) ?? 0;
                    if (native.Chunk.Kind == 0x10) ProjectTable(native, storyId.Value, shape, local);
                    else if (story.Paragraphs.Count != 0) ProjectText(story.Paragraphs, local, shape, 0, 0, width, height,
                        vertical switch { 1 => OfficeTextVerticalAlignment.Center, 2 => OfficeTextVerticalAlignment.Bottom, _ => OfficeTextVerticalAlignment.Top });
                    _context.Add("PUB_TEXT_LAYOUT_APPROXIMATED", "Native text stories use the shared drawing paragraph layout. Font availability, line breaking, fitting, and overflow may differ from Publisher.",
                        OfficeConversionLossKind.Approximation, "Quill");
                } else _context.Add("PUB_LINKED_TEXT_FRAME_OMITTED", "The complete linked story is placed in its first recovered frame; continuation flow is not reconstructed. TextStories retains the complete text.",
                    OfficeConversionLossKind.Omission, PublisherEscherReader.ShapeLocation(shape.Id));
            } else _context.Add("PUB_TEXT_REFERENCE_UNRESOLVED", "A text frame refers to an unavailable story.", OfficeConversionLossKind.Omission, PublisherEscherReader.ShapeLocation(shape.Id));
        }
        foreach (OfficeDrawingElement element in local.Elements) element.SourceElementIds = new[] { "publisher-object-" + shape.Id };
        var transform = new OfficeImageFrameTransform(rotation, x + width / 2, y + height / 2, shape.Transform.FlipHorizontal, shape.Transform.FlipVertical);
        _context.AccountProjection(local);
        page.AddDrawingForClippedRendering(local, x, y, transform);
        if (local.Elements.Count > 0) _projected.Add(shape.Id);
    }
    private OfficeTextPadding TextPadding(PublisherEscherShape shape) => new(
        (shape.Property(0x81) ?? 36576) / 12700D, (shape.Property(0x82) ?? 36576) / 12700D,
        (shape.Property(0x83) ?? 36576) / 12700D, (shape.Property(0x84) ?? 36576) / 12700D);
    private static bool FiniteRectangle(double x, double y, double width, double height) =>
        !double.IsNaN(x) && !double.IsInfinity(x) && !double.IsNaN(y) && !double.IsInfinity(y)
        && !double.IsNaN(width) && !double.IsInfinity(width) && !double.IsNaN(height) && !double.IsInfinity(height);
}

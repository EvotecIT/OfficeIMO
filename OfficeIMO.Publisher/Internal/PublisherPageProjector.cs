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
    private readonly Dictionary<uint, PublisherTable> _tables = new();
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
        PrepareTextFrames();
        var pageFrames = new Dictionary<uint, List<PublisherTextFrame>>();
        foreach (PreparedTextFrame frame in _textFrames.Values) {
            _context.Token.ThrowIfCancellationRequested();
            if (!pageFrames.TryGetValue(frame.Model.PageId, out List<PublisherTextFrame>? frames))
                pageFrames.Add(frame.Model.PageId, frames = new List<PublisherTextFrame>());
            frames.Add(frame.Model);
        }
        foreach (PublisherSourcePage page in _source.Pages) {
            _context.Token.ThrowIfCancellationRequested();
            var drawing = new OfficeDrawing(_source.Width, _source.Height);
            if (pageShapes.TryGetValue(page.Id, out List<PublisherEscherShape>? shapes))
                foreach (PublisherEscherShape shape in shapes) ProjectShape(shape, drawing, page.Id);
            scenes.Add(page.Id, drawing);
        }
        var pageTables = _tables.Values.GroupBy(table => table.PageId).ToDictionary(group => group.Key, group => group.ToArray());
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
            var projectedPage = new PublisherPage(page.Id, page.Name, page.Master, clipped,
                pageFrames.TryGetValue(page.Id, out List<PublisherTextFrame>? frames) ? frames : Array.Empty<PublisherTextFrame>(),
                pageTables.TryGetValue(page.Id, out PublisherTable[]? tables) ? tables : Array.Empty<PublisherTable>());
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
    private void ProjectShape(PublisherEscherShape shape, OfficeDrawing page, uint pageId) {
        _context.Token.ThrowIfCancellationRequested();
        if (!_source.Shapes.TryGetValue(shape.Id, out PublisherSourceShape? native)) {
            _context.Add("PUB_DRAWING_REFERENCE_UNRESOLVED", "An OfficeArt object has no publication descriptor.", OfficeConversionLossKind.Unassessed, PublisherEscherReader.ShapeLocation(shape.Id));
            return;
        }
        if (shape.IsGroup) { _projected.Add(shape.Id); return; }
        if (shape.Hidden) { _projected.Add(shape.Id); return; }
        if (!shape.Bounds.HasValue) {
            _context.Add("PUB_OBJECT_ANCHOR_UNRESOLVED", "A publication object has no resolved native anchor.", OfficeConversionLossKind.Omission, PublisherEscherReader.ShapeLocation(shape.Id)); return;
        }
        FrameRectangle bounds = FrameBounds(shape, page.Width, page.Height);
        double x = bounds.X, y = bounds.Y, width = bounds.Width, height = bounds.Height;
        double rotation = shape.Transform.RotationDegrees.GetValueOrDefault();
        double geometryWidth = width, geometryHeight = height;
        if (shape.Type == 20 && (width > 0 || height > 0)) { width = Math.Max(width, 0.01); height = Math.Max(height, 0.01); }
        if (width <= 0 || height <= 0) {
            _context.Add("PUB_DEGENERATE_OBJECT_OMITTED", "A zero-sized drawing object was omitted.", OfficeConversionLossKind.Omission, PublisherEscherReader.ShapeLocation(shape.Id)); return;
        }
        if (!FiniteRectangle(x, y, width, height)) throw new InvalidDataException("Publisher group projection produced invalid geometry.");
        var printableBounds = PageTransform(shape, bounds).TransformRectangleBounds(x, y, width, height);
        if (printableBounds.Left < 0 || printableBounds.Top < 0 || printableBounds.Right > page.Width || printableBounds.Bottom > page.Height)
            _context.Add("PUB_PAGE_BLEED_CLIPPED", "Artwork outside the publication page remains in the scene but is clipped to the printable page in exports.", OfficeConversionLossKind.Approximation, PublisherEscherReader.ShapeLocation(shape.Id));
        var local = new OfficeDrawing(width, height);
        ProjectGraphic(shape, local, geometryWidth, geometryHeight);
        uint? storyId = native.Value(0x27);
        if (native.Chunk.Kind == 0x10) ProjectTable(native, storyId, shape, bounds, pageId, local);
        else if (storyId.HasValue) {
            if (_text.Stories.TryGetValue(storyId.Value, out PublisherTextStory? story)) {
                if (_textFrames.TryGetValue(shape.Id, out PreparedTextFrame? frame)) {
                    uint vertical = native.Value(0x35) ?? 0;
                    foreach (PreparedTextRegion region in frame.Regions) local.AddRichTextParagraphs(region.Paragraphs,
                        region.X, region.Y, region.Width, region.Height,
                        verticalAlignment: vertical switch { 1 => OfficeTextVerticalAlignment.Center, 2 => OfficeTextVerticalAlignment.Bottom, _ => OfficeTextVerticalAlignment.Top });
                }
                _context.Add("PUB_TEXT_LAYOUT_APPROXIMATED", "Native text stories use the shared drawing paragraph layout. Font availability, line breaking, fitting, and overflow may differ from Publisher.",
                    OfficeConversionLossKind.Approximation, "Quill");
            } else _context.Add("PUB_TEXT_REFERENCE_UNRESOLVED", "A text frame refers to an unavailable story.", OfficeConversionLossKind.Omission, PublisherEscherReader.ShapeLocation(shape.Id));
        }
        TagSourceElements(local, new[] { "publisher-object-" + shape.Id });
        var transform = new OfficeImageFrameTransform(rotation, x + width / 2, y + height / 2, shape.Transform.FlipHorizontal, shape.Transform.FlipVertical);
        _context.AccountProjection(local);
        if (shape.GroupTransform == OfficeTransform.Identity) page.AddDrawingForClippedRendering(local, x, y, transform);
        else {
            var placed = new OfficeDrawing(page.Width, page.Height);
            placed.AddDrawingForClippedRendering(local, x, y, transform);
            OfficeTransform groupTransform = GroupPageTransform(shape);
            if (!groupTransform.TryInvert(out OfficeTransform inverse))
                throw new InvalidDataException("Publisher group projection produced a singular transform.");
            var visible = inverse.TransformRectangleBounds(0, 0, page.Width, page.Height);
            // Retain pre-transform paint that can enter the final page viewport.
            // This expands vector coordinates only; renderers enforce their own
            // pixel budgets before allocating the intermediate raster surface.
            if (!placed.TryExpandViewportCanvas(visible.Left, visible.Top, visible.Right, visible.Bottom,
                    double.MaxValue, double.MaxValue, out OfficeDrawing retained, out double left, out double top))
                throw new InvalidDataException("Publisher group projection produced invalid viewport geometry.");
            page.AddEffectDrawing(retained, OfficeTransform.Translate(left, top).Then(groupTransform));
        }
        if (local.Elements.Count > 0 || _textFrames.ContainsKey(shape.Id) || _tables.ContainsKey(shape.Id)) _projected.Add(shape.Id);
    }
    private static void TagSourceElements(OfficeDrawing drawing, IReadOnlyList<string> ids) {
        foreach (OfficeDrawingElement element in drawing.Elements) {
            element.SourceElementIds = ids;
            if (element is OfficeDrawingGroup group) TagSourceElements(group.InnerDrawing, ids);
            if (element is OfficeDrawingEffectGroup transformed) TagSourceElements(transformed.InnerDrawing, ids);
        }
    }
    private OfficeTextPadding TextPadding(PublisherEscherShape shape) => new(
        (shape.Property(0x81) ?? 36576) / 12700D, (shape.Property(0x82) ?? 36576) / 12700D,
        (shape.Property(0x83) ?? 36576) / 12700D, (shape.Property(0x84) ?? 36576) / 12700D);
    private static bool FiniteRectangle(double x, double y, double width, double height) =>
        !double.IsNaN(x) && !double.IsInfinity(x) && !double.IsNaN(y) && !double.IsInfinity(y)
        && !double.IsNaN(width) && !double.IsInfinity(width) && !double.IsNaN(height) && !double.IsInfinity(height);
}

using OfficeIMO.Drawing;
using OfficeIMO.Drawing.Binary;

namespace OfficeIMO.Publisher.Internal;

internal sealed partial class PublisherPageProjector {
    private readonly PublisherContents _source;
    private readonly PublisherQuillText _text;
    private readonly PublisherEscherData _escher;
    private readonly PublisherParseContext _context;
    private readonly HashSet<uint> _projected = new();
    private readonly HashSet<uint> _usedStories = new();
    private readonly Dictionary<uint, PublisherTable> _tables = new();
    private readonly PublisherPublicationState _publication;
    private readonly Dictionary<uint, PublisherPageState> _pages;
    internal PublisherPageProjector(PublisherContents source, PublisherQuillText text, PublisherEscherData escher, PublisherParseContext context,
        PublisherPublicationState? publication = null) {
        _source = source; _text = text; _escher = escher; _context = context;
        _publication = publication ?? PublisherModelImporter.Import(source, escher, context);
        _pages = _publication.Pages.ToDictionary(page => page.Id);
        foreach (PublisherPageState page in _publication.Pages.Where(page => !page.SupportsPlacementEditing))
            _context.Add("PUB_MASTER_COORDINATE_ORIGIN_UNASSESSED", "An unassessed master-page layout uses the centered page coordinate origin.",
                OfficeConversionLossKind.Unassessed, "Contents/page/" + page.Id);
        foreach (PublisherPicture picture in _publication.Pictures.Values.Where(picture => !picture.CanEdit))
            _context.Add("PUB_PICTURE_EDIT_UNSUPPORTED", picture.EditLimitation!, OfficeConversionLossKind.None,
                PublisherEscherReader.ShapeLocation(picture.Id));
    }
    internal PublisherDocument Project() {
        var scenes = new Dictionary<uint, OfficeDrawing>();
        var artwork = _escher.Shapes.ToDictionary(shape => shape.Id);
        var tableShapes = _escher.Shapes.Where(shape => _source.Shapes.TryGetValue(shape.Id, out PublisherSourceShape? native)
            && native.Table != null).ToDictionary(shape => shape.Id);
        foreach (PublisherSourceShape native in _source.Shapes.Values.Where(native => native.Table != null)) {
            _context.Token.ThrowIfCancellationRequested();
            uint? owner = Owner(native.Chunk.Id, 0);
            if (!owner.HasValue) continue;
            tableShapes.TryGetValue(native.Chunk.Id, out PublisherEscherShape? shape);
            ProjectTable(native, native.Value(0x27), shape,
                shape?.Bounds.HasValue == true ? FrameBounds(shape) : default, owner.Value, null);
        }
        PrepareTextFrames();
        var pageFrames = new Dictionary<uint, List<PublisherTextFrame>>();
        foreach (PreparedTextFrame frame in _textFrames.Values) {
            _context.Token.ThrowIfCancellationRequested();
            if (!pageFrames.TryGetValue(frame.Model.PageId, out List<PublisherTextFrame>? frames))
                pageFrames.Add(frame.Model.PageId, frames = new List<PublisherTextFrame>());
            frames.Add(frame.Model);
        }
        foreach (PublisherPageState page in _publication.Pages) {
            _context.Token.ThrowIfCancellationRequested();
            var drawing = new OfficeDrawing(page.Width, page.Height);
            foreach (uint id in page.ObjectIds) {
                _context.Token.ThrowIfCancellationRequested();
                if (_publication.Pictures.TryGetValue(id, out PublisherPicture? picture) && picture.CanEdit)
                    ProjectSemanticPicture(picture, drawing);
                else if (artwork.TryGetValue(id, out PublisherEscherShape? shape)) ProjectShape(shape, drawing, page.Id);
            }
            scenes.Add(page.Id, drawing);
        }
        var pageTables = _tables.Values.GroupBy(table => table.PageId).ToDictionary(group => group.Key, group => group.ToArray());
        var masters = new List<PublisherPage>();
        var pages = new List<PublisherPage>();
        foreach (PublisherPageState page in _publication.Pages) {
            _context.Token.ThrowIfCancellationRequested();
            var drawing = new OfficeDrawing(page.Width, page.Height);
            if (!page.IsMaster) drawing.AddShape(new OfficeShape { Kind = OfficeShapeKind.Rectangle, Width = drawing.Width, Height = drawing.Height, FillColor = OfficeColor.White }, 0, 0);
            if (page.MasterPageId.HasValue && page.MasterPageId.Value != 0) {
                if (_pages.TryGetValue(page.MasterPageId.Value, out PublisherPageState? master) && master.IsMaster) {
                    _context.AccountProjection(scenes[master.Id]);
                    // Definition coordinates and an applied master's page basis
                    // differ for the left and right members of a paired master.
                    var targetOrigin = PageOrigin(page);
                    var masterOrigin = PageOrigin(master);
                    drawing.AddDrawingForClippedRendering(scenes[master.Id],
                        targetOrigin.X - masterOrigin.X, targetOrigin.Y - masterOrigin.Y, null);
                }
                else _context.Add("PUB_MASTER_REFERENCE_UNRESOLVED", "A declared master page could not be applied.", OfficeConversionLossKind.Omission, "Contents/page/" + page.Id);
            }
            if (page.HasUnassessedBackground)
                _context.Add("PUB_PAGE_BACKGROUND_UNASSESSED", "A native page background definition was not reconstructed separately from its page artwork.",
                    OfficeConversionLossKind.Unassessed, "Contents/page/" + page.Id);
            if (page.IsMaster && page.MasterPageId.GetValueOrDefault() != 0)
                _context.Add("PUB_NESTED_MASTER_UNASSESSED", "Master-page inheritance beyond the directly referenced definition is not reconstructed on document pages.",
                    OfficeConversionLossKind.Unassessed, "Contents/page/" + page.Id);
            _context.AccountProjection(scenes[page.Id]);
            drawing.AddDrawingForClippedRendering(scenes[page.Id], 0, 0, null);
            var clipped = new OfficeDrawing(drawing.Width, drawing.Height);
            clipped.AddClippedDrawing(drawing, 0, 0, OfficeClipPath.Rectangle(drawing.Width, drawing.Height));
            var projectedPage = new PublisherPage(page.Id, page.Name, page.MasterPageId, page.Width, page.Height, clipped,
                pageFrames.TryGetValue(page.Id, out List<PublisherTextFrame>? frames) ? frames : Array.Empty<PublisherTextFrame>(),
                pageTables.TryGetValue(page.Id, out PublisherTable[]? tables) ? tables : Array.Empty<PublisherTable>(),
                page.ObjectIds.Where(id => _publication.Pictures.ContainsKey(id)).Select(id => _publication.Pictures[id]).ToArray());
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
            _text.Stories.Count, _publication.Images.Count, _context.Records);
        return new PublisherDocument(pages, masters, _text.Stories.Values.ToArray(), _publication.Images.Values.ToArray(), report);
    }
    private uint? Owner(uint id, int depth) {
        _context.CheckDepth(depth);
        return _publication.ObjectOwners.TryGetValue(id, out uint? owner) ? owner : null;
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
        FrameRectangle bounds = FrameBounds(shape);
        double x = bounds.X, y = bounds.Y, width = bounds.Width, height = bounds.Height;
        double rotation = shape.Transform.RotationDegrees.GetValueOrDefault();
        double geometryWidth = width, geometryHeight = height;
        if (!shape.LayoutTransform.TryInvert(out _) || !shape.GroupTransform.TryInvert(out _)) {
            _context.Add("PUB_DEGENERATE_OBJECT_OMITTED", "An object in a collapsed group was omitted.", OfficeConversionLossKind.Omission, PublisherEscherReader.ShapeLocation(shape.Id)); return;
        }
        if ((shape.Type == 20 || OfficeArtCustomPathProjector.HasPath(shape.Properties)) && (width > 0 || height > 0)) { width = Math.Max(width, 0.01); height = Math.Max(height, 0.01); }
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
        if (local.Elements.Count == 0) {
            // A valid empty frame remains recovered without a retained raster surface.
            _projected.Add(shape.Id);
            return;
        }
        TagSourceElements(local, new[] { "publisher-object-" + shape.Id });
        var transform = new OfficeImageFrameTransform(rotation, x + width / 2, y + height / 2, shape.Transform.FlipHorizontal, shape.Transform.FlipVertical);
        _context.AccountProjection(local);
        if (shape.GroupTransform == shape.LayoutTransform && (rotation == 0 || shape.LayoutTransform.M11 == shape.LayoutTransform.M22))
            page.AddDrawingForClippedRendering(local, x, y, transform);
        else {
            OfficeTransform groupTransform = GroupPageTransform(shape);
            if (!groupTransform.TryInvert(out _))
                throw new InvalidDataException("Publisher group projection produced a singular transform.");
            var paint = GroupChildPaintBounds(local);
            // Retain this child's paint before transforming it onto the page.
            // A page-sized intermediate per child needlessly exhausts raster budgets.
            if (!local.TryExpandViewportCanvas(paint.Left, paint.Top, paint.Right, paint.Bottom,
                    double.MaxValue, double.MaxValue, out OfficeDrawing retained, out double left, out double top))
                throw new InvalidDataException("Publisher group projection produced invalid viewport geometry.");
            page.AddEffectDrawing(retained, OfficeTransform.Translate(left + x, top + y)
                .Then(PageTransform(shape, bounds)));
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

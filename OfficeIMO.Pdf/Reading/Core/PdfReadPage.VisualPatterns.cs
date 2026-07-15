using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

public sealed partial class PdfReadPage {
    private static void AddTilingPatternFill(OfficeDrawing drawing, PdfPageVisualPrimitive primitive) {
        PdfPageTilingPatternPaint paint = primitive.FillTilingPattern!;
        if (primitive.Width <= 0D || primitive.Height <= 0D) return;
        PdfPageClipPath shapeClip;
        if (primitive.Kind == PdfPageVisualPrimitiveKind.Rectangle) {
            shapeClip = PdfPageClipPath.Rectangle(primitive.X, primitive.Y, primitive.Width, primitive.Height);
        } else if (!PdfPageClipPath.TryCreatePath(primitive.PathCommands, primitive.FillRule, out shapeClip)) {
            return;
        }

        if (primitive.ClipPath.HasValue) {
            shapeClip = PdfPageClipPath.ResolveActiveClip(shapeClip, primitive.ClipPath.Value);
        }
        if (!TryFitClipToDrawing(shapeClip, drawing.Width, drawing.Height, out PdfPageClipPath fitted)) return;
        OfficeClipPath? clip = fitted.ToOfficeClipPath(fitted.X, fitted.Y);
        if (clip == null) return;

        OfficeDrawing tile = paint.Resource.Tile.Clone();
        if (paint.Tint.HasValue) TintPatternTile(tile, paint.Tint.Value);
        var patternDrawing = new OfficeDrawing(fitted.Width, fitted.Height);
        OfficeTransform localTransform = paint.Transform.Then(OfficeTransform.Translate(-fitted.X, -fitted.Y));
        patternDrawing.AddTilingPattern(
            tile,
            new OfficeImagePlacement(0D, 0D, fitted.Width, fitted.Height),
            paint.Resource.HorizontalStep,
            paint.Resource.VerticalStep,
            localTransform,
            maximumTileCount: 16384,
            opacity: paint.Opacity);
        drawing.AddClippedDrawing(patternDrawing, fitted.X, fitted.Y, clip);
    }

    private static void TintPatternTile(OfficeDrawing tile, OfficeColor tint) {
        for (int i = 0; i < tile.Shapes.Count; i++) {
            OfficeShape shape = tile.Shapes[i].Shape;
            if (shape.FillColor.HasValue && shape.FillColor.Value.A > 0) shape.FillColor = OfficeColor.FromRgba(tint.R, tint.G, tint.B, shape.FillColor.Value.A);
            if (shape.StrokeColor.HasValue && shape.StrokeColor.Value.A > 0) shape.StrokeColor = OfficeColor.FromRgba(tint.R, tint.G, tint.B, shape.StrokeColor.Value.A);
        }
    }

    private Dictionary<string, PdfPageTilingPatternResource> GetTilingPatternResources(PdfDictionary? resources) {
        var result = new Dictionary<string, PdfPageTilingPatternResource>(StringComparer.Ordinal);
        if (resources == null || !resources.Items.TryGetValue("Pattern", out PdfObject? patternObject)) return result;
        PdfDictionary? patterns = ResolveDictionary(patternObject);
        if (patterns == null) return result;
        foreach (KeyValuePair<string, PdfObject> entry in patterns.Items) {
            if (TryReadTilingPattern(entry.Value, resources, out PdfPageTilingPatternResource? pattern)) result[entry.Key] = pattern;
        }
        return result;
    }

    private bool TryReadTilingPattern(PdfObject? value, PdfDictionary parentResources, out PdfPageTilingPatternResource pattern) {
        pattern = null!;
        int? paintType;
        int? tilingType;
        if (ResolveObject(value) is not PdfStream stream ||
            TryReadInteger(stream.Dictionary.Items.TryGetValue("PatternType", out PdfObject? typeObject) ? typeObject : null) != 1 ||
            ((paintType = TryReadInteger(stream.Dictionary.Items.TryGetValue("PaintType", out PdfObject? paintTypeObject) ? paintTypeObject : null)) != 1 && paintType != 2) ||
            ((tilingType = TryReadInteger(stream.Dictionary.Items.TryGetValue("TilingType", out PdfObject? tilingTypeObject) ? tilingTypeObject : null)) < 1 || tilingType > 3) ||
            !TryReadRectangle(stream.Dictionary.Items.TryGetValue("BBox", out PdfObject? boxObject) ? boxObject : null, out (double X1, double Y1, double X2, double Y2) box) ||
            ResolveObject(stream.Dictionary.Items.TryGetValue("XStep", out PdfObject? xStepObject) ? xStepObject : null) is not PdfNumber xStep ||
            ResolveObject(stream.Dictionary.Items.TryGetValue("YStep", out PdfObject? yStepObject) ? yStepObject : null) is not PdfNumber yStep ||
            !IsFinite(xStep.Value) || !IsFinite(yStep.Value) ||
            Math.Abs(xStep.Value) <= 0.0000001D || Math.Abs(yStep.Value) <= 0.0000001D) return false;
        double width = box.X2 - box.X1;
        double height = box.Y2 - box.Y1;
        if (width <= 0D || height <= 0D) return false;
        PdfDictionary? resources = ResolveDictionary(stream.Dictionary.Items.TryGetValue("Resources", out PdfObject? resourceObject) ? resourceObject : null) ?? parentResources;
        OfficeDrawing tile = CreatePatternTileDrawing(stream, resources, box, width, height);
        Matrix2D matrix = stream.Dictionary.Items.TryGetValue("Matrix", out PdfObject? matrixObject)
            ? ReadPatternMatrix(matrixObject)
            : Matrix2D.Identity;
        if (!IsUsableTilingPatternMatrix(matrix)) return false;
        bool uncolored = paintType == 2;
        pattern = new PdfPageTilingPatternResource(tile, Math.Abs(xStep.Value), Math.Abs(yStep.Value), matrix, box.X1, box.Y2, uncolored);
        return true;
    }

    private static bool IsUsableTilingPatternMatrix(Matrix2D matrix) =>
        IsFinite(matrix.A) && IsFinite(matrix.B) && IsFinite(matrix.C) &&
        IsFinite(matrix.D) && IsFinite(matrix.E) && IsFinite(matrix.F) &&
        Math.Abs((matrix.A * matrix.D) - (matrix.B * matrix.C)) > 0.000000000001D;

    private OfficeDrawing CreatePatternTileDrawing(
        PdfStream stream,
        PdfDictionary? resources,
        (double X1, double Y1, double X2, double Y2) box,
        double width,
        double height) {
        var drawing = new OfficeDrawing(width, height);
        RegisterEmbeddedFonts(drawing, resources, new HashSet<PdfStream>(), 0);
        string content = PdfEncoding.Latin1GetString(DecodeIfNeeded(stream));
        if (content.Length == 0) return drawing;
        Matrix2D transform = Matrix2D.Translation(-box.X1, -box.Y1);
        var activeForms = new HashSet<PdfStream>();
        var elements = new List<PdfPageDrawingElement>();
        var primitives = new List<PdfPageVisualPrimitive>();
        CollectVisualPrimitivesAndForms(content, resources, transform, width, height, primitives, activeForms, includeTilingPatterns: false);
        for (int i = 0; i < primitives.Count; i++) elements.Add(PdfPageDrawingElement.FromPrimitive(primitives[i], elements.Count));

        var spans = new List<PdfTextSpan>();
        Dictionary<string, Func<byte[], string>> decoders = ResourceResolver.GetFontDecodersForForm(stream.Dictionary, _objects);
        Dictionary<string, Func<byte[], double>> widthProviders = ResourceResolver.GetFontWidthProviders(stream.Dictionary, _objects);
        Dictionary<string, PdfFontResource> fonts = ResourceResolver.GetFontsForResources(resources, _objects);
        string transformedContent = WrapContentWithTransform(content, transform, out int transformedOffset);
        CollectTextAndForms(transformedContent, resources, decoders, widthProviders, fonts, spans, activeForms, height, paintOrderOffset: -transformedOffset, useLogicalTextFilters: false);
        for (int i = 0; i < spans.Count; i++) elements.Add(PdfPageDrawingElement.FromText(spans[i], elements.Count));

        var placements = new List<PdfImagePlacement>();
        CollectImagePlacementsAndForms(content, resources, 0, transform, height, placements, activeForms);
        if (placements.Count > 0) {
            IReadOnlyList<PdfExtractedImage> images = GetImagesForResources(resources, 0, placements, colorizeImageMasks: true);
            for (int i = 0; i < placements.Count; i++) {
                PdfExtractedImage? image = FindImage(images, placements[i]);
                if (image != null) elements.Add(PdfPageDrawingElement.FromImage(placements[i], image, elements.Count));
            }
        }
        SortDrawingElements(elements);
        for (int i = 0; i < elements.Count; i++) AddDrawingElementCore(drawing, height, elements[i]);
        return drawing;
    }
}

using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

public sealed partial class PdfReadPage {
    /// <summary>
    /// Projects supported page drawing operators, text spans, and image placements into a dependency-free drawing scene.
    /// </summary>
    public OfficeDrawing ToDrawing() {
        (double Width, double Height) size = GetVisualPageSize();
        Matrix2D pageTransform = GetVisualPageTransform();
        var drawing = new OfficeDrawing(size.Width, size.Height);

        AddVisualPrimitives(drawing, size.Width, size.Height, pageTransform);
        AddTextSpans(drawing, size.Height, pageTransform);
        AddImages(drawing, size.Height, pageTransform);
        AddAnnotationAppearances(drawing, size.Height, pageTransform);

        return drawing;
    }

    private void AddVisualPrimitives(OfficeDrawing drawing, double pageWidth, double pageHeight, Matrix2D pageTransform) {
        IReadOnlyList<PdfPageVisualPrimitive> primitives = GetVisualPrimitives(pageWidth, pageHeight, pageTransform);
        for (int i = 0; i < primitives.Count; i++) {
            PdfPageVisualPrimitive primitive = primitives[i];
            if (primitive.Kind == PdfPageVisualPrimitiveKind.Rectangle) {
                AddRectangle(drawing, primitive);
            } else if (primitive.Kind == PdfPageVisualPrimitiveKind.Line) {
                AddLine(drawing, primitive);
            } else if (primitive.Kind == PdfPageVisualPrimitiveKind.Path) {
                AddPath(drawing, primitive);
            }
        }
    }

    private static void AddRectangle(OfficeDrawing drawing, PdfPageVisualPrimitive primitive) {
        if (!HasPositiveArea(primitive.X, primitive.Y, primitive.Width, primitive.Height, drawing.Width, drawing.Height)) {
            return;
        }

        OfficeShape shape = OfficeShape.Rectangle(primitive.Width, primitive.Height);
        shape.FillColor = primitive.FillColor;
        shape.FillGradient = primitive.FillGradient;
        shape.FillRadialGradient = primitive.FillRadialGradient;
        shape.StrokeColor = primitive.StrokeColor;
        shape.StrokeGradient = primitive.StrokeGradient;
        shape.StrokeRadialGradient = primitive.StrokeRadialGradient;
        shape.StrokeWidth = primitive.StrokeWidth;
        shape.StrokeDashStyle = primitive.StrokeDashStyle;
        shape.StrokeLineCap = primitive.StrokeLineCap;
        shape.StrokeLineJoin = primitive.StrokeLineJoin;
        shape.FillOpacity = primitive.FillOpacity;
        shape.StrokeOpacity = primitive.StrokeOpacity;
        shape.FillRule = primitive.FillRule;
        if (TryAddClippedShape(drawing, shape, primitive.X, primitive.Y, primitive.ClipPath)) {
            return;
        }

        drawing.AddShape(shape, primitive.X, primitive.Y);
    }

    private static void AddLine(OfficeDrawing drawing, PdfPageVisualPrimitive primitive) {
        double left = Math.Min(primitive.X1, primitive.X2);
        double top = Math.Min(primitive.Y1, primitive.Y2);
        double right = Math.Max(primitive.X1, primitive.X2);
        double bottom = Math.Max(primitive.Y1, primitive.Y2);
        if ((NearlyEqual(left, right) && NearlyEqual(top, bottom)) ||
            left < 0D ||
            top < 0D ||
            right > drawing.Width ||
            bottom > drawing.Height) {
            return;
        }

        OfficeShape shape = OfficeShape.Line(primitive.X1 - left, primitive.Y1 - top, primitive.X2 - left, primitive.Y2 - top);
        shape.StrokeColor = primitive.StrokeColor;
        shape.StrokeGradient = primitive.StrokeGradient;
        shape.StrokeRadialGradient = primitive.StrokeRadialGradient;
        shape.StrokeWidth = primitive.StrokeWidth;
        shape.StrokeDashStyle = primitive.StrokeDashStyle;
        shape.StrokeLineCap = primitive.StrokeLineCap;
        shape.StrokeLineJoin = primitive.StrokeLineJoin;
        shape.StrokeOpacity = primitive.StrokeOpacity;
        if (TryAddClippedShape(drawing, shape, left, top, primitive.ClipPath)) {
            return;
        }

        drawing.AddShape(shape, left, top);
    }

    private static void AddPath(OfficeDrawing drawing, PdfPageVisualPrimitive primitive) {
        if (!HasPositiveArea(primitive.X, primitive.Y, primitive.Width, primitive.Height, drawing.Width, drawing.Height)) {
            return;
        }

        OfficeShape shape = OfficeShape.Path(primitive.PathCommands);
        shape.FillColor = primitive.FillColor;
        shape.FillGradient = primitive.FillGradient;
        shape.FillRadialGradient = primitive.FillRadialGradient;
        shape.StrokeColor = primitive.StrokeColor;
        shape.StrokeGradient = primitive.StrokeGradient;
        shape.StrokeRadialGradient = primitive.StrokeRadialGradient;
        shape.StrokeWidth = primitive.StrokeWidth;
        shape.StrokeDashStyle = primitive.StrokeDashStyle;
        shape.StrokeLineCap = primitive.StrokeLineCap;
        shape.StrokeLineJoin = primitive.StrokeLineJoin;
        shape.FillOpacity = primitive.FillOpacity;
        shape.StrokeOpacity = primitive.StrokeOpacity;
        shape.FillRule = primitive.FillRule;
        if (TryAddClippedShape(drawing, shape, primitive.X, primitive.Y, primitive.ClipPath)) {
            return;
        }

        drawing.AddShape(shape, primitive.X, primitive.Y);
    }

    private static bool TryAddClippedShape(OfficeDrawing drawing, OfficeShape shape, double x, double y, PdfPageClipPath? clipPath) {
        if (!clipPath.HasValue) {
            return false;
        }

        PdfPageClipPath clip = clipPath.Value;
        if (clip.X < 0D ||
            clip.Y < 0D ||
            clip.Width <= 0D ||
            clip.Height <= 0D ||
            clip.X + clip.Width > drawing.Width ||
            clip.Y + clip.Height > drawing.Height) {
            return false;
        }

        OfficeClipPath? localClip = clip.ToOfficeClipPath(x, y);
        if (localClip != null) {
            shape.ClipPath = localClip;
            return false;
        }

        OfficeClipPath? groupClip = clip.ToOfficeClipPath(clip.X, clip.Y);
        if (groupClip == null) {
            return false;
        }

        double localX = x - clip.X;
        double localY = y - clip.Y;
        if (localX < 0D || localY < 0D) {
            return false;
        }

        double innerWidth = Math.Max(clip.Width, localX + shape.Width);
        double innerHeight = Math.Max(clip.Height, localY + shape.Height);
        var innerDrawing = new OfficeDrawing(innerWidth, innerHeight);
        innerDrawing.AddShape(shape, localX, localY);
        drawing.AddClippedDrawing(innerDrawing, clip.X, clip.Y, groupClip);
        return true;
    }

    private IReadOnlyList<PdfPageVisualPrimitive> GetVisualPrimitives(double pageWidth, double pageHeight, Matrix2D pageTransform) {
        var primitives = new List<PdfPageVisualPrimitive>();
        PdfDictionary? pageResources = ResolveDictionary(GetInheritedValue("Resources"));
        var activeForms = new HashSet<PdfStream>();
        string content = GetContentStreamContent();
        if (content.Length > 0) {
            CollectVisualPrimitivesAndForms(content, pageResources, pageTransform, pageWidth, pageHeight, primitives, activeForms);
        }

        return primitives.Count == 0 ? Array.Empty<PdfPageVisualPrimitive>() : primitives.AsReadOnly();
    }

    private void CollectVisualPrimitivesAndForms(
        string content,
        PdfDictionary? resources,
        Matrix2D baseTransform,
        double pageWidth,
        double pageHeight,
        List<PdfPageVisualPrimitive> primitives,
        HashSet<PdfStream> activeForms) {
        string transformedContent = WrapContentWithTransform(content, baseTransform);
        primitives.AddRange(PdfPageContentVisualParser.Parse(
            transformedContent,
            pageWidth,
            pageHeight,
            GetGraphicsStateResources(resources),
            GetColorSpaceResources(resources),
            GetShadingResources(resources),
            GetShadingPatternResources(resources),
            GetOptionalContentVisibility(resources)));

        foreach (PdfPageXObjectInvocation invocation in PdfPageXObjectInvocationParser.Parse(content, baseTransform, pageHeight, GetGraphicsStateResources(resources), GetColorSpaceResources(resources), GetOptionalContentVisibility(resources))) {
            if (!TryGetFormStream(resources, invocation.Name, out PdfStream formStream)) {
                continue;
            }

            if (!activeForms.Add(formStream)) {
                continue;
            }

            try {
                PdfDictionary formDictionary = formStream.Dictionary;
                PdfDictionary? formResources = ResolveDictionary(formDictionary.Items.TryGetValue("Resources", out PdfObject? resourcesObject) ? resourcesObject : null) ?? resources;
                Matrix2D formTransform = ApplyFormMatrix(invocation.Transform, formDictionary);
                string formContent = WrapFormContentWithBoundingBoxClip(PdfEncoding.Latin1GetString(DecodeIfNeeded(formStream)), formDictionary);
                CollectVisualPrimitivesAndForms(formContent, formResources, formTransform, pageWidth, pageHeight, primitives, activeForms);
            } finally {
                activeForms.Remove(formStream);
            }
        }
    }

    private Dictionary<string, PdfPageShadingResource> GetShadingResources(PdfDictionary? resources) {
        var result = new Dictionary<string, PdfPageShadingResource>(StringComparer.Ordinal);
        if (resources == null ||
            !resources.Items.TryGetValue("Shading", out PdfObject? shadingObject)) {
            return result;
        }

        PdfDictionary? shadings = ResolveDictionary(shadingObject);
        if (shadings == null) {
            return result;
        }

        foreach (KeyValuePair<string, PdfObject> entry in shadings.Items) {
            if (TryReadShading(entry.Value, out PdfPageShadingResource shading)) {
                result[entry.Key] = shading;
            }
        }

        return result;
    }

    private Dictionary<string, PdfPageShadingPatternResource> GetShadingPatternResources(PdfDictionary? resources) {
        var result = new Dictionary<string, PdfPageShadingPatternResource>(StringComparer.Ordinal);
        if (resources == null ||
            !resources.Items.TryGetValue("Pattern", out PdfObject? patternObject)) {
            return result;
        }

        PdfDictionary? patterns = ResolveDictionary(patternObject);
        if (patterns == null) {
            return result;
        }

        foreach (KeyValuePair<string, PdfObject> entry in patterns.Items) {
            if (TryReadShadingPattern(entry.Value, out PdfPageShadingPatternResource pattern)) {
                result[entry.Key] = pattern;
            }
        }

        return result;
    }

    private bool TryReadShadingPattern(PdfObject? value, out PdfPageShadingPatternResource pattern) {
        pattern = default;
        PdfDictionary? dictionary = ResolveDictionary(value);
        if (dictionary == null ||
            TryReadInteger(dictionary.Items.TryGetValue("PatternType", out PdfObject? patternTypeObject) ? patternTypeObject : null) != 2 ||
            !dictionary.Items.TryGetValue("Shading", out PdfObject? shadingObject) ||
            !TryReadShading(shadingObject, out PdfPageShadingResource shading)) {
            return false;
        }

        Matrix2D matrix = dictionary.Items.TryGetValue("Matrix", out PdfObject? matrixObject)
            ? ReadPatternMatrix(matrixObject)
            : Matrix2D.Identity;
        pattern = new PdfPageShadingPatternResource(shading, matrix);
        return true;
    }

    private Matrix2D ReadPatternMatrix(PdfObject? matrixObject) {
        PdfArray? matrix = ResolveArray(matrixObject);
        if (matrix == null || matrix.Items.Count < 6) {
            return Matrix2D.Identity;
        }

        return new Matrix2D(
            ReadMatrixNumber(matrix, 0, 1D),
            ReadMatrixNumber(matrix, 1, 0D),
            ReadMatrixNumber(matrix, 2, 0D),
            ReadMatrixNumber(matrix, 3, 1D),
            ReadMatrixNumber(matrix, 4, 0D),
            ReadMatrixNumber(matrix, 5, 0D));
    }

    private bool TryReadShading(PdfObject? value, out PdfPageShadingResource shading) {
        shading = default;
        PdfDictionary? dictionary = ResolveDictionary(value);
        if (dictionary == null ||
            !dictionary.Items.TryGetValue("Coords", out PdfObject? coordsObject)) {
            return false;
        }

        int? shadingType = TryReadInteger(dictionary.Items.TryGetValue("ShadingType", out PdfObject? shadingTypeObject) ? shadingTypeObject : null);
        IReadOnlyList<double> coords = ReadNumberArray(coordsObject);
        if ((shadingType == 2 && coords.Count < 4) ||
            (shadingType == 3 && coords.Count < 6)) {
            return false;
        }

        PdfPageColorSpaceKind colorSpace = PdfPageColorSpaceKind.DeviceGray;
        if (dictionary.Items.TryGetValue("ColorSpace", out PdfObject? colorSpaceObject)) {
            TryReadColorSpaceResource(colorSpaceObject, out colorSpace);
        }

        PdfDictionary? function = ResolveFunctionDictionary(dictionary.Items.TryGetValue("Function", out PdfObject? functionObject) ? functionObject : null);
        if (function == null ||
            TryReadInteger(function.Items.TryGetValue("FunctionType", out PdfObject? functionTypeObject) ? functionTypeObject : null) != 2) {
            return false;
        }

        OfficeColor startColor = ReadFunctionColor(function, "C0", colorSpace, false);
        OfficeColor endColor = ReadFunctionColor(function, "C1", colorSpace, true);
        if (shadingType == 2) {
            shading = new PdfPageShadingResource(coords[0], coords[1], coords[2], coords[3], startColor, endColor);
            return true;
        }

        if (shadingType == 3) {
            shading = new PdfPageShadingResource(coords[0], coords[1], Math.Max(0D, coords[2]), coords[3], coords[4], Math.Max(0D, coords[5]), startColor, endColor);
            return true;
        }

        return false;
    }

    private PdfDictionary? ResolveFunctionDictionary(PdfObject? functionObject) {
        PdfObject? resolved = ResolveObject(functionObject);
        if (resolved is PdfArray array && array.Items.Count > 0) {
            resolved = ResolveObject(array.Items[0]);
        }

        return resolved is PdfDictionary dictionary ? dictionary : null;
    }

    private OfficeColor ReadFunctionColor(PdfDictionary function, string key, PdfPageColorSpaceKind colorSpace, bool endColor) {
        IReadOnlyList<double> components = function.Items.TryGetValue(key, out PdfObject? value)
            ? ReadNumberArray(value)
            : Array.Empty<double>();
        return ReadColorComponents(components, colorSpace, endColor);
    }

    private static OfficeColor ReadColorComponents(IReadOnlyList<double> components, PdfPageColorSpaceKind colorSpace, bool endColor) {
        switch (colorSpace) {
            case PdfPageColorSpaceKind.DeviceRgb:
                return OfficeColor.FromRgb(
                    ToColorByte(ComponentAt(components, 0, endColor ? 1D : 0D)),
                    ToColorByte(ComponentAt(components, 1, endColor ? 1D : 0D)),
                    ToColorByte(ComponentAt(components, 2, endColor ? 1D : 0D)));
            case PdfPageColorSpaceKind.DeviceCmyk:
                double cyan = ComponentAt(components, 0, 0D);
                double magenta = ComponentAt(components, 1, 0D);
                double yellow = ComponentAt(components, 2, 0D);
                double black = ComponentAt(components, 3, endColor ? 0D : 1D);
                return OfficeColor.FromRgb(
                    ToColorByte((1D - cyan) * (1D - black)),
                    ToColorByte((1D - magenta) * (1D - black)),
                    ToColorByte((1D - yellow) * (1D - black)));
            default:
                byte gray = ToColorByte(ComponentAt(components, 0, endColor ? 1D : 0D));
                return OfficeColor.FromRgb(gray, gray, gray);
        }
    }

    private static double ComponentAt(IReadOnlyList<double> components, int index, double fallback) =>
        index < components.Count ? Clamp01(components[index]) : fallback;

    private static byte ToColorByte(double value) =>
        (byte)Math.Round(Clamp01(value) * 255D);

    private static double Clamp01(double value) =>
        value < 0D ? 0D : value > 1D ? 1D : value;

    private Dictionary<string, PdfPageGraphicsStateResource> GetGraphicsStateResources(PdfDictionary? resources) {
        var result = new Dictionary<string, PdfPageGraphicsStateResource>(StringComparer.Ordinal);
        if (resources == null ||
            !resources.Items.TryGetValue("ExtGState", out PdfObject? extGStateObject)) {
            return result;
        }

        PdfDictionary? extGStates = ResolveDictionary(extGStateObject);
        if (extGStates == null) {
            return result;
        }

        foreach (KeyValuePair<string, PdfObject> entry in extGStates.Items) {
            PdfDictionary? state = ResolveDictionary(entry.Value);
            if (state == null) {
                continue;
            }

            double? fillOpacity = ReadOpacity(state, "ca");
            double? strokeOpacity = ReadOpacity(state, "CA");
            double? strokeWidth = ReadStrokeWidth(state);
            OfficeStrokeDashStyle? strokeDashStyle = ReadStrokeDashStyle(state);
            OfficeStrokeLineCap? strokeLineCap = ReadStrokeLineCap(state);
            OfficeStrokeLineJoin? strokeLineJoin = ReadStrokeLineJoin(state);
            if (fillOpacity.HasValue ||
                strokeOpacity.HasValue ||
                strokeWidth.HasValue ||
                strokeDashStyle.HasValue ||
                strokeLineCap.HasValue ||
                strokeLineJoin.HasValue) {
                result[entry.Key] = new PdfPageGraphicsStateResource(fillOpacity, strokeOpacity, strokeWidth, strokeDashStyle, strokeLineCap, strokeLineJoin);
            }
        }

        return result;
    }

    private Dictionary<string, PdfPageColorSpaceKind> GetColorSpaceResources(PdfDictionary? resources) {
        var result = new Dictionary<string, PdfPageColorSpaceKind>(StringComparer.Ordinal);
        if (resources == null ||
            !resources.Items.TryGetValue("ColorSpace", out PdfObject? colorSpacesObject)) {
            return result;
        }

        PdfDictionary? colorSpaces = ResolveDictionary(colorSpacesObject);
        if (colorSpaces == null) {
            return result;
        }

        foreach (KeyValuePair<string, PdfObject> entry in colorSpaces.Items) {
            if (TryReadColorSpaceResource(entry.Value, out PdfPageColorSpaceKind colorSpace)) {
                result[entry.Key] = colorSpace;
            }
        }

        return result;
    }

    private bool TryReadColorSpaceResource(PdfObject? value, out PdfPageColorSpaceKind colorSpace) {
        PdfObject? resolved = ResolveObject(value);
        if (resolved is PdfName directName) {
            return TryReadStandardColorSpaceName(directName.Name, out colorSpace);
        }

        if (resolved is PdfArray array && array.Items.Count > 0 &&
            ResolveObject(array.Items[0]) is PdfName arrayName) {
            return TryReadStandardColorSpaceName(arrayName.Name, out colorSpace);
        }

        colorSpace = PdfPageColorSpaceKind.DeviceGray;
        return false;
    }

    private static bool TryReadStandardColorSpaceName(string name, out PdfPageColorSpaceKind colorSpace) {
        switch (name) {
            case "DeviceRGB":
            case "RGB":
                colorSpace = PdfPageColorSpaceKind.DeviceRgb;
                return true;
            case "DeviceCMYK":
            case "CMYK":
                colorSpace = PdfPageColorSpaceKind.DeviceCmyk;
                return true;
            case "DeviceGray":
            case "G":
                colorSpace = PdfPageColorSpaceKind.DeviceGray;
                return true;
            case "Pattern":
                colorSpace = PdfPageColorSpaceKind.Pattern;
                return true;
            default:
                colorSpace = PdfPageColorSpaceKind.DeviceGray;
                return false;
        }
    }

    private double? ReadOpacity(PdfDictionary dictionary, string key) {
        if (!dictionary.Items.TryGetValue(key, out PdfObject? value) ||
            ResolveObject(value) is not PdfNumber number) {
            return null;
        }

        if (number.Value < 0D) {
            return 0D;
        }

        return number.Value > 1D ? 1D : number.Value;
    }

    private double? ReadStrokeWidth(PdfDictionary dictionary) {
        if (!dictionary.Items.TryGetValue("LW", out PdfObject? value) ||
            ResolveObject(value) is not PdfNumber number) {
            return null;
        }

        return Math.Max(0D, number.Value);
    }

    private OfficeStrokeDashStyle? ReadStrokeDashStyle(PdfDictionary dictionary) {
        if (!dictionary.Items.TryGetValue("D", out PdfObject? value) ||
            ResolveObject(value) is not PdfArray dash ||
            dash.Items.Count == 0 ||
            ResolveObject(dash.Items[0]) is not PdfArray dashArray) {
            return null;
        }

        IReadOnlyList<double> values = ReadNumberArray(dashArray);
        if (values.Count == 0) {
            return OfficeStrokeDashStyle.Solid;
        }

        if (values.Count >= 6) {
            return OfficeStrokeDashStyle.DashDotDot;
        }

        if (values.Count >= 4) {
            return OfficeStrokeDashStyle.DashDot;
        }

        if (values.Count >= 2) {
            return values[0] <= values[1] ? OfficeStrokeDashStyle.Dot : OfficeStrokeDashStyle.Dash;
        }

        return OfficeStrokeDashStyle.Solid;
    }

    private OfficeStrokeLineCap? ReadStrokeLineCap(PdfDictionary dictionary) {
        int? lineCap = TryReadInteger(dictionary.Items.TryGetValue("LC", out PdfObject? value) ? value : null);
        switch (lineCap) {
            case 0:
                return OfficeStrokeLineCap.Butt;
            case 1:
                return OfficeStrokeLineCap.Round;
            case 2:
                return OfficeStrokeLineCap.Square;
            default:
                return null;
        }
    }

    private OfficeStrokeLineJoin? ReadStrokeLineJoin(PdfDictionary dictionary) {
        int? lineJoin = TryReadInteger(dictionary.Items.TryGetValue("LJ", out PdfObject? value) ? value : null);
        switch (lineJoin) {
            case 0:
                return OfficeStrokeLineJoin.Miter;
            case 1:
                return OfficeStrokeLineJoin.Round;
            case 2:
                return OfficeStrokeLineJoin.Bevel;
            default:
                return null;
        }
    }

    private void AddAnnotationAppearances(OfficeDrawing drawing, double pageHeight, Matrix2D pageTransform) {
        if (!_pageDict.Items.TryGetValue("Annots", out PdfObject? annotationsObject)) {
            return;
        }

        PdfArray? annotations = ResolveArray(annotationsObject);
        if (annotations == null) {
            return;
        }

        PdfDictionary? pageResources = ResolveDictionary(GetInheritedValue("Resources"));
        Dictionary<string, Func<byte[], string>> pageDecoders = ResourceResolver.GetFontDecoders(_pageDict, _objects);
        Dictionary<string, Func<byte[], double>> pageWidthProviders = ResourceResolver.GetFontWidthProviders(_pageDict, _objects);
        Dictionary<string, PdfFontResource> pageFonts = ResourceResolver.GetFontsForResources(pageResources, _objects);
        var activeForms = new HashSet<PdfStream>();
        for (int i = 0; i < annotations.Items.Count; i++) {
            PdfDictionary? annotation = ResolveDictionary(annotations.Items[i]);
            if (annotation == null ||
                !TryReadRectangle(annotation.Items.TryGetValue("Rect", out PdfObject? rectangleObject) ? rectangleObject : null, out (double X1, double Y1, double X2, double Y2) rectangle) ||
                IsHiddenAnnotation(annotation) ||
                !TryGetNormalAppearanceStream(annotation, out PdfStream appearanceStream)) {
                continue;
            }

            PdfDictionary? appearanceResources = ResolveDictionary(appearanceStream.Dictionary.Items.TryGetValue("Resources", out PdfObject? resourcesObject) ? resourcesObject : null) ?? pageResources;
            string appearanceContent = PdfEncoding.Latin1GetString(DecodeIfNeeded(appearanceStream));
            if (appearanceContent.Length == 0) {
                continue;
            }

            Matrix2D appearanceTransform = Matrix2D.Multiply(pageTransform, CreateAnnotationAppearanceTransform(rectangle, appearanceStream.Dictionary));
            var primitives = new List<PdfPageVisualPrimitive>();
            CollectVisualPrimitivesAndForms(appearanceContent, appearanceResources, appearanceTransform, drawing.Width, pageHeight, primitives, activeForms);
            for (int primitiveIndex = 0; primitiveIndex < primitives.Count; primitiveIndex++) {
                PdfPageVisualPrimitive primitive = primitives[primitiveIndex];
                if (primitive.Kind == PdfPageVisualPrimitiveKind.Rectangle) {
                    AddRectangle(drawing, primitive);
                } else if (primitive.Kind == PdfPageVisualPrimitiveKind.Line) {
                    AddLine(drawing, primitive);
                } else if (primitive.Kind == PdfPageVisualPrimitiveKind.Path) {
                    AddPath(drawing, primitive);
                }
            }

            var textSpans = new List<PdfTextSpan>();
            Dictionary<string, Func<byte[], string>> appearanceDecoders = MergeDecoders(pageDecoders, ResourceResolver.GetFontDecodersForForm(appearanceStream.Dictionary, _objects));
            Dictionary<string, Func<byte[], double>> appearanceWidthProviders = MergeWidthProviders(pageWidthProviders, ResourceResolver.GetFontWidthProviders(appearanceStream.Dictionary, _objects));
            Dictionary<string, PdfFontResource> appearanceFonts = MergeFonts(pageFonts, ResourceResolver.GetFontsForResources(appearanceResources, _objects));
            string transformedAppearanceContent = WrapContentWithTransform(appearanceContent, appearanceTransform);
            CollectTextAndForms(transformedAppearanceContent, appearanceResources, appearanceDecoders, appearanceWidthProviders, appearanceFonts, textSpans, activeForms, pageHeight);
            for (int textIndex = 0; textIndex < textSpans.Count; textIndex++) {
                AddTextSpan(drawing, pageHeight, textSpans[textIndex]);
            }

            var imagePlacements = new List<PdfImagePlacement>();
            CollectImagePlacementsAndForms(appearanceContent, appearanceResources, 0, appearanceTransform, pageHeight, imagePlacements, activeForms);
            if (imagePlacements.Count > 0) {
                AddImagePlacements(drawing, pageHeight, imagePlacements, GetImagesForResources(appearanceResources, 0, imagePlacements));
            }
        }
    }

    private bool TryGetNormalAppearanceStream(PdfDictionary annotation, out PdfStream stream) {
        stream = null!;
        PdfDictionary? appearance = ResolveDictionary(annotation.Items.TryGetValue("AP", out PdfObject? appearanceObject) ? appearanceObject : null);
        if (appearance == null || !appearance.Items.TryGetValue("N", out PdfObject? normalAppearanceObject)) {
            return false;
        }

        PdfObject? normalAppearance = ResolveObject(normalAppearanceObject);
        if (normalAppearance is PdfStream directStream) {
            stream = directStream;
            return true;
        }

        if (normalAppearance is not PdfDictionary stateDictionary || stateDictionary.Items.Count == 0) {
            return false;
        }

        if (annotation.Items.TryGetValue("AS", out PdfObject? appearanceStateObject) &&
            ResolveObject(appearanceStateObject) is PdfName appearanceState &&
            stateDictionary.Items.TryGetValue(appearanceState.Name, out PdfObject? stateObject) &&
            ResolveObject(stateObject) is PdfStream stateStream) {
            stream = stateStream;
            return true;
        }

        foreach (KeyValuePair<string, PdfObject> state in stateDictionary.Items) {
            if (string.Equals(state.Key, "Off", StringComparison.Ordinal)) {
                continue;
            }

            if (ResolveObject(state.Value) is PdfStream fallbackStream) {
                stream = fallbackStream;
                return true;
            }
        }

        foreach (KeyValuePair<string, PdfObject> state in stateDictionary.Items) {
            if (ResolveObject(state.Value) is PdfStream fallbackStream) {
                stream = fallbackStream;
                return true;
            }
        }

        return false;
    }

    private Matrix2D CreateAnnotationAppearanceTransform((double X1, double Y1, double X2, double Y2) rectangle, PdfDictionary appearanceDictionary) {
        double bboxX1 = 0D;
        double bboxY1 = 0D;
        double bboxWidth = rectangle.X2 - rectangle.X1;
        double bboxHeight = rectangle.Y2 - rectangle.Y1;
        if (TryReadBox(appearanceDictionary.Items.TryGetValue("BBox", out PdfObject? bboxObject) ? bboxObject : null, out (double X1, double Y1, double X2, double Y2) bbox)) {
            bboxX1 = bbox.X1;
            bboxY1 = bbox.Y1;
            bboxWidth = bbox.X2 - bbox.X1;
            bboxHeight = bbox.Y2 - bbox.Y1;
        }

        double scaleX = bboxWidth > 0D ? (rectangle.X2 - rectangle.X1) / bboxWidth : 1D;
        double scaleY = bboxHeight > 0D ? (rectangle.Y2 - rectangle.Y1) / bboxHeight : 1D;
        var rectangleTransform = new Matrix2D(
            scaleX,
            0D,
            0D,
            scaleY,
            rectangle.X1 - (bboxX1 * scaleX),
            rectangle.Y1 - (bboxY1 * scaleY));
        return Matrix2D.Multiply(rectangleTransform, ReadAppearanceMatrix(appearanceDictionary));
    }

    private Matrix2D ReadAppearanceMatrix(PdfDictionary appearanceDictionary) {
        if (!appearanceDictionary.Items.TryGetValue("Matrix", out PdfObject? matrixObject) ||
            ResolveObject(matrixObject) is not PdfArray matrix ||
            matrix.Items.Count < 6) {
            return Matrix2D.Identity;
        }

        return new Matrix2D(
            ReadMatrixNumber(matrix, 0, 1D),
            ReadMatrixNumber(matrix, 1, 0D),
            ReadMatrixNumber(matrix, 2, 0D),
            ReadMatrixNumber(matrix, 3, 1D),
            ReadMatrixNumber(matrix, 4, 0D),
            ReadMatrixNumber(matrix, 5, 0D));
    }

    private double ReadMatrixNumber(PdfArray matrix, int index, double fallback) =>
        ResolveObject(matrix.Items[index]) is PdfNumber number ? number.Value : fallback;

    private bool TryReadBox(PdfObject? obj, out (double X1, double Y1, double X2, double Y2) box) =>
        TryReadRectangle(obj, out box);

    private bool IsHiddenAnnotation(PdfDictionary annotation) {
        int? flags = TryReadInteger(annotation.Items.TryGetValue("F", out PdfObject? flagsObject) ? flagsObject : null);
        if (!flags.HasValue) {
            return false;
        }

        const int invisible = 1;
        const int hidden = 2;
        const int noView = 32;
        return (flags.Value & (invisible | hidden | noView)) != 0;
    }

    private void AddTextSpans(OfficeDrawing drawing, double pageHeight, Matrix2D pageTransform) {
        IReadOnlyList<PdfTextSpan> spans = GetVisualTextSpans(pageHeight, pageTransform);
        for (int i = 0; i < spans.Count; i++) {
            AddTextSpan(drawing, pageHeight, spans[i]);
        }
    }

    private IReadOnlyList<PdfTextSpan> GetVisualTextSpans(double pageHeight, Matrix2D pageTransform) {
        var spans = new List<PdfTextSpan>();
        PdfDictionary? pageResources = ResolveDictionary(GetInheritedValue("Resources"));
        Dictionary<string, Func<byte[], string>> pageDecoders = ResourceResolver.GetFontDecoders(_pageDict, _objects);
        Dictionary<string, Func<byte[], double>> pageWidthProviders = ResourceResolver.GetFontWidthProviders(_pageDict, _objects);
        Dictionary<string, PdfFontResource> pageFonts = ResourceResolver.GetFontsForResources(pageResources, _objects);
        var activeForms = new HashSet<PdfStream>();

        string content = GetContentStreamContent();
        if (content.Length > 0) {
            CollectTextAndForms(
                WrapContentWithTransform(content, pageTransform),
                pageResources,
                pageDecoders,
                pageWidthProviders,
                pageFonts,
                spans,
                activeForms,
                pageHeight);
        }

        return spans.Count == 0 ? Array.Empty<PdfTextSpan>() : spans.AsReadOnly();
    }

    private static void AddTextSpan(OfficeDrawing drawing, double pageHeight, PdfTextSpan span) {
        if (string.IsNullOrEmpty(span.Text) || !span.IsVisible) {
            return;
        }

        double height = Math.Max(1D, span.FontSize * 1.25D);
        double width = Math.Max(span.Advance, span.Text.Length * span.FontSize * 0.55D);
        double x = Clamp(span.X, 0D, drawing.Width);
        double y = Clamp(pageHeight - span.Y - span.FontSize, 0D, drawing.Height);
        double baselineY = Clamp(pageHeight - span.Y, 0D, drawing.Height);
        width = Math.Min(width, Math.Max(1D, drawing.Width - x));
        height = Math.Min(height, Math.Max(1D, drawing.Height - y));
        if (TryAddClippedTextSpan(drawing, span, x, y, width, height, baselineY)) {
            return;
        }

        drawing.AddText(
            span.Text,
            x,
            y,
            width,
            height,
            ToOfficeFontInfo(span.BaseFont, span.FontSize),
            span.Color ?? OfficeColor.Black,
            rotationDegrees: -span.RotationDegrees,
            rotationCenterX: x,
            rotationCenterY: baselineY,
            wrapText: false);
    }

    private static bool TryAddClippedTextSpan(OfficeDrawing drawing, PdfTextSpan span, double x, double y, double width, double height, double baselineY) {
        if (!span.ClipPath.HasValue) {
            return false;
        }

        PdfPageClipPath clip = span.ClipPath.Value;
        OfficeClipPath? officeClipPath = clip.ToOfficeClipPath(clip.X, clip.Y);
        if (officeClipPath == null) {
            return false;
        }

        double clipRight = clip.X + clip.Width;
        double clipBottom = clip.Y + clip.Height;
        if (clip.IsRectangle && x >= clip.X && y >= clip.Y && x + width <= clipRight && y + height <= clipBottom) {
            return false;
        }

        if (x + width <= clip.X || y + height <= clip.Y || x >= clipRight || y >= clipBottom) {
            return true;
        }

        double localX = x - clip.X;
        double localY = y - clip.Y;
        if (clip.X < 0D ||
            clip.Y < 0D ||
            clip.Width <= 0D ||
            clip.Height <= 0D ||
            clipRight > drawing.Width ||
            clipBottom > drawing.Height) {
            return false;
        }

        double innerWidth = Math.Max(clip.Width, localX + width);
        double innerHeight = Math.Max(clip.Height, localY + height);
        var innerDrawing = new OfficeDrawing(innerWidth, innerHeight);
        innerDrawing.AddText(
            span.Text,
            localX,
            localY,
            width,
            height,
            ToOfficeFontInfo(span.BaseFont, span.FontSize),
            span.Color ?? OfficeColor.Black,
            rotationDegrees: -span.RotationDegrees,
            rotationCenterX: x - clip.X,
            rotationCenterY: baselineY - clip.Y,
            wrapText: false);
        drawing.AddClippedDrawing(innerDrawing, clip.X, clip.Y, officeClipPath);
        return true;
    }

    private static OfficeFontInfo ToOfficeFontInfo(string? baseFont, double size) {
        string normalized = StripSubsetPrefix(baseFont);
        OfficeFontStyle style = OfficeFontStyle.Regular;
        if (ContainsFontStyleToken(normalized, "Bold") ||
            ContainsFontStyleToken(normalized, "Black") ||
            ContainsFontStyleToken(normalized, "Heavy") ||
            ContainsFontStyleToken(normalized, "Demi") ||
            ContainsFontStyleToken(normalized, "SemiBold")) {
            style |= OfficeFontStyle.Bold;
        }

        if (ContainsFontStyleToken(normalized, "Italic") ||
            ContainsFontStyleToken(normalized, "Oblique")) {
            style |= OfficeFontStyle.Italic;
        }

        string family = ResolveOfficeFontFamily(normalized);
        return new OfficeFontInfo(family, size, style);
    }

    private static string ResolveOfficeFontFamily(string baseFont) {
        if (string.IsNullOrWhiteSpace(baseFont)) {
            return "Helvetica";
        }

        string normalized = baseFont.Replace('_', ' ');
        if (normalized.StartsWith("Times-", StringComparison.Ordinal) ||
            normalized.StartsWith("TimesNewRoman", StringComparison.OrdinalIgnoreCase) ||
            normalized.StartsWith("Times New Roman", StringComparison.OrdinalIgnoreCase)) {
            return "Times New Roman";
        }

        if (normalized.StartsWith("Courier", StringComparison.OrdinalIgnoreCase)) {
            return "Courier New";
        }

        if (normalized.StartsWith("Helvetica", StringComparison.OrdinalIgnoreCase)) {
            return "Helvetica";
        }

        int hyphen = normalized.IndexOf('-');
        if (hyphen > 0) {
            normalized = normalized.Substring(0, hyphen);
        }

        normalized = RemoveFontSuffix(normalized, "BoldItalic");
        normalized = RemoveFontSuffix(normalized, "BoldOblique");
        normalized = RemoveFontSuffix(normalized, "SemiBold");
        normalized = RemoveFontSuffix(normalized, "DemiBold");
        normalized = RemoveFontSuffix(normalized, "Bold");
        normalized = RemoveFontSuffix(normalized, "Italic");
        normalized = RemoveFontSuffix(normalized, "Oblique");
        normalized = RemoveFontSuffix(normalized, "Regular");
        normalized = RemoveFontSuffix(normalized, "PSMT");
        normalized = RemoveFontSuffix(normalized, "MT");
        return string.IsNullOrWhiteSpace(normalized) ? "Helvetica" : normalized.Trim();
    }

    private static string StripSubsetPrefix(string? baseFont) {
        if (string.IsNullOrWhiteSpace(baseFont)) {
            return string.Empty;
        }

        string value = baseFont!.Trim();
        if (value.Length > 7 && value[6] == '+') {
            for (int i = 0; i < 6; i++) {
                char ch = value[i];
                if (ch < 'A' || ch > 'Z') {
                    return value;
                }
            }

            return value.Substring(7);
        }

        return value;
    }

    private static bool ContainsFontStyleToken(string fontName, string token) =>
        System.Globalization.CultureInfo.InvariantCulture.CompareInfo.IndexOf(fontName, token, System.Globalization.CompareOptions.IgnoreCase) >= 0;

    private static string RemoveFontSuffix(string value, string suffix) =>
        value.EndsWith(suffix, StringComparison.OrdinalIgnoreCase)
            ? value.Substring(0, value.Length - suffix.Length)
            : value;

    private void AddImages(OfficeDrawing drawing, double pageHeight, Matrix2D pageTransform) {
        IReadOnlyList<PdfImagePlacement> placements = GetVisualImagePlacements(pageHeight, pageTransform);
        if (placements.Count == 0) {
            return;
        }

        IReadOnlyList<PdfExtractedImage> images = GetImages(0, placements);
        AddImagePlacements(drawing, pageHeight, placements, images);
    }

    private IReadOnlyList<PdfImagePlacement> GetVisualImagePlacements(double pageHeight, Matrix2D pageTransform) {
        var placements = new List<PdfImagePlacement>();
        PdfDictionary? pageResources = ResolveDictionary(GetInheritedValue("Resources"));
        var activeForms = new HashSet<PdfStream>();

        string content = GetContentStreamContent();
        if (content.Length > 0) {
            CollectImagePlacementsAndForms(
                content,
                pageResources,
                0,
                pageTransform,
                pageHeight,
                placements,
                activeForms);
        }

        return placements.Count == 0 ? Array.Empty<PdfImagePlacement>() : placements.AsReadOnly();
    }

    private static void AddImagePlacements(OfficeDrawing drawing, double pageHeight, IReadOnlyList<PdfImagePlacement> placements, IReadOnlyList<PdfExtractedImage> images) {
        for (int i = 0; i < placements.Count; i++) {
            PdfImagePlacement placement = placements[i];
            PdfExtractedImage? image = FindImage(images, placement);
            if (image == null || !image.IsImageFile || placement.Width <= 0D || placement.Height <= 0D) {
                continue;
            }

            if (!TryCreateImageProjection(placement, pageHeight, drawing.Width, drawing.Height, out OfficeImageProjection projection)) {
                continue;
            }

            if (TryAddClippedImagePlacement(drawing, placement, image, projection)) {
                continue;
            }

            drawing.AddImage(image.Bytes, image.MimeType, projection, opacity: placement.ImageOpacity ?? 1D);
        }
    }

    private static bool TryAddClippedImagePlacement(OfficeDrawing drawing, PdfImagePlacement placement, PdfExtractedImage image, OfficeImageProjection projection) {
        if (!placement.ClipPath.HasValue || placement.ClipPath.Value.IsRectangle) {
            return false;
        }

        PdfPageClipPath clip = placement.ClipPath.Value;
        OfficeClipPath? clipPath = clip.ToOfficeClipPath(clip.X, clip.Y);
        if (clipPath == null) {
            return false;
        }

        drawing.AddClippedImage(image.Bytes, image.MimeType, projection, clip.X, clip.Y, clipPath, opacity: placement.ImageOpacity ?? 1D);
        return true;
    }

    private static bool TryCreateImageProjection(PdfImagePlacement placement, double pageHeight, double drawingWidth, double drawingHeight, out OfficeImageProjection projection) {
        if (!placement.ClipPath.HasValue &&
            TryCreateTransformedImageProjection(placement, pageHeight, drawingWidth, drawingHeight, out projection)) {
            return true;
        }

        double imageX = placement.X;
        double imageY = pageHeight - placement.Y - placement.Height;
        projection = default;

        if (!HasPositiveArea(imageX, imageY, placement.Width, placement.Height, drawingWidth, drawingHeight)) {
            return false;
        }

        PdfPageClipPath? clip = placement.ClipPath;
        if (!clip.HasValue || !clip.Value.IsRectangle) {
            projection = new OfficeImageProjection(new OfficeImagePlacement(imageX, imageY, placement.Width, placement.Height));
            return true;
        }

        double clipLeft = clip.Value.X;
        double clipTop = clip.Value.Y;
        double clipRight = clipLeft + clip.Value.Width;
        double clipBottom = clipTop + clip.Value.Height;
        double imageRight = imageX + placement.Width;
        double imageBottom = imageY + placement.Height;
        double visibleLeft = Math.Max(imageX, clipLeft);
        double visibleTop = Math.Max(imageY, clipTop);
        double visibleRight = Math.Min(imageRight, clipRight);
        double visibleBottom = Math.Min(imageBottom, clipBottom);
        double visibleWidth = visibleRight - visibleLeft;
        double visibleHeight = visibleBottom - visibleTop;
        if (visibleWidth <= 0D || visibleHeight <= 0D ||
            !HasPositiveArea(visibleLeft, visibleTop, visibleWidth, visibleHeight, drawingWidth, drawingHeight)) {
            return false;
        }

        var crop = OfficeImageSourceCrop.FromClampedFractions(
            (visibleLeft - imageX) / placement.Width,
            (visibleTop - imageY) / placement.Height,
            (imageRight - visibleRight) / placement.Width,
            (imageBottom - visibleBottom) / placement.Height);
        projection = new OfficeImageProjection(new OfficeImagePlacement(visibleLeft, visibleTop, visibleWidth, visibleHeight), crop);
        return true;
    }

    private static bool TryCreateTransformedImageProjection(PdfImagePlacement placement, double pageHeight, double drawingWidth, double drawingHeight, out OfficeImageProjection projection) {
        projection = default;
        double m11 = placement.A;
        double m12 = -placement.B;
        double m21 = -placement.C;
        double m22 = placement.D;
        double offsetX = placement.C + placement.E;
        double offsetY = pageHeight - placement.D - placement.F;
        double width = Math.Sqrt((m11 * m11) + (m12 * m12));
        double height = Math.Sqrt((m21 * m21) + (m22 * m22));
        if (width <= 0D || height <= 0D) {
            return false;
        }

        double dot = (m11 * m21) + (m12 * m22);
        if (!NearlyEqual(dot, 0D)) {
            return false;
        }

        return TryCreateImageProjectionCandidate(m11, m12, m21, m22, offsetX, offsetY, width, height, flipHorizontal: false, flipVertical: false, drawingWidth, drawingHeight, out projection) ||
               TryCreateImageProjectionCandidate(m11, m12, m21, m22, offsetX, offsetY, width, height, flipHorizontal: true, flipVertical: false, drawingWidth, drawingHeight, out projection) ||
               TryCreateImageProjectionCandidate(m11, m12, m21, m22, offsetX, offsetY, width, height, flipHorizontal: false, flipVertical: true, drawingWidth, drawingHeight, out projection) ||
               TryCreateImageProjectionCandidate(m11, m12, m21, m22, offsetX, offsetY, width, height, flipHorizontal: true, flipVertical: true, drawingWidth, drawingHeight, out projection);
    }

    private static bool TryCreateImageProjectionCandidate(
        double m11,
        double m12,
        double m21,
        double m22,
        double offsetX,
        double offsetY,
        double width,
        double height,
        bool flipHorizontal,
        bool flipVertical,
        double drawingWidth,
        double drawingHeight,
        out OfficeImageProjection projection) {
        projection = default;
        double columnSign = flipHorizontal ? -1D : 1D;
        double rowSign = flipVertical ? -1D : 1D;
        double cos = m11 / (columnSign * width);
        double sin = m12 / (columnSign * width);
        double baseColumnX = width * cos;
        double baseColumnY = width * sin;
        double baseRowX = -height * sin;
        double baseRowY = height * cos;
        if (!NearlyEqual(m21, rowSign * baseRowX) ||
            !NearlyEqual(m22, rowSign * baseRowY)) {
            return false;
        }

        double unflippedOffsetX = offsetX;
        double unflippedOffsetY = offsetY;
        if (flipHorizontal) {
            unflippedOffsetX -= baseColumnX;
            unflippedOffsetY -= baseColumnY;
        }

        if (flipVertical) {
            unflippedOffsetX -= baseRowX;
            unflippedOffsetY -= baseRowY;
        }

        double x = unflippedOffsetX - (width / 2D) + (cos * width / 2D) - (sin * height / 2D);
        double y = unflippedOffsetY - (height / 2D) + (sin * width / 2D) + (cos * height / 2D);
        if (!IsFinite(x) || !IsFinite(y)) {
            return false;
        }

        double rotationDegrees = Math.Atan2(sin, cos) * 180D / Math.PI;
        projection = new OfficeImageProjection(
            new OfficeImagePlacement(x, y, width, height),
            rotationDegrees: rotationDegrees,
            flipHorizontal: flipHorizontal,
            flipVertical: flipVertical);
        (double left, double top, double right, double bottom) = projection.GetDestinationBounds();
        return HasPositiveArea(left, top, right - left, bottom - top, drawingWidth, drawingHeight);
    }

    private static PdfExtractedImage? FindImage(IReadOnlyList<PdfExtractedImage> images, PdfImagePlacement placement) {
        for (int i = 0; i < images.Count; i++) {
            PdfExtractedImage image = images[i];
            if (string.Equals(image.ResourceName, placement.ResourceName, StringComparison.Ordinal) &&
                image.ObjectNumber == placement.ObjectNumber &&
                image.DirectStreamIdentity == placement.DirectStreamIdentity &&
                (!image.IsImageMask || image.ImageMaskColor.Equals(placement.ImageMaskColor))) {
                return image;
            }
        }

        return null;
    }

    private static bool HasPositiveArea(double x, double y, double width, double height, double maxWidth, double maxHeight) =>
        width > 0D &&
        height > 0D &&
        x >= 0D &&
        y >= 0D &&
        x + width <= maxWidth &&
        y + height <= maxHeight;

    private static double Clamp(double value, double min, double max) {
        if (value < min) {
            return min;
        }

        return value > max ? max : value;
    }

    private static bool NearlyEqual(double left, double right) => Math.Abs(left - right) <= 0.001D;
}

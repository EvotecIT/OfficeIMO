namespace OfficeIMO.Pdf;

/// <summary>
/// Represents a single page parsed from the PDF.
/// Provides access to plain text and basic text spans based on content stream operators.
/// </summary>
public sealed class PdfReadPage {
    private readonly PdfDictionary _pageDict;
    private readonly Dictionary<int, PdfIndirectObject> _objects;

    internal PdfReadPage(int objectNumber, PdfDictionary pageDict, Dictionary<int, PdfIndirectObject> objects) {
        ObjectNumber = objectNumber; _pageDict = pageDict; _objects = objects;
    }

    /// <summary>Underlying object number for the page.</summary>
    public int ObjectNumber { get; }

    /// <summary>Extracts plain text from this page without column reordering.</summary>
    public string ExtractText() {
        var spans = GetTextSpans();
        var opts = new TextLayoutEngine.Options { ForceSingleColumn = true };
        var lines = TextLayoutEngine.BuildLines(spans, opts);
        return TextLayoutEngine.EmitText(lines, TextLayoutEngine.DetectColumns(lines, GetPageSize().Width, opts), null);
    }

    /// <summary>
    /// Attempts to read page size from CropBox (or MediaBox) and returns width/height in points.
    /// Falls back to 612x792 (US Letter) when not present or malformed.
    /// </summary>
    public (double Width, double Height) GetPageSize() {
        var crop = GetInheritedValue("CropBox");
        if (TryParseBox(crop, out var cropSize)) return cropSize;

        var media = GetInheritedValue("MediaBox");
        if (TryParseBox(media, out var mediaSize)) return mediaSize;

        return (612, 792);
    }

    /// <summary>Gets inherited page rotation in degrees normalized to 0, 90, 180, or 270.</summary>
    public int GetRotationDegrees() {
        var rotate = GetInheritedValue("Rotate");
        if (rotate is PdfNumber number) {
            int degrees = (int)Math.Round(number.Value);
            degrees %= 360;
            if (degrees < 0) {
                degrees += 360;
            }

            return degrees;
        }

        return 0;
    }

    /// <summary>Gets text spans (text with position and font info) from this page.</summary>
    public IReadOnlyList<PdfTextSpan> GetTextSpans() {
        var spans = new List<PdfTextSpan>();
        var pageResources = ResolveDictionary(GetInheritedValue("Resources"));
        var pageDecoders = ResourceResolver.GetFontDecoders(_pageDict, _objects);
        var pageWidthProviders = ResourceResolver.GetFontWidthProviders(_pageDict, _objects);
        var activeForms = new HashSet<PdfStream>();

        string content = GetContentStreamContent();
        if (content.Length > 0) {
            CollectTextAndForms(
                content,
                pageResources,
                pageDecoders,
                pageWidthProviders,
                spans,
                activeForms);
        }

        return spans;
    }

    /// <summary>Reads simple URI and named-destination link annotations from this page.</summary>
    public IReadOnlyList<PdfLinkAnnotation> GetLinkAnnotations() {
        if (!_pageDict.Items.TryGetValue("Annots", out var annotsObject)) {
            return Array.Empty<PdfLinkAnnotation>();
        }

        var annotations = ResolveArray(annotsObject);
        if (annotations is null) {
            return Array.Empty<PdfLinkAnnotation>();
        }

        var result = new List<PdfLinkAnnotation>();
        foreach (var item in annotations.Items) {
            var annotation = ResolveDictionary(item);
            if (annotation is null ||
                annotation.Get<PdfName>("Subtype")?.Name != "Link" ||
                !TryReadRectangle(annotation.Items.TryGetValue("Rect", out var rectObject) ? rectObject : null, out var rect)) {
                continue;
            }

            var action = ResolveDictionary(annotation.Items.TryGetValue("A", out var actionObject) ? actionObject : null);
            TryGetString(annotation.Items.TryGetValue("Contents", out var contentsObject) ? contentsObject : null, out string? contents);

            if (action != null &&
                action.Get<PdfName>("S")?.Name == "URI" &&
                TryGetString(action.Items.TryGetValue("URI", out var uriObject) ? uriObject : null, out string? uri) &&
                Uri.TryCreate(uri, UriKind.Absolute, out _)) {
                result.Add(new PdfLinkAnnotation(uri!, contents, rect.X1, rect.Y1, rect.X2, rect.Y2));
                continue;
            }

            if (action != null &&
                action.Get<PdfName>("S")?.Name == "GoTo" &&
                TryGetDestinationName(action.Items.TryGetValue("D", out var actionDestination) ? actionDestination : null, out string? actionDestinationName)) {
                result.Add(new PdfLinkAnnotation(null, actionDestinationName, contents, rect.X1, rect.Y1, rect.X2, rect.Y2));
                continue;
            }

            if (TryGetDestinationName(annotation.Items.TryGetValue("Dest", out var directDestination) ? directDestination : null, out string? directDestinationName)) {
                result.Add(new PdfLinkAnnotation(null, directDestinationName, contents, rect.X1, rect.Y1, rect.X2, rect.Y2));
            }
        }

        return result.AsReadOnly();
    }

    /// <summary>Extracts image XObjects referenced by this page.</summary>
    public IReadOnlyList<PdfExtractedImage> GetImages() => GetImages(0);

    internal IReadOnlyList<PdfExtractedImage> GetImages(int pageNumber) {
        return ResourceResolver.GetImageXObjectsForPage(_pageDict, _objects, pageNumber);
    }

    internal List<string> GetUnsupportedContentStreamFilters() {
        var unsupported = new List<string>();
        var pageResources = ResolveDictionary(GetInheritedValue("Resources"));
        var activeForms = new HashSet<PdfStream>();
        var content = new System.Text.StringBuilder();
        bool canInspectFormInvocations = true;
        foreach (var stream in GetContentStreamObjects()) {
            AddUnsupportedFilters(stream, unsupported);
            if (Filters.StreamDecoder.GetUnsupportedFilters(stream.Dictionary, _objects).Count != 0) {
                canInspectFormInvocations = false;
                continue;
            }

            content.Append(PdfEncoding.Latin1GetString(DecodeIfNeeded(stream)));
        }

        if (canInspectFormInvocations && content.Length > 0) {
            CollectUnsupportedFormFilters(content.ToString(), pageResources, unsupported, activeForms);
        }

        return unsupported;
    }

    private void CollectUnsupportedFormFilters(
        string content,
        PdfDictionary? resources,
        List<string> unsupported,
        HashSet<PdfStream> activeForms) {
        foreach (var invocation in TextContentParser.ExtractFormInvocations(content)) {
            if (!TryGetFormStream(resources, invocation.Name, out var formStream)) {
                continue;
            }

            if (!activeForms.Add(formStream)) {
                continue;
            }

            try {
                AddUnsupportedFilters(formStream, unsupported);
                if (Filters.StreamDecoder.GetUnsupportedFilters(formStream.Dictionary, _objects).Count != 0) {
                    continue;
                }

                var formResources = ResolveDictionary(formStream.Dictionary.Items.TryGetValue("Resources", out var resObj) ? resObj : null) ?? resources;
                CollectUnsupportedFormFilters(PdfEncoding.Latin1GetString(DecodeIfNeeded(formStream)), formResources, unsupported, activeForms);
            } finally {
                activeForms.Remove(formStream);
            }
        }
    }

    private void AddUnsupportedFilters(PdfStream stream, List<string> unsupported) {
        foreach (string filterName in Filters.StreamDecoder.GetUnsupportedFilters(stream.Dictionary, _objects)) {
            if (!ContainsFilter(unsupported, filterName)) {
                unsupported.Add(filterName);
            }
        }
    }

    private void CollectTextAndForms(
        string content,
        PdfDictionary? resources,
        Dictionary<string, Func<byte[], string>> decoders,
        Dictionary<string, Func<byte[], double>> widthProviders,
        List<PdfTextSpan> spans,
        HashSet<PdfStream> activeForms) {
        string DecodeWithFont(string fontRes, byte[] bytes) =>
            decoders.TryGetValue(fontRes, out var dec) ? dec(bytes) : PdfWinAnsiEncoding.Decode(bytes);
        double SumWidth1000(string fontRes, byte[] bytes) =>
            widthProviders.TryGetValue(fontRes, out var wp) ? wp(bytes) : (bytes?.Length ?? 0) * 500.0;

        spans.AddRange(TextContentParser.Parse(content, DecodeWithFont, SumWidth1000));

        foreach (var invocation in TextContentParser.ExtractFormInvocations(content)) {
            if (!TryGetFormStream(resources, invocation.Name, out var formStream)) {
                continue;
            }

            if (!activeForms.Add(formStream)) {
                continue;
            }

            try {
                var formDict = formStream.Dictionary;
                var formResources = ResolveDictionary(formDict.Items.TryGetValue("Resources", out var resObj) ? resObj : null) ?? resources;
                var formDecoders = MergeDecoders(decoders, ResourceResolver.GetFontDecodersForForm(formDict, _objects));
                var formWidths = MergeWidthProviders(widthProviders, ResourceResolver.GetFontWidthProviders(formDict, _objects));
                var combinedTransform = ApplyFormMatrix(invocation.Transform, formDict);
                var formContent = WrapContentWithTransform(PdfEncoding.Latin1GetString(DecodeIfNeeded(formStream)), combinedTransform);

                CollectTextAndForms(formContent, formResources, formDecoders, formWidths, spans, activeForms);
            } finally {
                activeForms.Remove(formStream);
            }
        }
    }

    private bool TryGetFormStream(PdfDictionary? resources, string name, out PdfStream formStream) {
        if (resources is null || !resources.Items.TryGetValue("XObject", out var xoObj)) {
            formStream = null!;
            return false;
        }

        var xoDict = ResolveDictionary(xoObj);
        if (xoDict is null || !xoDict.Items.TryGetValue(name, out var formObj)) {
            formStream = null!;
            return false;
        }

        if (formObj is PdfReference formRef &&
            PdfObjectLookup.TryGet(_objects, formRef, out var indirectForm) &&
            indirectForm.Value is PdfStream stream &&
            string.Equals(stream.Dictionary.Get<PdfName>("Subtype")?.Name, "Form", StringComparison.Ordinal)) {
            formStream = stream;
            return true;
        }

        if (formObj is PdfStream directStream &&
            string.Equals(directStream.Dictionary.Get<PdfName>("Subtype")?.Name, "Form", StringComparison.Ordinal)) {
            formStream = directStream;
            return true;
        }

        formStream = null!;
        return false;
    }

    private static Dictionary<string, Func<byte[], string>> MergeDecoders(
        Dictionary<string, Func<byte[], string>> parent,
        Dictionary<string, Func<byte[], string>> local) {
        var merged = new Dictionary<string, Func<byte[], string>>(parent, StringComparer.Ordinal);
        foreach (var entry in local) {
            merged[entry.Key] = entry.Value;
        }

        return merged;
    }

    private static Dictionary<string, Func<byte[], double>> MergeWidthProviders(
        Dictionary<string, Func<byte[], double>> parent,
        Dictionary<string, Func<byte[], double>> local) {
        var merged = new Dictionary<string, Func<byte[], double>>(parent, StringComparer.Ordinal);
        foreach (var entry in local) {
            merged[entry.Key] = entry.Value;
        }

        return merged;
    }

    private static string WrapContentWithTransform(string content, Matrix2D transform) {
        string prefix = string.Format(
            System.Globalization.CultureInfo.InvariantCulture,
            "q {0} {1} {2} {3} {4} {5} cm ",
            transform.A,
            transform.B,
            transform.C,
            transform.D,
            transform.E,
            transform.F);
        return prefix + content + " Q";
    }

    private static Matrix2D ApplyFormMatrix(Matrix2D invocationTransform, PdfDictionary? formDict) {
        if (formDict is null ||
            !formDict.Items.TryGetValue("Matrix", out var matrixObj) ||
            matrixObj is not PdfArray arr ||
            arr.Items.Count < 6) {
            return invocationTransform;
        }

        var formMatrix = new Matrix2D(
            (arr.Items[0] as PdfNumber)?.Value ?? 1,
            (arr.Items[1] as PdfNumber)?.Value ?? 0,
            (arr.Items[2] as PdfNumber)?.Value ?? 0,
            (arr.Items[3] as PdfNumber)?.Value ?? 1,
            (arr.Items[4] as PdfNumber)?.Value ?? 0,
            (arr.Items[5] as PdfNumber)?.Value ?? 0);

        return Matrix2D.Multiply(invocationTransform, formMatrix);
    }

    private PdfObject? GetInheritedValue(string key) {
        PdfDictionary? current = _pageDict;
        int guard = 0;
        while (current is not null && guard++ < 100) {
            if (current.Items.TryGetValue(key, out var value)) {
                return value;
            }

            if (!current.Items.TryGetValue("Parent", out var parentObj) ||
                parentObj is not PdfReference parentRef ||
                !PdfObjectLookup.TryGet(_objects, parentRef, out var parentIndirect) ||
                parentIndirect.Value is not PdfDictionary parentDict) {
                break;
            }

            current = parentDict;
        }

        return null;
    }

    private PdfDictionary? ResolveDictionary(PdfObject? obj) {
        if (obj is PdfDictionary dictionary) {
            return dictionary;
        }

        if (obj is PdfReference reference &&
            PdfObjectLookup.TryGet(_objects, reference, out var indirect) &&
            indirect.Value is PdfDictionary referencedDictionary) {
            return referencedDictionary;
        }

        return null;
    }

    private PdfObject? ResolveObject(PdfObject? obj) {
        return PdfObjectLookup.Resolve(_objects, obj);
    }

    private PdfArray? ResolveArray(PdfObject? obj) {
        if (obj is PdfArray array) {
            return array;
        }

        if (obj is PdfReference reference &&
            PdfObjectLookup.TryGet(_objects, reference, out var indirect) &&
            indirect.Value is PdfArray referencedArray) {
            return referencedArray;
        }

        return null;
    }

    private bool TryGetString(PdfObject? obj, out string? value) {
        if (ResolveObject(obj) is PdfStringObj text) {
            value = text.Value;
            return true;
        }

        value = null;
        return false;
    }

    private bool TryGetDestinationName(PdfObject? obj, out string? value) {
        switch (ResolveObject(obj)) {
            case PdfStringObj text when !string.IsNullOrEmpty(text.Value):
                value = text.Value;
                return true;
            case PdfName name when !string.IsNullOrEmpty(name.Name):
                value = name.Name;
                return true;
            default:
                value = null;
                return false;
        }
    }

    private bool TryReadRectangle(PdfObject? obj, out (double X1, double Y1, double X2, double Y2) rect) {
        rect = default;
        var array = ResolveArray(obj);
        if (array is null || array.Items.Count < 4) {
            return false;
        }

        if (ResolveObject(array.Items[0]) is not PdfNumber x1 ||
            ResolveObject(array.Items[1]) is not PdfNumber y1 ||
            ResolveObject(array.Items[2]) is not PdfNumber x2 ||
            ResolveObject(array.Items[3]) is not PdfNumber y2) {
            return false;
        }

        double left = Math.Min(x1.Value, x2.Value);
        double right = Math.Max(x1.Value, x2.Value);
        double bottom = Math.Min(y1.Value, y2.Value);
        double top = Math.Max(y1.Value, y2.Value);
        if (double.IsNaN(left) || double.IsInfinity(left) ||
            double.IsNaN(right) || double.IsInfinity(right) ||
            double.IsNaN(bottom) || double.IsInfinity(bottom) ||
            double.IsNaN(top) || double.IsInfinity(top) ||
            right <= left ||
            top <= bottom) {
            return false;
        }

        rect = (left, bottom, right, top);
        return true;
    }

    private bool TryParseBox(PdfObject? box, out (double Width, double Height) size) {
        var arr = ResolveArray(box);
        if (arr is not null &&
            arr.Items.Count >= 4 &&
            arr.Items[0] is PdfNumber llx &&
            arr.Items[1] is PdfNumber lly &&
            arr.Items[2] is PdfNumber urx &&
            arr.Items[3] is PdfNumber ury) {
            double width = urx.Value - llx.Value;
            double height = ury.Value - lly.Value;
            if (width > 0 && height > 0) {
                size = (width, height);
                return true;
            }
        }

        size = default;
        return false;
    }

    private static double GlyphWidthEmForBase(string baseFont) {
        if (string.IsNullOrEmpty(baseFont)) return 0.55;
        if (ContainsIgnoreCase(baseFont, "courier")) return 0.6;
        if (ContainsIgnoreCase(baseFont, "times")) return 0.5;
        if (ContainsIgnoreCase(baseFont, "helvetica")) return 0.55;
        return 0.55;
    }

    private static bool ContainsIgnoreCase(string source, string value) {
#if NET8_0_OR_GREATER
        return source.Contains(value, System.StringComparison.OrdinalIgnoreCase);
#else
        return source.IndexOf(value, System.StringComparison.OrdinalIgnoreCase) >= 0;
#endif
    }

    /// <summary>
    /// Returns decoded page content with stream arrays concatenated in PDF processing order.
    /// </summary>
    private string GetContentStreamContent() {
        var builder = new System.Text.StringBuilder();
        foreach (var stream in GetContentStreamObjects()) {
            builder.Append(PdfEncoding.Latin1GetString(DecodeIfNeeded(stream)));
        }

        return builder.ToString();
    }

    private List<PdfStream> GetContentStreamObjects() {
        var result = new List<PdfStream>();
        var contents = _pageDict.Items.TryGetValue("Contents", out var obj) ? obj : null;
        if (contents is PdfReference r) {
            if (PdfObjectLookup.TryGet(_objects, r, out var ind) && ind.Value is PdfStream s) {
                result.Add(s);
                return result;
            }
        }

        var contentArray = ResolveArray(contents);
        if (contentArray is null) {
            return result;
        }

        foreach (var item in contentArray.Items) {
            if (item is PdfReference rr &&
                PdfObjectLookup.TryGet(_objects, rr, out var ind2) &&
                ind2.Value is PdfStream s2) {
                result.Add(s2);
            } else if (item is PdfStream directStream) {
                result.Add(directStream);
            }
        }

        return result;
    }

    private static bool ContainsFilter(List<string> filters, string filterName) {
        for (int i = 0; i < filters.Count; i++) {
            if (string.Equals(filters[i], filterName, StringComparison.Ordinal)) {
                return true;
            }
        }

        return false;
    }

    private byte[] DecodeIfNeeded(PdfStream s) {
        return Filters.StreamDecoder.Decode(s.Dictionary, s.Data, _objects);
    }
}

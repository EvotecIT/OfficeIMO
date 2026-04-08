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

    /// <summary>Gets text spans (text with position and font info) from this page.</summary>
    public IReadOnlyList<PdfTextSpan> GetTextSpans() {
        var spans = new List<PdfTextSpan>();
        var pageResources = ResolveDictionary(GetInheritedValue("Resources"));
        var pageDecoders = ResourceResolver.GetFontDecoders(_pageDict, _objects);
        var pageWidthProviders = ResourceResolver.GetFontWidthProviders(_pageDict, _objects);
        var activeForms = new HashSet<int>();

        foreach (var stream in GetContentStreams()) {
            CollectTextAndForms(
                PdfEncoding.Latin1GetString(stream),
                pageResources,
                pageDecoders,
                pageWidthProviders,
                spans,
                activeForms);
        }

        return spans;
    }

    private void CollectTextAndForms(
        string content,
        PdfDictionary? resources,
        Dictionary<string, Func<byte[], string>> decoders,
        Dictionary<string, Func<byte[], double>> widthProviders,
        List<PdfTextSpan> spans,
        HashSet<int> activeForms) {
        string DecodeWithFont(string fontRes, byte[] bytes) =>
            decoders.TryGetValue(fontRes, out var dec) ? dec(bytes) : PdfWinAnsiEncoding.Decode(bytes);
        double SumWidth1000(string fontRes, byte[] bytes) =>
            widthProviders.TryGetValue(fontRes, out var wp) ? wp(bytes) : (bytes?.Length ?? 0) * 500.0;

        spans.AddRange(TextContentParser.Parse(content, DecodeWithFont, SumWidth1000));

        foreach (var invocation in TextContentParser.ExtractFormInvocations(content)) {
            if (!TryGetFormStream(resources, invocation.Name, out var formStream, out int formObjectNumber)) {
                continue;
            }

            bool trackRecursion = formObjectNumber > 0;
            if (trackRecursion && !activeForms.Add(formObjectNumber)) {
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
                if (trackRecursion) {
                    activeForms.Remove(formObjectNumber);
                }
            }
        }
    }

    private bool TryGetFormStream(PdfDictionary? resources, string name, out PdfStream formStream, out int objectNumber) {
        if (resources is null || !resources.Items.TryGetValue("XObject", out var xoObj)) {
            formStream = null!;
            objectNumber = 0;
            return false;
        }

        var xoDict = ResolveDictionary(xoObj);
        if (xoDict is null || !xoDict.Items.TryGetValue(name, out var formObj)) {
            formStream = null!;
            objectNumber = 0;
            return false;
        }

        if (formObj is PdfReference formRef &&
            _objects.TryGetValue(formRef.ObjectNumber, out var indirectForm) &&
            indirectForm.Value is PdfStream stream &&
            string.Equals(stream.Dictionary.Get<PdfName>("Subtype")?.Name, "Form", StringComparison.Ordinal)) {
            formStream = stream;
            objectNumber = formRef.ObjectNumber;
            return true;
        }

        if (formObj is PdfStream directStream &&
            string.Equals(directStream.Dictionary.Get<PdfName>("Subtype")?.Name, "Form", StringComparison.Ordinal)) {
            formStream = directStream;
            objectNumber = 0;
            return true;
        }

        formStream = null!;
        objectNumber = 0;
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
                !_objects.TryGetValue(parentRef.ObjectNumber, out var parentIndirect) ||
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
            _objects.TryGetValue(reference.ObjectNumber, out var indirect) &&
            indirect.Value is PdfDictionary referencedDictionary) {
            return referencedDictionary;
        }

        return null;
    }

    private PdfArray? ResolveArray(PdfObject? obj) {
        if (obj is PdfArray array) {
            return array;
        }

        if (obj is PdfReference reference &&
            _objects.TryGetValue(reference.ObjectNumber, out var indirect) &&
            indirect.Value is PdfArray referencedArray) {
            return referencedArray;
        }

        return null;
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
    /// Returns a shallow list of content stream bytes for the page (handles single or array of streams).
    /// </summary>
    private List<byte[]> GetContentStreams() {
        var result = new List<byte[]>();
        var contents = _pageDict.Items.TryGetValue("Contents", out var obj) ? obj : null;
        if (contents is PdfReference r) {
            if (_objects.TryGetValue(r.ObjectNumber, out var ind) && ind.Value is PdfStream s) {
                result.Add(DecodeIfNeeded(s));
                return result;
            }
        }

        var contentArray = ResolveArray(contents);
        if (contentArray is null) {
            return result;
        }

        foreach (var item in contentArray.Items) {
            if (item is PdfReference rr &&
                _objects.TryGetValue(rr.ObjectNumber, out var ind2) &&
                ind2.Value is PdfStream s2) {
                result.Add(DecodeIfNeeded(s2));
            } else if (item is PdfStream directStream) {
                result.Add(DecodeIfNeeded(directStream));
            }
        }

        return result;
    }

    private byte[] DecodeIfNeeded(PdfStream s) {
        return Filters.StreamDecoder.Decode(s.Dictionary, s.Data, _objects);
    }
}

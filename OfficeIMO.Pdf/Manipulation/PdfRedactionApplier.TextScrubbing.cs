using OfficeIMO.Pdf.Filters;
using System.Globalization;
using System.Text.RegularExpressions;

namespace OfficeIMO.Pdf;

public static partial class PdfRedactionApplier {
    private const double RedactionFallbackTextHeight = 18D;

    private static bool RemoveMatchedTextObjects(
        Dictionary<int, PdfIndirectObject> objects,
        PdfDictionary pageDictionary,
        IReadOnlyList<PdfRedactionMatch> matches) {
        RedactionTextTarget[] textTargets = BuildTextTargets(matches);
        if (textTargets.Length == 0 ||
            !pageDictionary.Items.TryGetValue("Contents", out PdfObject? contentsObject)) {
            return false;
        }

        bool changed = false;
        Dictionary<string, Func<byte[], string>> fontDecoders = ResourceResolver.GetFontDecoders(pageDictionary, objects);
        foreach (PdfReference reference in EnumerateContentReferences(objects, contentsObject)) {
            if (!PdfObjectLookup.TryGet(objects, reference, out PdfIndirectObject? indirect) ||
                indirect.Value is not PdfStream stream ||
                stream.DecodingFailed) {
                continue;
            }

            byte[] contentBytes = StreamDecoder.Decode(stream.Dictionary, stream.Data, objects);
            string content = PdfEncoding.Latin1GetString(contentBytes);
            string scrubbed = ScrubTextObjects(content, textTargets, fontDecoders, new[] { Matrix2D.Identity });
            if (string.Equals(content, scrubbed, StringComparison.Ordinal)) {
                continue;
            }

            objects[reference.ObjectNumber] = new PdfIndirectObject(reference.ObjectNumber, reference.Generation, new PdfStream(CleanStreamDictionary(stream.Dictionary), PdfEncoding.Latin1GetBytes(scrubbed)));
            changed = true;
        }

        return ScrubMatchedFormXObjects(objects, pageDictionary, contentsObject, textTargets, fontDecoders) || changed;
    }

    private static RedactionTextTarget[] BuildTextTargets(IReadOnlyList<PdfRedactionMatch> matches) {
        return matches
            .Where(match => match.Kind == PdfRedactionMatchKind.TextBlock && !string.IsNullOrWhiteSpace(match.Text))
            .Select(match => new RedactionTextTarget(
                NormalizeText(match.Text!),
                match.X,
                match.Y,
                match.Width,
                match.Height))
            .Where(target => target.Text.Length > 0)
            .ToArray();
    }

    private static bool ScrubMatchedFormXObjects(
        Dictionary<int, PdfIndirectObject> objects,
        PdfDictionary pageDictionary,
        PdfObject contentsObject,
        RedactionTextTarget[] textTargets,
        IReadOnlyDictionary<string, Func<byte[], string>> pageFontDecoders) {
        PdfDictionary? resources = GetInheritedDictionary(objects, pageDictionary, "Resources");
        if (resources is null) {
            return false;
        }

        bool changed = false;
        foreach (PdfReference reference in EnumerateContentReferences(objects, contentsObject)) {
            if (!PdfObjectLookup.TryGet(objects, reference, out PdfIndirectObject? indirect) ||
                indirect.Value is not PdfStream stream ||
                stream.DecodingFailed) {
                continue;
            }

            string content = PdfEncoding.Latin1GetString(StreamDecoder.Decode(stream.Dictionary, stream.Data, objects));
            changed = ScrubFormInvocations(objects, resources, content, textTargets, pageFontDecoders, new[] { Matrix2D.Identity }, new HashSet<int>()) || changed;
        }

        return changed;
    }

    private static bool ScrubFormInvocations(
        Dictionary<int, PdfIndirectObject> objects,
        PdfDictionary resources,
        string content,
        RedactionTextTarget[] textTargets,
        IReadOnlyDictionary<string, Func<byte[], string>> parentFontDecoders,
        IReadOnlyList<Matrix2D> parentTransforms,
        HashSet<int> activeForms) {
        bool changed = false;
        foreach (TextContentParser.FormInvocation invocation in TextContentParser.ExtractFormInvocations(content)) {
            if (!TryGetFormXObject(objects, resources, invocation.Name, out PdfReference reference, out PdfStream formStream) ||
                formStream.DecodingFailed ||
                !activeForms.Add(reference.ObjectNumber)) {
                continue;
            }

            try {
                PdfDictionary formResources = ResolveDictionary(objects, formStream.Dictionary.Items.TryGetValue("Resources", out PdfObject? resourcesObject) ? resourcesObject : null) ?? resources;
                Dictionary<string, Func<byte[], string>> formDecoders = MergeDecoders(parentFontDecoders, ResourceResolver.GetFontDecodersForForm(formStream.Dictionary, objects));
                Matrix2D[] effectiveTransforms = parentTransforms
                    .Select(parent => ApplyFormMatrix(Matrix2D.Multiply(parent, invocation.Transform), formStream.Dictionary))
                    .ToArray();
                byte[] formBytes = StreamDecoder.Decode(formStream.Dictionary, formStream.Data, objects);
                string formContent = PdfEncoding.Latin1GetString(formBytes);
                string scrubbed = ScrubTextObjects(formContent, textTargets, formDecoders, effectiveTransforms);
                if (!string.Equals(formContent, scrubbed, StringComparison.Ordinal)) {
                    objects[reference.ObjectNumber] = new PdfIndirectObject(reference.ObjectNumber, reference.Generation, new PdfStream(CleanStreamDictionary(formStream.Dictionary), PdfEncoding.Latin1GetBytes(scrubbed)));
                    formContent = scrubbed;
                    changed = true;
                }

                changed = ScrubFormInvocations(objects, formResources, formContent, textTargets, formDecoders, effectiveTransforms, activeForms) || changed;
            } finally {
                activeForms.Remove(reference.ObjectNumber);
            }
        }

        return changed;
    }

    private static bool TryGetFormXObject(Dictionary<int, PdfIndirectObject> objects, PdfDictionary resources, string name, out PdfReference reference, out PdfStream stream) {
        if (resources.Items.TryGetValue("XObject", out PdfObject? xObjectValue) &&
            ResolveDictionary(objects, xObjectValue) is PdfDictionary xObjects &&
            xObjects.Items.TryGetValue(name, out PdfObject? value) &&
            value is PdfReference formReference &&
            PdfObjectLookup.TryGet(objects, formReference, out PdfIndirectObject? indirect) &&
            indirect.Value is PdfStream formStream &&
            string.Equals(formStream.Dictionary.Get<PdfName>("Subtype")?.Name, "Form", StringComparison.Ordinal)) {
            reference = formReference;
            stream = formStream;
            return true;
        }

        reference = default!;
        stream = default!;
        return false;
    }

    private static PdfDictionary? ResolveDictionary(Dictionary<int, PdfIndirectObject> objects, PdfObject? value) {
        if (value is PdfDictionary dictionary) {
            return dictionary;
        }

        return value is PdfReference reference &&
            PdfObjectLookup.TryGet(objects, reference, out PdfIndirectObject? indirect) &&
            indirect.Value is PdfDictionary referencedDictionary
            ? referencedDictionary
            : null;
    }

    private static string ScrubTextObjects(
        string content,
        RedactionTextTarget[] targets,
        IReadOnlyDictionary<string, Func<byte[], string>> fontDecoders,
        IReadOnlyList<Matrix2D> transforms) {
        List<RedactionTextObject> textObjects = CollectTextObjects(content, fontDecoders, transforms);
        if (textObjects.Count == 0) {
            return content;
        }

        var removeByIndex = new HashSet<int>();
        for (int targetIndex = 0; targetIndex < targets.Length; targetIndex++) {
            MarkMatchingTextObjects(textObjects, targets[targetIndex], removeByIndex);
        }

        if (removeByIndex.Count == 0) {
            return content;
        }

        return TextObjectRegex.Replace(content, match => removeByIndex.Contains(match.Index) ? string.Empty : match.Value);
    }

    private static List<RedactionTextObject> CollectTextObjects(
        string content,
        IReadOnlyDictionary<string, Func<byte[], string>> fontDecoders,
        IReadOnlyList<Matrix2D> transforms) {
        var textObjects = new List<RedactionTextObject>();
        foreach (Match match in TextObjectRegex.Matches(content)) {
            string shownText = NormalizeText(ExtractTextFromTextObject(match.Value, fontDecoders));
            if (shownText.Length == 0) {
                continue;
            }

            textObjects.Add(BuildRedactionTextObject(match.Index, match.Value, shownText, fontDecoders, transforms));
        }

        return textObjects;
    }

    private static RedactionTextObject BuildRedactionTextObject(
        int index,
        string textObject,
        string shownText,
        IReadOnlyDictionary<string, Func<byte[], string>> fontDecoders,
        IReadOnlyList<Matrix2D> transforms) {
        RedactionTextBounds? bounds = null;
        for (int transformIndex = 0; transformIndex < transforms.Count; transformIndex++) {
            string transformedContent = WrapContentWithTransform(textObject, transforms[transformIndex]);
            List<PdfTextSpan> spans = ParseTextSpans(transformedContent, fontDecoders);
            for (int spanIndex = 0; spanIndex < spans.Count; spanIndex++) {
                bounds = AddSpanBounds(bounds, spans[spanIndex]);
            }
        }

        return new RedactionTextObject(index, shownText, bounds);
    }

    private static List<PdfTextSpan> ParseTextSpans(string content, IReadOnlyDictionary<string, Func<byte[], string>> fontDecoders) {
        string DecodeWithFont(string fontResource, byte[] bytes) =>
            fontDecoders.TryGetValue(fontResource, out Func<byte[], string>? decoder)
                ? decoder(bytes)
                : PdfWinAnsiEncoding.Decode(bytes);
        double SumWidth1000(string fontResource, byte[] bytes) =>
            bytes is null ? 0D : bytes.Length * 500D;

        return TextContentParser.Parse(content, DecodeWithFont, SumWidth1000);
    }

    private static RedactionTextBounds AddSpanBounds(RedactionTextBounds? current, PdfTextSpan span) {
        double left = Math.Min(span.X, span.X + Math.Max(span.Advance, 0D));
        double right = Math.Max(span.X, span.X + Math.Max(span.Advance, 0D));
        double bottom = span.Y - Math.Max(span.FontSize, 1D);
        double top = span.Y + Math.Max(span.FontSize * 0.25D, 1D);
        if (current is null) {
            return new RedactionTextBounds(left, bottom, right, top);
        }

        return new RedactionTextBounds(
            Math.Min(current.Value.Left, left),
            Math.Min(current.Value.Bottom, bottom),
            Math.Max(current.Value.Right, right),
            Math.Max(current.Value.Top, top));
    }

    private static void MarkMatchingTextObjects(
        List<RedactionTextObject> textObjects,
        RedactionTextTarget target,
        HashSet<int> removeByIndex) {
        for (int start = 0; start < textObjects.Count; start++) {
            if (ContainsOrdinal(textObjects[start].Text, target.Text)) {
                if (IntersectsTarget(textObjects[start], target)) {
                    removeByIndex.Add(textObjects[start].Index);
                }

                continue;
            }

            var builder = new StringBuilder();
            RedactionTextBounds? bounds = null;
            for (int end = start; end < textObjects.Count; end++) {
                if (builder.Length > 0) {
                    builder.Append(' ');
                }

                builder.Append(textObjects[end].Text);
                bounds = MergeBounds(bounds, textObjects[end].Bounds);
                string combined = NormalizeText(builder.ToString());
                if (!combined.StartsWith(target.Text, StringComparison.Ordinal)) {
                    continue;
                }

                if (!IntersectsTarget(bounds, target)) {
                    break;
                }

                for (int remove = start; remove <= end; remove++) {
                    removeByIndex.Add(textObjects[remove].Index);
                }

                break;
            }
        }
    }

    private static bool IntersectsTarget(RedactionTextObject textObject, RedactionTextTarget target) =>
        IntersectsTarget(textObject.Bounds, target);

    private static bool IntersectsTarget(RedactionTextBounds? bounds, RedactionTextTarget target) {
        if (bounds is null) {
            return true;
        }

        return Intersects(
            target.X,
            target.Y,
            target.Width,
            target.Height <= 0D ? RedactionFallbackTextHeight : target.Height,
            bounds.Value.Left,
            bounds.Value.Bottom,
            bounds.Value.Right - bounds.Value.Left,
            bounds.Value.Top - bounds.Value.Bottom);
    }

    private static RedactionTextBounds? MergeBounds(RedactionTextBounds? left, RedactionTextBounds? right) {
        if (left is null) {
            return right;
        }

        if (right is null) {
            return left;
        }

        return new RedactionTextBounds(
            Math.Min(left.Value.Left, right.Value.Left),
            Math.Min(left.Value.Bottom, right.Value.Bottom),
            Math.Max(left.Value.Right, right.Value.Right),
            Math.Max(left.Value.Top, right.Value.Top));
    }

    private static string ExtractTextFromTextObject(
        string textObject,
        IReadOnlyDictionary<string, Func<byte[], string>> fontDecoders) {
        var builder = new StringBuilder();
        string? currentFont = null;
        int cursor = 0;
        foreach (RedactionTextStringToken token in EnumerateTextStringTokens(textObject)) {
            currentFont = ReadLastFontName(textObject.Substring(cursor, token.Index - cursor)) ?? currentFont;
            if (token.IsHex) {
                builder.Append(DecodeHexString(token.Value, currentFont, fontDecoders));
            } else {
                builder.Append(DecodeLiteralString(token.Value, currentFont, fontDecoders));
            }

            cursor = token.Index + token.Length;
        }

        return builder.ToString();
    }

    private static IEnumerable<RedactionTextStringToken> EnumerateTextStringTokens(string value) {
        for (int i = 0; i < value.Length;) {
            char current = value[i];
            if (current == '(') {
                if (TryReadLiteralStringToken(value, i, out RedactionTextStringToken token)) {
                    yield return token;
                    i += token.Length;
                    continue;
                }

                yield break;
            }

            if (current == '<' && (i + 1 >= value.Length || value[i + 1] != '<')) {
                if (TryReadHexStringToken(value, i, out RedactionTextStringToken token)) {
                    yield return token;
                    i += token.Length;
                    continue;
                }
            }

            i++;
        }
    }

    private static bool TryReadLiteralStringToken(string value, int start, out RedactionTextStringToken token) {
        int depth = 1;
        bool escaped = false;
        int index = start + 1;
        while (index < value.Length && depth > 0) {
            char current = value[index++];
            if (escaped) {
                escaped = false;
            } else if (current == '\\') {
                escaped = true;
            } else if (current == '(') {
                depth++;
            } else if (current == ')') {
                depth--;
            }
        }

        if (depth != 0) {
            token = default;
            return false;
        }

        int length = index - start;
        token = new RedactionTextStringToken(start, length, isHex: false, value.Substring(start, length));
        return true;
    }

    private static bool TryReadHexStringToken(string value, int start, out RedactionTextStringToken token) {
        int index = start + 1;
        while (index < value.Length && value[index] != '>') {
            if (!IsHexStringCharacter(value[index])) {
                token = default;
                return false;
            }

            index++;
        }

        if (index >= value.Length || value[index] != '>') {
            token = default;
            return false;
        }

        token = new RedactionTextStringToken(start, index - start + 1, isHex: true, value.Substring(start + 1, index - start - 1));
        return true;
    }

    private static bool IsHexStringCharacter(char value) {
        return char.IsWhiteSpace(value) ||
            (value >= '0' && value <= '9') ||
            (value >= 'A' && value <= 'F') ||
            (value >= 'a' && value <= 'f');
    }

    private static string? ReadLastFontName(string value) {
        string? fontName = null;
        foreach (Match match in FontSelectionRegex.Matches(value)) {
            fontName = match.Groups[1].Value;
        }

        return fontName;
    }

    private static string DecodeHexString(
        string value,
        string? currentFont,
        IReadOnlyDictionary<string, Func<byte[], string>> fontDecoders) {
        string hex = RemoveWhitespace(value);
        if (hex.Length == 0) {
            return string.Empty;
        }

        if ((hex.Length & 1) == 1) {
            hex += "0";
        }

        var bytes = new byte[hex.Length / 2];
        for (int i = 0; i < bytes.Length; i++) {
            bytes[i] = Convert.ToByte(hex.Substring(i * 2, 2), 16);
        }

        return DecodeWithCurrentFont(bytes, currentFont, fontDecoders);
    }

    private static string DecodeLiteralString(
        string value,
        string? currentFont,
        IReadOnlyDictionary<string, Func<byte[], string>> fontDecoders) {
        if (value.Length < 2) {
            return string.Empty;
        }

        string body = value.Substring(1, value.Length - 2);
        return DecodeWithCurrentFont(PdfStringParser.ParseLiteralToBytes(body), currentFont, fontDecoders);
    }

    private static string DecodeWithCurrentFont(
        byte[] bytes,
        string? currentFont,
        IReadOnlyDictionary<string, Func<byte[], string>> fontDecoders) {
        if (!string.IsNullOrEmpty(currentFont) &&
            fontDecoders.TryGetValue(currentFont!, out Func<byte[], string>? decoder)) {
            return decoder(bytes);
        }

        return PdfWinAnsiEncoding.Decode(bytes);
    }

    private static Dictionary<string, Func<byte[], string>> MergeDecoders(
        IReadOnlyDictionary<string, Func<byte[], string>> parent,
        Dictionary<string, Func<byte[], string>> local) {
        var merged = new Dictionary<string, Func<byte[], string>>(StringComparer.Ordinal);
        foreach (KeyValuePair<string, Func<byte[], string>> entry in parent) {
            merged[entry.Key] = entry.Value;
        }

        foreach (KeyValuePair<string, Func<byte[], string>> entry in local) {
            merged[entry.Key] = entry.Value;
        }

        return merged;
    }

    private static PdfDictionary? GetInheritedDictionary(Dictionary<int, PdfIndirectObject> objects, PdfDictionary dictionary, string key) {
        PdfDictionary? current = dictionary;
        int guard = 0;
        while (current is not null && guard++ < 100) {
            if (current.Items.TryGetValue(key, out PdfObject? value) &&
                PdfObjectLookup.Resolve(objects, value) is PdfDictionary resolved) {
                return resolved;
            }

            if (!current.Items.TryGetValue("Parent", out PdfObject? parentObject) ||
                parentObject is not PdfReference parentReference ||
                !PdfObjectLookup.TryGet(objects, parentReference, out PdfIndirectObject? parentIndirect) ||
                parentIndirect.Value is not PdfDictionary parentDictionary) {
                break;
            }

            current = parentDictionary;
        }

        return null;
    }

    private static Matrix2D ApplyFormMatrix(Matrix2D invocationTransform, PdfDictionary? formDict) {
        if (formDict is null ||
            !formDict.Items.TryGetValue("Matrix", out PdfObject? matrixObj) ||
            matrixObj is not PdfArray array ||
            array.Items.Count < 6) {
            return invocationTransform;
        }

        var formMatrix = new Matrix2D(
            (array.Items[0] as PdfNumber)?.Value ?? 1,
            (array.Items[1] as PdfNumber)?.Value ?? 0,
            (array.Items[2] as PdfNumber)?.Value ?? 0,
            (array.Items[3] as PdfNumber)?.Value ?? 1,
            (array.Items[4] as PdfNumber)?.Value ?? 0,
            (array.Items[5] as PdfNumber)?.Value ?? 0);

        return Matrix2D.Multiply(invocationTransform, formMatrix);
    }

    private static string WrapContentWithTransform(string content, Matrix2D transform) {
        return string.Format(
            CultureInfo.InvariantCulture,
            "q {0} {1} {2} {3} {4} {5} cm {6} Q",
            transform.A,
            transform.B,
            transform.C,
            transform.D,
            transform.E,
            transform.F,
            content);
    }

    private static bool Intersects(double ax, double ay, double aw, double ah, double bx, double by, double bw, double bh) {
        return ax < bx + bw &&
            ax + aw > bx &&
            ay < by + bh &&
            ay + ah > by;
    }

    private readonly struct RedactionTextTarget {
        public RedactionTextTarget(string text, double x, double y, double width, double height) {
            Text = text;
            X = x;
            Y = y;
            Width = width;
            Height = height;
        }

        public string Text { get; }

        public double X { get; }

        public double Y { get; }

        public double Width { get; }

        public double Height { get; }
    }

    private readonly struct RedactionTextObject {
        public RedactionTextObject(int index, string text, RedactionTextBounds? bounds) {
            Index = index;
            Text = text;
            Bounds = bounds;
        }

        public int Index { get; }

        public string Text { get; }

        public RedactionTextBounds? Bounds { get; }
    }

    private readonly struct RedactionTextStringToken {
        public RedactionTextStringToken(int index, int length, bool isHex, string value) {
            Index = index;
            Length = length;
            IsHex = isHex;
            Value = value;
        }

        public int Index { get; }

        public int Length { get; }

        public bool IsHex { get; }

        public string Value { get; }
    }

    private readonly struct RedactionTextBounds {
        public RedactionTextBounds(double left, double bottom, double right, double top) {
            Left = left;
            Bottom = bottom;
            Right = right;
            Top = top;
        }

        public double Left { get; }

        public double Bottom { get; }

        public double Right { get; }

        public double Top { get; }
    }
}

using OfficeIMO.Pdf.Filters;
using System.Globalization;
using System.Text.RegularExpressions;

namespace OfficeIMO.Pdf;

internal static partial class PdfRedactionApplier {
    private const double RedactionFallbackTextHeight = 18D;

    private static bool RemoveMatchedTextObjects(
        Dictionary<int, PdfIndirectObject> objects,
        PdfDictionary pageDictionary,
        IReadOnlyList<PdfRedactionMatch> matches,
        IReadOnlyList<PdfRedactionArea> areas,
        ref int nextObjectNumber) {
        RedactionTextTarget[] textTargets = BuildTextTargets(matches, areas);
        if (textTargets.Length == 0 ||
            !pageDictionary.Items.TryGetValue("Contents", out PdfObject? contentsObject)) {
            return false;
        }

        bool changed = false;
        Dictionary<int, int> referenceCounts = CountIndirectReferenceUsage(objects);
        Dictionary<string, Func<byte[], string>> fontDecoders = ResourceResolver.GetFontDecoders(pageDictionary, objects);
        PdfObject currentContentsObject = contentsObject;
        PdfReference[] contentReferences = EnumerateContentReferences(objects, contentsObject).ToArray();
        var contentSegments = new List<string>(contentReferences.Length);
        bool allStreamsDecoded = true;
        foreach (PdfReference reference in contentReferences) {
            if (!PdfObjectLookup.TryGet(objects, reference, out PdfIndirectObject? indirect) ||
                indirect.Value is not PdfStream stream ||
                stream.DecodingFailed) {
                allStreamsDecoded = false;
                break;
            }

            byte[] contentBytes = StreamDecoder.Decode(stream.Dictionary, stream.Data, objects);
            contentSegments.Add(PdfEncoding.Latin1GetString(contentBytes));
        }

        if (allStreamsDecoded && contentSegments.Count > 0) {
            string combinedContent = string.Concat(contentSegments);
            TextObjectSpan[] spansToRemove = FindMatchingTextObjectSpans(
                combinedContent,
                textTargets,
                fontDecoders,
                new[] { Matrix2D.Identity },
                graphicsState: null);
            int contentOffset = 0;
            for (int index = 0; index < contentReferences.Length; index++) {
                string content = contentSegments[index];
                string scrubbed = RemoveTextObjectSpans(content, contentOffset, spansToRemove);
                changed = ReplacePageContentStreamIfChanged(
                    objects,
                    pageDictionary,
                    ref currentContentsObject,
                    contentReferences[index],
                    index,
                    content,
                    scrubbed,
                    referenceCounts,
                    ref nextObjectNumber) || changed;
                contentOffset += content.Length;
            }
        } else {
            var graphicsState = new TextScrubGraphicsState();
            for (int index = 0; index < contentReferences.Length; index++) {
                PdfReference reference = contentReferences[index];
                if (!PdfObjectLookup.TryGet(objects, reference, out PdfIndirectObject? indirect) ||
                    indirect.Value is not PdfStream stream ||
                    stream.DecodingFailed) {
                    graphicsState.Reset();
                    continue;
                }

                string content = PdfEncoding.Latin1GetString(StreamDecoder.Decode(stream.Dictionary, stream.Data, objects));
                string scrubbed = ScrubTextObjects(content, textTargets, fontDecoders, new[] { Matrix2D.Identity }, graphicsState);
                changed = ReplacePageContentStreamIfChanged(
                    objects,
                    pageDictionary,
                    ref currentContentsObject,
                    reference,
                    index,
                    content,
                    scrubbed,
                    referenceCounts,
                    ref nextObjectNumber) || changed;
            }
        }

        return ScrubMatchedFormXObjects(objects, pageDictionary, currentContentsObject, textTargets, fontDecoders, referenceCounts, ref nextObjectNumber) || changed;
    }

    private static bool ReplacePageContentStreamIfChanged(
        Dictionary<int, PdfIndirectObject> objects,
        PdfDictionary pageDictionary,
        ref PdfObject currentContentsObject,
        PdfReference reference,
        int contentIndex,
        string content,
        string scrubbed,
        IReadOnlyDictionary<int, int> referenceCounts,
        ref int nextObjectNumber) {
        if (string.Equals(content, scrubbed, StringComparison.Ordinal) ||
            !PdfObjectLookup.TryGet(objects, reference, out PdfIndirectObject? indirect) ||
            indirect.Value is not PdfStream stream) {
            return false;
        }

        PdfReference targetReference = reference;
        if (IsSharedReference(referenceCounts, reference)) {
            targetReference = CloneIndirectObject(objects, reference, indirect, ref nextObjectNumber);
            ReplacePageContentReferenceAtIndex(objects, pageDictionary, currentContentsObject, contentIndex, targetReference);
            currentContentsObject = pageDictionary.Items.TryGetValue("Contents", out PdfObject? updatedContentsObject)
                ? updatedContentsObject
                : currentContentsObject;
        }

        objects[targetReference.ObjectNumber] = new PdfIndirectObject(
            targetReference.ObjectNumber,
            targetReference.Generation,
            new PdfStream(CleanStreamDictionary(stream.Dictionary), PdfEncoding.Latin1GetBytes(scrubbed)));
        return true;
    }

    private static RedactionTextTarget[] BuildTextTargets(
        IReadOnlyList<PdfRedactionMatch> matches,
        IReadOnlyList<PdfRedactionArea> areas) {
        return matches
            .Where(match => match.Kind == PdfRedactionMatchKind.TextBlock && !string.IsNullOrWhiteSpace(match.Text))
            .Select(match => new RedactionTextTarget(
                NormalizeText(match.Text!),
                match.X,
                match.Y,
                match.Width,
                match.Height))
            .Where(target => target.Text.Length > 0)
            .Concat(areas.Select(area => new RedactionTextTarget(
                string.Empty,
                area.X,
                area.Y,
                area.Width,
                area.Height)))
            .ToArray();
    }

    private static bool ScrubMatchedFormXObjects(
        Dictionary<int, PdfIndirectObject> objects,
        PdfDictionary pageDictionary,
        PdfObject contentsObject,
        RedactionTextTarget[] textTargets,
        IReadOnlyDictionary<string, Func<byte[], string>> pageFontDecoders,
        IReadOnlyDictionary<int, int> referenceCounts,
        ref int nextObjectNumber) {
        PdfDictionary? resources = GetInheritedDictionary(objects, pageDictionary, "Resources");
        if (resources is null ||
            !resources.Items.ContainsKey("XObject")) {
            return false;
        }

        PdfDictionary xObjects = PdfPageResourceHelper.EnsurePageXObjects(objects, pageDictionary, "redaction text scrubbing");
        resources = ResolveDictionary(objects, pageDictionary.Items.TryGetValue("Resources", out PdfObject? pageResources) ? pageResources : null) ?? resources;
        PdfReference[] contentReferences = EnumerateContentReferences(objects, contentsObject).ToArray();
        var contentSegments = new string?[contentReferences.Length];
        bool allStreamsDecoded = true;
        for (int index = 0; index < contentReferences.Length; index++) {
            PdfReference reference = contentReferences[index];
            if (!PdfObjectLookup.TryGet(objects, reference, out PdfIndirectObject? indirect) ||
                indirect.Value is not PdfStream stream ||
                stream.DecodingFailed) {
                allStreamsDecoded = false;
                continue;
            }

            contentSegments[index] = PdfEncoding.Latin1GetString(StreamDecoder.Decode(stream.Dictionary, stream.Data, objects));
        }

        bool changed = false;
        if (allStreamsDecoded && contentSegments.Length > 0) {
            string combinedContent = string.Concat(contentSegments);
            TextFormScrubContentResult result = ScrubFormInvocations(objects, resources, xObjects, combinedContent, textTargets, pageFontDecoders, new[] { Matrix2D.Identity }, referenceCounts, new HashSet<int>(), ref nextObjectNumber);
            if (!string.Equals(result.Content, combinedContent, StringComparison.Ordinal)) {
                PdfObject currentContentsObject = contentsObject;
                for (int index = 0; index < contentReferences.Length; index++) {
                    string replacement = index == 0 ? result.Content : string.Empty;
                    changed = ReplacePageContentStreamIfChanged(
                        objects,
                        pageDictionary,
                        ref currentContentsObject,
                        contentReferences[index],
                        index,
                        contentSegments[index]!,
                        replacement,
                        referenceCounts,
                        ref nextObjectNumber) || changed;
                }
            }

            return result.HasChanges || changed;
        }

        PdfObject fallbackContentsObject = contentsObject;
        for (int index = 0; index < contentSegments.Length; index++) {
            string? content = contentSegments[index];
            if (content is null) {
                continue;
            }

            TextFormScrubContentResult result = ScrubFormInvocations(objects, resources, xObjects, content, textTargets, pageFontDecoders, new[] { Matrix2D.Identity }, referenceCounts, new HashSet<int>(), ref nextObjectNumber);
            changed = result.HasChanges || changed;
            changed = ReplacePageContentStreamIfChanged(
                objects,
                pageDictionary,
                ref fallbackContentsObject,
                contentReferences[index],
                index,
                content,
                result.Content,
                referenceCounts,
                ref nextObjectNumber) || changed;
        }

        return changed;
    }

    private static bool TryGetFormXObject(Dictionary<int, PdfIndirectObject> objects, PdfDictionary xObjects, string name, out PdfReference reference, out PdfStream stream) {
        if (xObjects.Items.TryGetValue(name, out PdfObject? value) &&
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
        IReadOnlyList<Matrix2D> transforms,
        TextScrubGraphicsState? graphicsState = null) {
        TextObjectSpan[] spansToRemove = FindMatchingTextObjectSpans(content, targets, fontDecoders, transforms, graphicsState);
        return RemoveTextObjectSpans(content, 0, spansToRemove);
    }

    private static TextObjectSpan[] FindMatchingTextObjectSpans(
        string content,
        RedactionTextTarget[] targets,
        IReadOnlyDictionary<string, Func<byte[], string>> fontDecoders,
        IReadOnlyList<Matrix2D> transforms,
        TextScrubGraphicsState? graphicsState) {
        List<RedactionTextObject> textObjects = CollectTextObjects(content, fontDecoders, transforms, graphicsState);
        if (textObjects.Count == 0) {
            return Array.Empty<TextObjectSpan>();
        }

        var removeByIndex = new HashSet<int>();
        for (int targetIndex = 0; targetIndex < targets.Length; targetIndex++) {
            MarkMatchingTextObjects(textObjects, targets[targetIndex], removeByIndex);
        }

        if (removeByIndex.Count == 0) {
            return Array.Empty<TextObjectSpan>();
        }

        return EnumerateTextObjectSpans(content)
            .Where(span => removeByIndex.Contains(span.Index))
            .ToArray();
    }

    private static List<RedactionTextObject> CollectTextObjects(
        string content,
        IReadOnlyDictionary<string, Func<byte[], string>> fontDecoders,
        IReadOnlyList<Matrix2D> transforms,
        TextScrubGraphicsState? graphicsState) {
        var textObjects = new List<RedactionTextObject>();
        Dictionary<int, Matrix2D> localTransforms = CollectTextObjectTransforms(content, graphicsState);
        foreach (TextObjectSpan span in EnumerateTextObjectSpans(content)) {
            string shownText = NormalizeText(ExtractTextFromTextObject(span.Value, fontDecoders));
            Matrix2D localTransform = localTransforms.TryGetValue(span.Index, out Matrix2D resolved)
                ? resolved
                : Matrix2D.Identity;
            Matrix2D[] effectiveTransforms = transforms
                .Select(parent => Matrix2D.Multiply(parent, localTransform))
                .ToArray();
            textObjects.Add(BuildRedactionTextObject(span.Index, span.Value, shownText, fontDecoders, effectiveTransforms));
        }

        return textObjects;
    }

    private static Dictionary<int, Matrix2D> CollectTextObjectTransforms(string content, TextScrubGraphicsState? graphicsState) {
        var transforms = new Dictionary<int, Matrix2D>();
        TextScrubGraphicsState state = graphicsState ?? new TextScrubGraphicsState();
        Stack<Matrix2D> stack = state.Stack;
        Matrix2D current = state.Current;
        PdfContentStreamInterpreter.Interpret(
            content,
            PdfReadLimits.DefaultMaxContentOperations,
            operation => {
                switch (operation.Name) {
                    case "q":
                        stack.Push(current);
                        break;
                    case "Q":
                        current = stack.Count > 0 ? stack.Pop() : Matrix2D.Identity;
                        break;
                    case "cm" when operation.Operands.Count >= 6:
                        int start = operation.Operands.Count - 6;
                        current = Matrix2D.Multiply(current, new Matrix2D(
                            Convert.ToDouble(operation.Operands[start], CultureInfo.InvariantCulture),
                            Convert.ToDouble(operation.Operands[start + 1], CultureInfo.InvariantCulture),
                            Convert.ToDouble(operation.Operands[start + 2], CultureInfo.InvariantCulture),
                            Convert.ToDouble(operation.Operands[start + 3], CultureInfo.InvariantCulture),
                            Convert.ToDouble(operation.Operands[start + 4], CultureInfo.InvariantCulture),
                            Convert.ToDouble(operation.Operands[start + 5], CultureInfo.InvariantCulture)));
                        break;
                    case "BT":
                        transforms[operation.OperatorOffset] = current;
                        break;
                }
            });
        state.Current = current;
        return transforms;
    }

    private sealed class TextScrubGraphicsState {
        internal Matrix2D Current { get; set; } = Matrix2D.Identity;
        internal Stack<Matrix2D> Stack { get; } = new Stack<Matrix2D>();

        internal void Reset() {
            Current = Matrix2D.Identity;
            Stack.Clear();
        }
    }

    private static string RemoveTextObjectSpans(string content, int contentOffset, IReadOnlyList<TextObjectSpan> spansToRemove) {
        if (spansToRemove.Count == 0) {
            return content;
        }

        var builder = new StringBuilder(content.Length);
        int cursor = 0;
        int contentEnd = contentOffset + content.Length;
        foreach (TextObjectSpan span in spansToRemove) {
            int spanStart = Math.Max(contentOffset, span.Index);
            int spanEnd = Math.Min(contentEnd, span.Index + span.Length);
            if (spanStart >= spanEnd) {
                continue;
            }

            int localStart = spanStart - contentOffset;
            int localEnd = spanEnd - contentOffset;
            if (localEnd <= cursor) {
                continue;
            }

            int copyEnd = Math.Max(cursor, localStart);
            builder.Append(content, cursor, copyEnd - cursor);
            cursor = localEnd;
        }

        if (cursor == 0) {
            return content;
        }

        builder.Append(content, cursor, content.Length - cursor);
        return builder.ToString();
    }

    private static IEnumerable<TextObjectSpan> EnumerateTextObjectSpans(string content) {
        int start = -1;
        for (int i = 0; i < content.Length;) {
            if (TrySkipPdfStringOrComment(content, i, out int nextIndex)) {
                i = nextIndex;
                continue;
            }

            if (start < 0 && IsPdfOperatorAt(content, i, "BT")) {
                start = i;
                i += 2;
                continue;
            }

            if (start >= 0 && IsPdfOperatorAt(content, i, "ET")) {
                int end = i + 2;
                yield return new TextObjectSpan(start, end - start, content.Substring(start, end - start));
                start = -1;
                i = end;
                continue;
            }

            i++;
        }
    }

    private static bool TrySkipPdfStringOrComment(string content, int index, out int nextIndex) {
        nextIndex = index;
        if (content[index] == '%') {
            nextIndex = index + 1;
            while (nextIndex < content.Length && content[nextIndex] != '\r' && content[nextIndex] != '\n') {
                nextIndex++;
            }

            return true;
        }

        if (content[index] == '(') {
            nextIndex = SkipLiteralString(content, index);
            return true;
        }

        if (content[index] == '<' && (index + 1 >= content.Length || content[index + 1] != '<')) {
            nextIndex = SkipHexString(content, index);
            return true;
        }

        return false;
    }

    private static int SkipLiteralString(string content, int index) {
        int depth = 1;
        bool escaped = false;
        index++;
        while (index < content.Length && depth > 0) {
            char current = content[index++];
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

        return index;
    }

    private static int SkipHexString(string content, int index) {
        index++;
        while (index < content.Length && content[index] != '>') {
            index++;
        }

        return index < content.Length ? index + 1 : index;
    }

    private static bool IsPdfOperatorAt(string content, int index, string operatorName) {
        if (index + operatorName.Length > content.Length ||
            !string.Equals(content.Substring(index, operatorName.Length), operatorName, StringComparison.Ordinal)) {
            return false;
        }

        return IsPdfTokenBoundary(content, index - 1) &&
            IsPdfTokenBoundary(content, index + operatorName.Length);
    }

    private static bool IsPdfTokenBoundary(string content, int index) {
        if (index < 0 || index >= content.Length) {
            return true;
        }

        char value = content[index];
        return char.IsWhiteSpace(value) ||
            value == '(' ||
            value == ')' ||
            value == '<' ||
            value == '>' ||
            value == '[' ||
            value == ']' ||
            value == '{' ||
            value == '}' ||
            value == '/' ||
            value == '%';
    }

    private static RedactionTextObject BuildRedactionTextObject(
        int index,
        string textObject,
        string shownText,
        IReadOnlyDictionary<string, Func<byte[], string>> fontDecoders,
        Matrix2D[] transforms) {
        RedactionTextBounds? bounds = null;
        for (int transformIndex = 0; transformIndex < transforms.Length; transformIndex++) {
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
        if (target.Text.Length == 0) {
            foreach (RedactionTextObject textObject in textObjects) {
                if (IntersectsTarget(textObject, target)) {
                    removeByIndex.Add(textObject.Index);
                }
            }

            return;
        }

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

    private readonly struct TextObjectSpan {
        public TextObjectSpan(int index, int length, string value) {
            Index = index;
            Length = length;
            Value = value;
        }

        public int Index { get; }

        public int Length { get; }

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

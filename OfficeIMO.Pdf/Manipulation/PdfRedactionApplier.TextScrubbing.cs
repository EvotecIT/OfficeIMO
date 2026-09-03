using OfficeIMO.Pdf.Filters;
using System.Globalization;
using System.Text.RegularExpressions;

namespace OfficeIMO.Pdf;

internal static partial class PdfRedactionApplier {
    private const double RedactionFallbackTextHeight = 18D;

    private static bool RemoveMatchedTextObjects(
        Dictionary<int, PdfIndirectObject> objects,
        PdfDictionary pageDictionary,
        IReadOnlyList<PdfRedactionArea> areas,
        PdfReadLimits limits,
        HashSet<PdfStream> sourceStreamIdentities,
        ref int nextObjectNumber) {
        RedactionTextTarget[] textTargets = BuildTextTargets(areas);
        if (textTargets.Length == 0 ||
            !pageDictionary.Items.TryGetValue("Contents", out PdfObject? contentsObject)) {
            return false;
        }

        bool changed = false;
        Dictionary<int, int> referenceCounts = CountIndirectReferenceUsage(objects);
        Dictionary<string, Func<byte[], string>> fontDecoders = ResourceResolver.GetFontDecoders(pageDictionary, objects);
        Dictionary<string, Func<byte[], double>> fontWidthProviders = ResourceResolver.GetFontWidthProviders(pageDictionary, objects);
        PdfDictionary? pageResources = GetInheritedDictionary(objects, pageDictionary, "Resources");
        IReadOnlyDictionary<string, PdfExtGStateFontSelection> extGStateFonts = pageResources == null
            ? new Dictionary<string, PdfExtGStateFontSelection>(StringComparer.Ordinal)
            : ResolveExtGStateFontSelections(objects, pageResources);
        HashSet<string> verticalWritingFonts = GetVerticalWritingFontResources(pageResources, objects);
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

            byte[] contentBytes = StreamDecoder.DecodeRequired(stream.Dictionary, stream.Data, objects, GetMutationDecodeLimit(stream, limits, sourceStreamIdentities));
            contentSegments.Add(PdfEncoding.Latin1GetString(contentBytes));
        }

        if (allStreamsDecoded && contentSegments.Count > 0) {
            string combinedContent = string.Concat(contentSegments);
            TextObjectRewrite[] rewrites = FindMatchingTextObjectRewrites(
                combinedContent,
                textTargets,
                fontDecoders,
                fontWidthProviders,
                new[] { Matrix2D.Identity },
                graphicsState: null,
                limits,
                verticalWritingFonts,
                extGStateFonts: extGStateFonts);
            int contentOffset = 0;
            for (int index = 0; index < contentReferences.Length; index++) {
                string content = contentSegments[index];
                string scrubbed = RewriteTextObjectSpans(content, contentOffset, rewrites);
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

                string content = PdfEncoding.Latin1GetString(StreamDecoder.DecodeRequired(stream.Dictionary, stream.Data, objects, GetMutationDecodeLimit(stream, limits, sourceStreamIdentities)));
                string scrubbed = ScrubTextObjects(content, textTargets, fontDecoders, fontWidthProviders, new[] { Matrix2D.Identity }, limits, graphicsState, extGStateFonts, verticalWritingFonts);
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

        return ScrubMatchedFormXObjects(objects, pageDictionary, currentContentsObject, textTargets, fontDecoders, fontWidthProviders, verticalWritingFonts, referenceCounts, limits, sourceStreamIdentities, ref nextObjectNumber) || changed;
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

    private static RedactionTextTarget[] BuildTextTargets(IReadOnlyList<PdfRedactionArea> areas) {
        return areas
            .Select(area => new RedactionTextTarget(
                string.Empty,
                area.X,
                area.Y,
                area.Width,
                area.Height))
            .ToArray();
    }

    private static bool ScrubMatchedFormXObjects(
        Dictionary<int, PdfIndirectObject> objects,
        PdfDictionary pageDictionary,
        PdfObject contentsObject,
        RedactionTextTarget[] textTargets,
        IReadOnlyDictionary<string, Func<byte[], string>> pageFontDecoders,
        IReadOnlyDictionary<string, Func<byte[], double>> pageFontWidthProviders,
        ISet<string> pageVerticalWritingFonts,
        IReadOnlyDictionary<int, int> referenceCounts,
        PdfReadLimits limits,
        HashSet<PdfStream> sourceStreamIdentities,
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

            contentSegments[index] = PdfEncoding.Latin1GetString(StreamDecoder.DecodeRequired(stream.Dictionary, stream.Data, objects, GetMutationDecodeLimit(stream, limits, sourceStreamIdentities)));
        }

        bool changed = false;
        if (allStreamsDecoded && contentSegments.Length > 0) {
            string combinedContent = string.Concat(contentSegments);
            TextFormScrubContentResult result = ScrubFormInvocations(objects, resources, xObjects, combinedContent, textTargets, pageFontDecoders, pageFontWidthProviders, pageVerticalWritingFonts, new[] { Matrix2D.Identity }, PdfTextStateSnapshot.Default, referenceCounts, new HashSet<int>(), limits, sourceStreamIdentities, ref nextObjectNumber);
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

            TextFormScrubContentResult result = ScrubFormInvocations(objects, resources, xObjects, content, textTargets, pageFontDecoders, pageFontWidthProviders, pageVerticalWritingFonts, new[] { Matrix2D.Identity }, PdfTextStateSnapshot.Default, referenceCounts, new HashSet<int>(), limits, sourceStreamIdentities, ref nextObjectNumber);
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
        IReadOnlyDictionary<string, Func<byte[], double>> fontWidthProviders,
        IReadOnlyList<Matrix2D> transforms,
        PdfReadLimits limits,
        TextScrubGraphicsState? graphicsState = null,
        IReadOnlyDictionary<string, PdfExtGStateFontSelection>? extGStateFonts = null,
        ISet<string>? verticalWritingFonts = null) {
        TextObjectRewrite[] rewrites = FindMatchingTextObjectRewrites(content, targets, fontDecoders, fontWidthProviders, transforms, graphicsState, limits, verticalWritingFonts, extGStateFonts);
        return RewriteTextObjectSpans(content, 0, rewrites);
    }

    private static TextObjectRewrite[] FindMatchingTextObjectRewrites(
        string content,
        RedactionTextTarget[] targets,
        IReadOnlyDictionary<string, Func<byte[], string>> fontDecoders,
        IReadOnlyDictionary<string, Func<byte[], double>> fontWidthProviders,
        IReadOnlyList<Matrix2D> transforms,
        TextScrubGraphicsState? graphicsState,
        PdfReadLimits limits,
        ISet<string>? verticalWritingFonts = null,
        IReadOnlyDictionary<string, PdfExtGStateFontSelection>? extGStateFonts = null) {
        List<RedactionTextObject> textObjects = CollectTextObjects(content, fontDecoders, fontWidthProviders, transforms, graphicsState, limits, extGStateFonts);
        if (textObjects.Count == 0) {
            return Array.Empty<TextObjectRewrite>();
        }

        var targetsByIndex = new Dictionary<int, List<RedactionTextTarget>>();
        for (int targetIndex = 0; targetIndex < targets.Length; targetIndex++) {
            MarkMatchingTextObjects(textObjects, targets[targetIndex], targetsByIndex);
        }

        if (targetsByIndex.Count == 0) {
            return Array.Empty<TextObjectRewrite>();
        }

        return EnumerateTextObjectSpans(content, limits)
            .Where(span => targetsByIndex.ContainsKey(span.Index))
            .Select(span => {
                List<RedactionTextTarget> selectedTargets = targetsByIndex[span.Index];
                RedactionTextObject textObject = textObjects.First(value => value.Index == span.Index);
                var rewriteTargets = selectedTargets
                    .Select(static target => new PdfContentStreamTextRewriteTarget(
                        target.X,
                        target.Y,
                        target.Width,
                        target.Height <= 0D ? RedactionFallbackTextHeight : target.Height))
                    .ToArray();
                string replacement = PdfContentStreamTextRewriter.TryRemoveIntersectingGlyphs(
                    span.Value,
                    fontDecoders,
                    fontWidthProviders,
                    textObject.Transforms,
                    textObject.TextState,
                    rewriteTargets,
                    limits,
                    verticalWritingFonts,
                    extGStateFonts,
                    out string rewritten)
                    ? rewritten
                    : string.Empty;
                return new TextObjectRewrite(span.Index, span.Length, replacement);
            })
            .ToArray();
    }

    private static List<RedactionTextObject> CollectTextObjects(
        string content,
        IReadOnlyDictionary<string, Func<byte[], string>> fontDecoders,
        IReadOnlyDictionary<string, Func<byte[], double>> fontWidthProviders,
        IReadOnlyList<Matrix2D> transforms,
        TextScrubGraphicsState? graphicsState,
        PdfReadLimits limits,
        IReadOnlyDictionary<string, PdfExtGStateFontSelection>? extGStateFonts = null) {
        var textObjects = new List<RedactionTextObject>();
        Dictionary<int, TextObjectContext> contexts = CollectTextObjectContexts(content, graphicsState, limits, extGStateFonts);
        foreach (TextObjectSpan span in EnumerateTextObjectSpans(content, limits)) {
            string shownText = NormalizeText(ExtractTextFromTextObject(span.Value, fontDecoders));
            TextObjectContext context = contexts.TryGetValue(span.Index, out TextObjectContext resolved)
                ? resolved
                : new TextObjectContext(Matrix2D.Identity, PdfTextStateSnapshot.Default);
            Matrix2D[] effectiveTransforms = transforms
                .Select(parent => Matrix2D.Multiply(parent, context.Transform))
                .ToArray();
            textObjects.Add(BuildRedactionTextObject(span.Index, span.Value, shownText, fontDecoders, fontWidthProviders, effectiveTransforms, context.TextState));
        }

        return textObjects;
    }

    private static string RewriteTextObjectSpans(string content, int contentOffset, IReadOnlyList<TextObjectRewrite> rewrites) {
        if (rewrites.Count == 0) {
            return content;
        }

        var builder = new StringBuilder(content.Length);
        int cursor = 0;
        int contentEnd = contentOffset + content.Length;
        foreach (TextObjectRewrite rewrite in rewrites) {
            int spanStart = Math.Max(contentOffset, rewrite.Index);
            int spanEnd = Math.Min(contentEnd, rewrite.Index + rewrite.Length);
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
            if (rewrite.Index >= contentOffset && rewrite.Index < contentEnd) {
                builder.Append(rewrite.Replacement);
            }
            cursor = localEnd;
        }

        if (cursor == 0) {
            return content;
        }

        builder.Append(content, cursor, content.Length - cursor);
        return builder.ToString();
    }

    private static List<TextObjectSpan> EnumerateTextObjectSpans(string content, PdfReadLimits limits) {
        var spans = new List<TextObjectSpan>();
        int start = -1;
        PdfContentStreamInterpreter.Interpret(
            content,
            limits.MaxContentOperations,
            operation => {
                if (start < 0 && string.Equals(operation.Name, "BT", StringComparison.Ordinal)) {
                    start = operation.OperatorOffset;
                    return;
                }
                if (start >= 0 && string.Equals(operation.Name, "ET", StringComparison.Ordinal)) {
                    int end = operation.OperatorOffset + 2;
                    spans.Add(new TextObjectSpan(start, end - start, content.Substring(start, end - start)));
                    start = -1;
                }
            },
            maxNestingDepth: limits.MaxContentNestingDepth,
            maxOperands: limits.MaxContentOperands);
        return spans;
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
        IReadOnlyDictionary<string, Func<byte[], double>> fontWidthProviders,
        Matrix2D[] transforms,
        PdfTextStateSnapshot textState) {
        RedactionTextBounds? bounds = null;
        for (int transformIndex = 0; transformIndex < transforms.Length; transformIndex++) {
            string transformedContent = WrapContentWithTransform(textObject, transforms[transformIndex]);
            List<PdfTextSpan> spans = ParseTextSpans(transformedContent, fontDecoders, fontWidthProviders, textState);
            for (int spanIndex = 0; spanIndex < spans.Count; spanIndex++) {
                bounds = AddSpanBounds(bounds, spans[spanIndex]);
            }
        }

        return new RedactionTextObject(index, shownText, bounds, transforms, textState);
    }

    private static List<PdfTextSpan> ParseTextSpans(
        string content,
        IReadOnlyDictionary<string, Func<byte[], string>> fontDecoders,
        IReadOnlyDictionary<string, Func<byte[], double>> fontWidthProviders,
        PdfTextStateSnapshot? initialTextState = null) {
        string DecodeWithFont(string fontResource, byte[] bytes) =>
            fontDecoders.TryGetValue(fontResource, out Func<byte[], string>? decoder)
                ? decoder(bytes)
                : PdfWinAnsiEncoding.Decode(bytes);
        double SumWidth1000(string fontResource, byte[] bytes) =>
            fontWidthProviders.TryGetValue(fontResource, out Func<byte[], double>? provider)
                ? provider(bytes)
                : bytes is null ? 0D : bytes.Length * 500D;

        return TextContentParser.Parse(content, DecodeWithFont, SumWidth1000, initialTextState: initialTextState);
    }

    private static RedactionTextBounds AddSpanBounds(RedactionTextBounds? current, PdfTextSpan span) {
        PdfTextSpanBounds spanBounds = PdfTextSpanGeometry.GetAxisAlignedBounds(span);
        double left = spanBounds.Left;
        double right = spanBounds.Right;
        double bottom = spanBounds.Bottom;
        double top = spanBounds.Top;
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
        Dictionary<int, List<RedactionTextTarget>> targetsByIndex) {
        if (target.Text.Length == 0) {
            foreach (RedactionTextObject textObject in textObjects) {
                if (IntersectsTarget(textObject, target)) {
                    AddRewriteTarget(targetsByIndex, textObject.Index, target);
                }
            }

            return;
        }

        for (int start = 0; start < textObjects.Count; start++) {
            if (ContainsOrdinal(textObjects[start].Text, target.Text)) {
                if (IntersectsTarget(textObjects[start], target)) {
                    AddRewriteTarget(targetsByIndex, textObjects[start].Index, target);
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
                    AddRewriteTarget(targetsByIndex, textObjects[remove].Index, target);
                }

                break;
            }
        }
    }

    private static void AddRewriteTarget(
        Dictionary<int, List<RedactionTextTarget>> targetsByIndex,
        int textObjectIndex,
        RedactionTextTarget target) {
        if (!targetsByIndex.TryGetValue(textObjectIndex, out List<RedactionTextTarget>? targets)) {
            targets = new List<RedactionTextTarget>();
            targetsByIndex[textObjectIndex] = targets;
        }
        targets.Add(target);
    }

    private static bool IntersectsTarget(RedactionTextObject textObject, RedactionTextTarget target) =>
        IntersectsTarget(textObject.Bounds, target);

    private static bool IntersectsTarget(RedactionTextBounds? bounds, RedactionTextTarget target) {
        if (bounds is null) {
            // An area redaction is a confidentiality boundary. If a text object cannot be
            // located, retaining it could leave extractable text inside the painted area.
            // Text-targeted redactions still require their textual match before reaching here.
            return target.Text.Length == 0;
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

    private static Dictionary<string, Func<byte[], double>> MergeWidthProviders(
        IReadOnlyDictionary<string, Func<byte[], double>> parent,
        Dictionary<string, Func<byte[], double>> local) {
        var merged = new Dictionary<string, Func<byte[], double>>(StringComparer.Ordinal);
        foreach (KeyValuePair<string, Func<byte[], double>> entry in parent) {
            merged[entry.Key] = entry.Value;
        }

        foreach (KeyValuePair<string, Func<byte[], double>> entry in local) {
            merged[entry.Key] = entry.Value;
        }

        return merged;
    }

    private static HashSet<string> GetVerticalWritingFontResources(
        PdfDictionary? resources,
        Dictionary<int, PdfIndirectObject> objects) =>
        resources == null
            ? new HashSet<string>(StringComparer.Ordinal)
            : ResourceResolver.GetFontsForResources(resources, objects)
                .Where(static entry => string.Equals(entry.Value.Encoding, "Identity-V", StringComparison.Ordinal))
                .Select(static entry => entry.Key)
                .ToHashSet(StringComparer.Ordinal);

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
        public RedactionTextObject(int index, string text, RedactionTextBounds? bounds, Matrix2D[] transforms, PdfTextStateSnapshot textState) {
            Index = index;
            Text = text;
            Bounds = bounds;
            Transforms = transforms;
            TextState = textState;
        }

        public int Index { get; }

        public string Text { get; }

        public RedactionTextBounds? Bounds { get; }

        public Matrix2D[] Transforms { get; }

        public PdfTextStateSnapshot TextState { get; }
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

    private readonly struct TextObjectRewrite {
        public TextObjectRewrite(int index, int length, string replacement) {
            Index = index;
            Length = length;
            Replacement = replacement;
        }

        public int Index { get; }

        public int Length { get; }

        public string Replacement { get; }
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

using System.Globalization;
using System.Text;

namespace OfficeIMO.Pdf;

/// <summary>
/// Rewrites text-show operators while retaining the original encoded glyphs, font resources,
/// text state, and text-object structure. Unsupported mappings fail closed so callers can
/// remove the complete text object instead of retaining uncertain content.
/// </summary>
internal static class PdfContentStreamTextRewriter {
    internal static bool TryRemoveIntersectingGlyphs(
        string textObject,
        IReadOnlyDictionary<string, Func<byte[], string>> fontDecoders,
        IReadOnlyDictionary<string, Func<byte[], double>> fontWidthProviders,
        IReadOnlyList<Matrix2D> transforms,
        PdfTextStateSnapshot initialTextState,
        IReadOnlyList<PdfContentStreamTextRewriteTarget> targets,
        PdfReadLimits limits,
        ISet<string>? verticalWritingFonts,
        IReadOnlyDictionary<string, PdfExtGStateFontSelection>? extGStateFonts,
        out string rewritten) {
        rewritten = textObject;
        if (targets.Count == 0 || transforms.Count == 0) return false;

        List<Dictionary<int, List<PdfTextSpan>>> spansByTransform = ParseSpansByOperator(
            textObject,
            fontDecoders,
            fontWidthProviders,
            transforms,
            initialTextState,
            extGStateFonts);
        if (spansByTransform.Count != transforms.Count) return false;

        PdfTextStateSnapshot currentTextState = initialTextState;
        var textStateStack = new Stack<PdfTextStateSnapshot>();
        int previousOperatorEnd = 0;
        bool sawTextShowOperator = false;
        bool removedAnyGlyph = false;
        bool safe = true;
        var edits = new List<TextShowEdit>();

        PdfContentStreamInterpreter.InterpretUntil(
            textObject,
            limits.MaxContentOperations,
            operation => {
                int operationStart = FindOperationStart(textObject, previousOperatorEnd, operation.OperatorOffset);
                previousOperatorEnd = Math.Min(textObject.Length, operation.OperatorOffset + operation.Name.Length);

                switch (operation.Name) {
                    case "Tf" when operation.Operands.Count >= 2:
                        currentTextState = currentTextState.WithFont(
                            operation.Operands[operation.Operands.Count - 2] as string ?? currentTextState.FontResource,
                            NumberAt(operation.Operands, operation.Operands.Count - 1, currentTextState.FontSize));
                        break;
                    case "Tc" when operation.Operands.Count >= 1:
                        currentTextState = currentTextState.WithCharacterSpacing(NumberAt(operation.Operands, operation.Operands.Count - 1, currentTextState.CharacterSpacing));
                        break;
                    case "Tw" when operation.Operands.Count >= 1:
                        currentTextState = currentTextState.WithWordSpacing(NumberAt(operation.Operands, operation.Operands.Count - 1, currentTextState.WordSpacing));
                        break;
                    case "Tz" when operation.Operands.Count >= 1:
                        currentTextState = currentTextState.WithHorizontalScaling(NumberAt(operation.Operands, operation.Operands.Count - 1, currentTextState.HorizontalScaling * 100D) / 100D);
                        break;
                    case "TL" when operation.Operands.Count >= 1:
                        currentTextState = currentTextState.WithLeading(NumberAt(operation.Operands, operation.Operands.Count - 1, currentTextState.Leading));
                        break;
                    case "TD" when operation.Operands.Count >= 2:
                        currentTextState = currentTextState.WithLeading(-NumberAt(operation.Operands, operation.Operands.Count - 1, -currentTextState.Leading));
                        break;
                    case "Ts" when operation.Operands.Count >= 1:
                        currentTextState = currentTextState.WithTextRise(NumberAt(operation.Operands, operation.Operands.Count - 1, currentTextState.TextRise));
                        break;
                    case "Tr" when operation.Operands.Count >= 1:
                        currentTextState = currentTextState.WithTextRenderingMode((int)NumberAt(operation.Operands, operation.Operands.Count - 1, currentTextState.TextRenderingMode));
                        break;
                    case "gs" when operation.Operands.Count >= 1:
                        if (operation.Operands[operation.Operands.Count - 1] is string graphicsStateName &&
                            extGStateFonts != null &&
                            extGStateFonts.TryGetValue(graphicsStateName, out PdfExtGStateFontSelection fontSelection)) {
                            if (!fontSelection.IsValid) {
                                safe = false;
                                return false;
                            }
                            currentTextState = currentTextState.WithFont(fontSelection.FontResource, fontSelection.FontSize);
                        }
                        break;
                    case "q":
                        textStateStack.Push(currentTextState);
                        break;
                    case "Q" when textStateStack.Count > 0:
                        currentTextState = textStateStack.Pop();
                        break;
                }

                if (!IsTextShowOperator(operation.Name)) return true;
                sawTextShowOperator = true;

                double effectiveCharacterSpacing = currentTextState.CharacterSpacing;
                double effectiveWordSpacing = currentTextState.WordSpacing;
                if (operation.Name == "\"") {
                    if (operation.Operands.Count < 3) {
                        safe = false;
                        return false;
                    }
                    effectiveWordSpacing = NumberAt(operation.Operands, operation.Operands.Count - 3, currentTextState.WordSpacing);
                    effectiveCharacterSpacing = NumberAt(operation.Operands, operation.Operands.Count - 2, currentTextState.CharacterSpacing);
                    currentTextState = currentTextState
                        .WithWordSpacing(effectiveWordSpacing)
                        .WithCharacterSpacing(effectiveCharacterSpacing);
                }

                if (!targets.Any(target => target.MatchesRenderingMode(currentTextState.TextRenderingMode))) return true;
                if (verticalWritingFonts != null && verticalWritingFonts.Contains(currentTextState.FontResource)) {
                    safe = false;
                    return false;
                }

                if (!TryRewriteShowOperation(
                        operation,
                        currentTextState.FontResource,
                        currentTextState.FontSize,
                        effectiveCharacterSpacing,
                        effectiveWordSpacing,
                        currentTextState.HorizontalScaling,
                        fontDecoders,
                        fontWidthProviders,
                        spansByTransform,
                        targets,
                        out string replacement,
                        out bool changed)) {
                    safe = false;
                    return false;
                }

                if (changed) {
                    removedAnyGlyph = true;
                    edits.Add(new TextShowEdit(operationStart, previousOperatorEnd - operationStart, replacement));
                }
                return true;
            },
            maxNestingDepth: limits.MaxContentNestingDepth,
            maxOperands: limits.MaxContentOperands,
            dispatchInvalidOperations: true);

        if (!safe || !sawTextShowOperator) return false;
        rewritten = removedAnyGlyph ? ApplyEdits(textObject, edits) : textObject;
        return true;
    }

    private static List<Dictionary<int, List<PdfTextSpan>>> ParseSpansByOperator(
        string textObject,
        IReadOnlyDictionary<string, Func<byte[], string>> fontDecoders,
        IReadOnlyDictionary<string, Func<byte[], double>> fontWidthProviders,
        IReadOnlyList<Matrix2D> transforms,
        PdfTextStateSnapshot initialTextState,
        IReadOnlyDictionary<string, PdfExtGStateFontSelection>? extGStateFonts) {
        var result = new List<Dictionary<int, List<PdfTextSpan>>>(transforms.Count);
        IReadOnlyDictionary<string, PdfPageGraphicsStateResource>? graphicsStates = extGStateFonts?.ToDictionary(
            static entry => entry.Key,
            static entry => entry.Value.IsValid ? new PdfPageGraphicsStateResource(
                null, null, null, null, null, null,
                fontResource: entry.Value.FontResource,
                fontSize: entry.Value.FontSize) : new PdfPageGraphicsStateResource(
                    null, null, null, null, null, null,
                    hasUnsupportedTextRestampEffect: true),
            StringComparer.Ordinal);
        for (int transformIndex = 0; transformIndex < transforms.Count; transformIndex++) {
            string prefix = BuildTransformPrefix(transforms[transformIndex]);
            string wrapped = prefix + textObject + " Q";
            List<PdfTextSpan> spans = TextContentParser.Parse(
                wrapped,
                (font, bytes) => fontDecoders.TryGetValue(font, out Func<byte[], string>? decoder)
                    ? decoder(bytes)
                    : PdfWinAnsiEncoding.Decode(bytes),
                (font, bytes) => fontWidthProviders.TryGetValue(font, out Func<byte[], double>? provider)
                    ? provider(bytes)
                    : bytes.Length * 500D,
                initialTextState: initialTextState,
                graphicsStates: graphicsStates);
            var byOperator = new Dictionary<int, List<PdfTextSpan>>();
            for (int spanIndex = 0; spanIndex < spans.Count; spanIndex++) {
                int operatorOffset = checked((int)Math.Round(spans[spanIndex].PaintOrder, MidpointRounding.AwayFromZero)) - prefix.Length;
                if (!byOperator.TryGetValue(operatorOffset, out List<PdfTextSpan>? operatorSpans)) {
                    operatorSpans = new List<PdfTextSpan>();
                    byOperator[operatorOffset] = operatorSpans;
                }
                operatorSpans.Add(spans[spanIndex]);
            }
            result.Add(byOperator);
        }
        return result;
    }

    private static bool TryRewriteShowOperation(
        PdfContentOperation operation,
        string font,
        double fontSize,
        double characterSpacing,
        double wordSpacing,
        double horizontalScaling,
        IReadOnlyDictionary<string, Func<byte[], string>> fontDecoders,
        IReadOnlyDictionary<string, Func<byte[], double>> fontWidthProviders,
        List<Dictionary<int, List<PdfTextSpan>>> spansByTransform,
        IReadOnlyList<PdfContentStreamTextRewriteTarget> targets,
        out string replacement,
        out bool changed) {
        replacement = string.Empty;
        changed = false;
        if (operation.HasInvalidOperands || Math.Abs(fontSize) <= 0.000001D) return false;

        if (!TryGetTextItems(operation, out List<object> items, out int byteStringCount)) return false;
        if (byteStringCount == 0) return true;
        var spansForOperation = new List<IReadOnlyList<PdfTextSpan>>(spansByTransform.Count);
        for (int transformIndex = 0; transformIndex < spansByTransform.Count; transformIndex++) {
            if (!spansByTransform[transformIndex].TryGetValue(operation.OperatorOffset, out List<PdfTextSpan>? spans) ||
                spans.Count != byteStringCount) {
                return false;
            }
            spansForOperation.Add(spans);
        }

        var output = new List<string>();
        int textItemIndex = 0;
        for (int itemIndex = 0; itemIndex < items.Count; itemIndex++) {
            object item = items[itemIndex];
            if (item is double adjustment) {
                output.Add(FormatNumber(adjustment));
                continue;
            }
            if (item is not byte[] bytes) return false;

            if (bytes.Length == 0) {
                output.Add("<>");
                continue;
            }

            var transformSpans = new PdfTextSpan[spansForOperation.Count];
            for (int transformIndex = 0; transformIndex < spansForOperation.Count; transformIndex++) {
                transformSpans[transformIndex] = spansForOperation[transformIndex][textItemIndex];
            }
            textItemIndex++;

            if (!IntersectsAnyTarget(transformSpans, targets)) {
                FlushKeptBytes(output, bytes.ToList());
                continue;
            }

            if (transformSpans.Any(static span => span.HasActualText)) return false;

            if (!TryRewriteByteString(
                    bytes,
                    font,
                    fontSize,
                    characterSpacing,
                    wordSpacing,
                    horizontalScaling,
                    fontDecoders,
                    fontWidthProviders,
                    transformSpans,
                    targets,
                    output,
                    out bool stringChanged)) {
                return false;
            }
            changed |= stringChanged;
        }

        if (!changed) return true;
        string array = "[" + string.Join(" ", output) + "] TJ";
        if (operation.Name == "'") {
            replacement = "T* " + array;
        } else if (operation.Name == "\"") {
            double nextWordSpacing = NumberAt(operation.Operands, operation.Operands.Count - 3, wordSpacing);
            double nextCharacterSpacing = NumberAt(operation.Operands, operation.Operands.Count - 2, characterSpacing);
            replacement = FormatNumber(nextWordSpacing) + " Tw " +
                FormatNumber(nextCharacterSpacing) + " Tc T* " + array;
        } else {
            replacement = array;
        }
        return true;
    }

    private static bool IntersectsAnyTarget(
        PdfTextSpan[] spans,
        IReadOnlyList<PdfContentStreamTextRewriteTarget> targets) {
        for (int spanIndex = 0; spanIndex < spans.Length; spanIndex++) {
            PdfTextSpanBounds bounds = PdfTextSpanGeometry.GetAxisAlignedBounds(spans[spanIndex]);
            for (int targetIndex = 0; targetIndex < targets.Count; targetIndex++) {
                if (targets[targetIndex].Intersects(bounds, spans[spanIndex].TextRenderingMode)) return true;
            }
        }
        return false;
    }

    private static bool TryRewriteByteString(
        byte[] bytes,
        string font,
        double fontSize,
        double characterSpacing,
        double wordSpacing,
        double horizontalScaling,
        IReadOnlyDictionary<string, Func<byte[], string>> fontDecoders,
        IReadOnlyDictionary<string, Func<byte[], double>> fontWidthProviders,
        PdfTextSpan[] spans,
        IReadOnlyList<PdfContentStreamTextRewriteTarget> targets,
        List<string> output,
        out bool changed) {
        changed = false;
        if (bytes.Length == 0) {
            output.Add("<>");
            return true;
        }

        Func<byte[], string> decoder = fontDecoders.TryGetValue(font, out Func<byte[], string>? resolvedDecoder)
            ? resolvedDecoder
            : PdfWinAnsiEncoding.Decode;
        Func<byte[], double> widthProvider = fontWidthProviders.TryGetValue(font, out Func<byte[], double>? resolvedWidthProvider)
            ? resolvedWidthProvider
            : value => value.Length * 500D;
        if (!TrySplitGlyphs(bytes, decoder, widthProvider, out List<EncodedGlyph> glyphs)) return false;

        string decoded = string.Concat(glyphs.Select(static glyph => glyph.Text));
        var characterBoundariesByTransform = new double[spans.Length][];
        for (int transformIndex = 0; transformIndex < spans.Length; transformIndex++) {
            PdfTextSpan span = spans[transformIndex];
            if (!string.Equals(span.RestampText, decoded, StringComparison.Ordinal) ||
                !TryResolveCharacterAdvances(span, glyphs, fontSize, characterSpacing, wordSpacing, out IReadOnlyList<double> characterAdvances) ||
                !TryResolveCharacterBoundaries(characterAdvances, span.CharacterAdvanceDirection, out double[] boundaries)) return false;
            characterBoundariesByTransform[transformIndex] = boundaries;
        }

        int characterOffset = 0;
        var kept = new List<byte>();
        double removedAdvance1000 = 0D;
        for (int glyphIndex = 0; glyphIndex < glyphs.Count; glyphIndex++) {
            EncodedGlyph glyph = glyphs[glyphIndex];
            bool remove = false;
            for (int transformIndex = 0; transformIndex < spans.Length && !remove; transformIndex++) {
                PdfTextSpan span = spans[transformIndex];
                double[] boundaries = characterBoundariesByTransform[transformIndex];
                double start = boundaries[characterOffset];
                double end = boundaries[characterOffset + glyph.Text.Length];
                double offset = start;
                double baselineScale = span.TextToPageTransform.HasValue
                    ? Math.Sqrt(
                        span.TextToPageTransform.Value.A * span.TextToPageTransform.Value.A +
                        span.TextToPageTransform.Value.B * span.TextToPageTransform.Value.B)
                    : 1D;
                double advance = Math.Abs(glyph.Width1000 / 1000D * fontSize * horizontalScaling * baselineScale);
                PdfTextSpanBounds bounds = PdfTextSpanGeometry.GetAxisAlignedBounds(span, offset, advance);
                for (int targetIndex = 0; targetIndex < targets.Count; targetIndex++) {
                    if (targets[targetIndex].Intersects(bounds, span.TextRenderingMode)) {
                        remove = true;
                        break;
                    }
                }
            }

            if (remove) {
                FlushKeptBytes(output, kept);
                double spacing = characterSpacing + (IsWordSpacingCode(glyph.Bytes) ? wordSpacing : 0D);
                removedAdvance1000 += glyph.Width1000 + (spacing * 1000D / fontSize);
                changed = true;
            } else {
                FlushRemovedAdvance(output, ref removedAdvance1000);
                kept.AddRange(glyph.Bytes);
            }
            characterOffset += glyph.Text.Length;
        }
        FlushKeptBytes(output, kept);
        FlushRemovedAdvance(output, ref removedAdvance1000);
        return true;
    }

    private static bool TryResolveCharacterBoundaries(
        IReadOnlyList<double> advances,
        double characterAdvanceDirection,
        out double[] boundaries) {
        double signedTotal = 0D;
        for (int index = 0; index < advances.Count; index++) {
            double value = advances[index];
            if (double.IsNaN(value) || double.IsInfinity(value)) {
                boundaries = Array.Empty<double>();
                return false;
            }
            signedTotal += value;
        }
        if (Math.Abs(signedTotal) <= 0.000001D || double.IsNaN(signedTotal) || double.IsInfinity(signedTotal)) {
            boundaries = Array.Empty<double>();
            return false;
        }

        double directionSign = characterAdvanceDirection != 0D
            ? characterAdvanceDirection
            : signedTotal < 0D ? -1D : 1D;
        boundaries = new double[advances.Count + 1];
        for (int index = 0; index < advances.Count; index++) {
            boundaries[index + 1] = boundaries[index] + advances[index] * directionSign;
            if (double.IsNaN(boundaries[index + 1]) || double.IsInfinity(boundaries[index + 1])) {
                boundaries = Array.Empty<double>();
                return false;
            }
        }
        return true;
    }

    private static bool TryResolveCharacterAdvances(
        PdfTextSpan span,
        IReadOnlyList<EncodedGlyph> glyphs,
        double fontSize,
        double characterSpacing,
        double wordSpacing,
        out IReadOnlyList<double> characterAdvances) {
        if (span.CharacterAdvances is not null && span.CharacterAdvances.Count == span.RestampText.Length) {
            characterAdvances = span.CharacterAdvances;
            return true;
        }

        var unscaled = new List<double>(span.RestampText.Length);
        double unscaledTotal = 0D;
        for (int glyphIndex = 0; glyphIndex < glyphs.Count; glyphIndex++) {
            EncodedGlyph glyph = glyphs[glyphIndex];
            double spacing = characterSpacing + (IsWordSpacingCode(glyph.Bytes) ? wordSpacing : 0D);
            double glyphAdvance = glyph.Width1000 + (spacing * 1000D / fontSize);
            if (double.IsNaN(glyphAdvance) || double.IsInfinity(glyphAdvance)) {
                characterAdvances = Array.Empty<double>();
                return false;
            }
            double perCharacterAdvance = glyphAdvance / glyph.Text.Length;
            for (int characterIndex = 0; characterIndex < glyph.Text.Length; characterIndex++) unscaled.Add(perCharacterAdvance);
            unscaledTotal += glyphAdvance;
        }

        if (unscaled.Count != span.RestampText.Length || Math.Abs(unscaledTotal) <= 0.000001D) {
            characterAdvances = Array.Empty<double>();
            return false;
        }
        double scale = span.Advance / unscaledTotal;
        characterAdvances = unscaled.Select(value => value * scale).ToArray();
        return true;
    }

    private static bool TrySplitGlyphs(
        byte[] bytes,
        Func<byte[], string> decoder,
        Func<byte[], double> widthProvider,
        out List<EncodedGlyph> glyphs) {
        glyphs = new List<EncodedGlyph>();
        bool twoByte = false;
        if (bytes.Length >= 2) {
            byte[] first = { bytes[0] };
            byte[] second = { bytes[1] };
            byte[] pair = { bytes[0], bytes[1] };
            string one = decoder(first) ?? string.Empty;
            string two = decoder(pair) ?? string.Empty;
            twoByte = (string.IsNullOrEmpty(one.Trim('\0')) && !string.IsNullOrEmpty(two.Trim('\0'))) ||
                (widthProvider(first) <= 0D && widthProvider(second) <= 0D && widthProvider(pair) > 0D);
        }

        for (int index = 0; index < bytes.Length;) {
            int length = twoByte && index + 1 < bytes.Length ? 2 : 1;
            byte[] glyphBytes = new byte[length];
            Array.Copy(bytes, index, glyphBytes, 0, length);
            string text = (decoder(glyphBytes) ?? string.Empty).Replace("\0", string.Empty);
            if (text.Length == 0) return false;
            double width = widthProvider(glyphBytes);
            if (double.IsNaN(width) || double.IsInfinity(width)) return false;
            glyphs.Add(new EncodedGlyph(glyphBytes, text, width));
            index += length;
        }
        return glyphs.Count > 0;
    }

    private static bool IsWordSpacingCode(byte[] glyphBytes) =>
        glyphBytes.Length == 1 && glyphBytes[0] == 0x20;

    private static bool TryGetTextItems(PdfContentOperation operation, out List<object> items, out int byteStringCount) {
        items = new List<object>();
        byteStringCount = 0;
        object? operand;
        if (operation.Name == "TJ") {
            operand = operation.Operands.Count > 0 ? operation.Operands[operation.Operands.Count - 1] : null;
            if (operand is List<object> array) items.AddRange(array);
            else if (operand is double[] numericItems) items.AddRange(numericItems.Cast<object>());
            else return false;
        } else {
            operand = operation.Operands.Count > 0 ? operation.Operands[operation.Operands.Count - 1] : null;
            if (operand is not byte[] bytes) return false;
            items.Add(bytes);
        }
        for (int index = 0; index < items.Count; index++) {
            if (items[index] is byte[] bytes) {
                if (bytes.Length > 0) byteStringCount++;
                continue;
            }
            if (items[index] is not double) return false;
        }
        return items.Count > 0;
    }

    private static string ApplyEdits(string content, IReadOnlyList<TextShowEdit> edits) {
        var builder = new StringBuilder(content.Length);
        int cursor = 0;
        foreach (TextShowEdit edit in edits.OrderBy(static value => value.Index)) {
            builder.Append(content, cursor, edit.Index - cursor);
            builder.Append(edit.Replacement);
            cursor = edit.Index + edit.Length;
        }
        builder.Append(content, cursor, content.Length - cursor);
        return builder.ToString();
    }

    private static int FindOperationStart(string content, int previousOperatorEnd, int operatorOffset) {
        int index = Math.Max(0, previousOperatorEnd);
        while (index < operatorOffset) {
            if (char.IsWhiteSpace(content[index])) {
                index++;
                continue;
            }
            if (content[index] == '%') {
                while (index < operatorOffset && content[index] != '\r' && content[index] != '\n') index++;
                continue;
            }
            break;
        }
        return index;
    }

    private static string BuildTransformPrefix(Matrix2D transform) => string.Format(
        CultureInfo.InvariantCulture,
        "q {0} {1} {2} {3} {4} {5} cm ",
        transform.A,
        transform.B,
        transform.C,
        transform.D,
        transform.E,
        transform.F);

    private static bool IsTextShowOperator(string name) => name is "Tj" or "TJ" or "'" or "\"";

    private static double NumberAt(IReadOnlyList<object> values, int index, double fallback) =>
        index >= 0 && index < values.Count && values[index] is double value ? value : fallback;

    private static string FormatNumber(double value) =>
        (Math.Abs(value) <= 0.0000001D ? 0D : value).ToString("0.########", CultureInfo.InvariantCulture);

    private static double Sum(IReadOnlyList<double> values, int start, int count) {
        double total = 0D;
        int end = Math.Min(values.Count, start + count);
        for (int index = start; index < end; index++) total += values[index];
        return total;
    }

    private static void FlushKeptBytes(List<string> output, List<byte> kept) {
        if (kept.Count == 0) return;
        var hex = new StringBuilder(kept.Count * 2);
        for (int index = 0; index < kept.Count; index++) {
            hex.Append(kept[index].ToString("X2", CultureInfo.InvariantCulture));
        }
        output.Add("<" + hex + ">");
        kept.Clear();
    }

    private static void FlushRemovedAdvance(List<string> output, ref double removedAdvance1000) {
        if (Math.Abs(removedAdvance1000) <= 0.0000001D) return;
        output.Add(FormatNumber(-removedAdvance1000));
        removedAdvance1000 = 0D;
    }

    private readonly struct EncodedGlyph {
        internal EncodedGlyph(byte[] bytes, string text, double width1000) {
            Bytes = bytes;
            Text = text;
            Width1000 = width1000;
        }

        internal byte[] Bytes { get; }
        internal string Text { get; }
        internal double Width1000 { get; }
    }

    private readonly struct TextShowEdit {
        internal TextShowEdit(int index, int length, string replacement) {
            Index = index;
            Length = length;
            Replacement = replacement;
        }

        internal int Index { get; }
        internal int Length { get; }
        internal string Replacement { get; }
    }
}

internal readonly struct PdfContentStreamTextRewriteTarget {
    internal PdfContentStreamTextRewriteTarget(PdfRedactionArea area, int? textRenderingMode = null) {
        Area = area;
        TextRenderingMode = textRenderingMode;
    }

    internal PdfRedactionArea Area { get; }
    internal int? TextRenderingMode { get; }

    internal bool MatchesRenderingMode(int textRenderingMode) =>
        !TextRenderingMode.HasValue || TextRenderingMode.Value == textRenderingMode;

    internal bool Intersects(PdfTextSpanBounds bounds, int textRenderingMode) =>
        MatchesRenderingMode(textRenderingMode) &&
        Area.IntersectsRectangle(bounds.Left, bounds.Bottom, bounds.Width, bounds.Height);
}

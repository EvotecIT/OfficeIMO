namespace OfficeIMO.IWork.Internal;

internal static partial class IWorkTextReader {
    private static (int Level, string? Label, string? FontName, IWorkListMarkerKind Kind) ResolveList(IWorkObjectIndex index,
        ulong? identifier, double? paragraphLeftIndentPoints, int? explicitLevel,
        IWorkProjectionBudget projectionBudget,
        Dictionary<(ulong Identifier, double? LeftIndentPoints, int? ExplicitLevel), Cached<(int Level, string? Label, string? FontName, IWorkListMarkerKind Kind)>> cache,
        Dictionary<ulong, Cached<ListStyleData>> decodedStyles,
        bool tolerateStyleDepth,
        IWorkSourceReferenceIssueCollector references,
        ref bool complete) {
        if (!identifier.HasValue) return (-1, null, null, IWorkListMarkerKind.None);
        var cacheKey = (identifier.Value, paragraphLeftIndentPoints, explicitLevel);
        if (cache.TryGetValue(cacheKey, out Cached<(int Level, string? Label, string? FontName, IWorkListMarkerKind Kind)> cached)) {
            if (!cached.IsComplete) complete = false;
            return cached.Value;
        }
        if (!decodedStyles.TryGetValue(identifier.Value, out Cached<ListStyleData> decoded)) {
            bool decodedCompletely = true;
            var style = new ListStyleData();
            var chain = IWorkStyleReader.ReadChain(index, identifier.Value,
                projectionBudget.MaximumTextStyleInheritanceDepth,
                type => type == ListStyleArchive, tolerateStyleDepth, references, ref decodedCompletely);
            for (int styleIndex = chain.Count - 1; styleIndex >= 0; styleIndex--) {
                IWorkWireMessage message = chain[styleIndex].Message;
                ApplyStyleName(message, value => style.Name = value, projectionBudget, chain[styleIndex].Record, references, ref decodedCompletely);
                OverlayList(message, style, projectionBudget,
                    new StylePropertyEvidence(chain[styleIndex].Record, "", references.Declarations),
                    ref decodedCompletely);
            }
            decoded = new Cached<ListStyleData>(style, decodedCompletely);
            decodedStyles.Add(identifier.Value, decoded);
        }
        bool resolvedCompletely = decoded.IsComplete;
        ListStyleData data = decoded.Value;
        int level = explicitLevel ?? ResolveListLevel(data, paragraphLeftIndentPoints, ref resolvedCompletely);
        if (level >= data.LabelTypes.Count) resolvedCompletely = false;
        ulong labelType = level >= 0 && level < data.LabelTypes.Count
            ? data.LabelTypes[level]
            : 0;
        string? selectedLabel = level >= 0 && level < data.Labels.Count
            ? data.Labels[level]
            : null;
        if (labelType == 1) {
            selectedLabel = null;
            resolvedCompletely = false;
            if (data.LabelDeclaration != null) data.LabelEvidence.Record(data.LabelDeclaration, 11,
                IWorkSourceDeclarationIssueKind.UnsupportedField);
        }
        if (labelType == 3) {
            selectedLabel = level >= 0 && level < data.NumberTypes.Count
                ? NumberMarker(data.NumberTypes[level]) : null;
            // Number kind is recoverable; paragraph-level starts/continuations and
            // tiered numbering are not yet qualified. Keep strict conversion gated.
            resolvedCompletely = false;
        }
        if (labelType != 0 && selectedLabel == null) resolvedCompletely = false;
        (int Level, string? Label, string? FontName, IWorkListMarkerKind Kind) result = labelType == 0
            || string.Equals(data.Name, "None", StringComparison.OrdinalIgnoreCase)
            ? (-1, null, null, IWorkListMarkerKind.None)
            : (level, selectedLabel, data.FontName, (IWorkListMarkerKind)labelType);
        cache.Add(cacheKey, new Cached<(int Level, string? Label, string? FontName, IWorkListMarkerKind Kind)>(result, resolvedCompletely));
        if (!resolvedCompletely) complete = false;
        return result;
    }

    // Each repeated field is one level-indexed vector. Removing a rejected entry would
    // assign its readable siblings to different levels, so only valid vectors overlay a parent.
    private static void OverlayList(IWorkWireMessage message, ListStyleData data,
        IWorkProjectionBudget projectionBudget, StylePropertyEvidence evidence, ref bool complete) {
        bool typesComplete = !message.HasUnexpectedWireKind(11, IWorkWireKind.Varint, IWorkWireKind.Bytes);
        IReadOnlyList<ulong> types = Array.Empty<ulong>();
        if (typesComplete) {
            try {
                types = message.GetRepeatedUnsigned(11, packed: true);
                typesComplete = types.All(type => type <= 3); // None, image, string, number.
            } catch (InvalidDataException exception) when (!IWorkProtobuf.IsLimitException(exception)) {
                typesComplete = false;
            }
        }
        if (!typesComplete) { evidence.Record(message, 11); complete = false; }
        else if (types.Count > 0) {
            data.LabelTypes = types;
            data.LabelDeclaration = message;
            data.LabelEvidence = evidence;
        }

        bool numbersComplete = !message.HasUnexpectedWireKind(15, IWorkWireKind.Varint, IWorkWireKind.Bytes);
        IReadOnlyList<ulong> numbers = Array.Empty<ulong>();
        if (numbersComplete) {
            try {
                numbers = message.GetRepeatedUnsigned(15, packed: true);
                numbersComplete = numbers.All(number => number <= 64);
            } catch (InvalidDataException exception) when (!IWorkProtobuf.IsLimitException(exception)) {
                numbersComplete = false;
            }
        }
        if (!numbersComplete) { evidence.Record(message, 15); complete = false; }
        else if (numbers.Count > 0) data.NumberTypes = numbers;

        bool labelsComplete = !message.HasUnexpectedWireKind(16, IWorkWireKind.Bytes);
        var labels = new List<string>();
        foreach (byte[] bytes in message.EnumerateRepeatedBytes(16)) {
            if (TryDecodeUtf8(bytes, projectionBudget, out string label)) labels.Add(label);
            else labelsComplete = false;
        }
        if (!labelsComplete) { evidence.Record(message, 16); complete = false; }
        else if (labels.Count > 0) data.Labels = labels;

        bool? clearFont = ReadBoolean(message, 22, evidence, ref complete);
        if (clearFont == true) {
            data.FontName = null;
            if (message.HasField(23)) { evidence.Record(message, 23); complete = false; }
        } else if ((!message.HasField(22) || clearFont.HasValue) && message.HasField(23)) {
            // A rejected child font must not silently reuse the parent's marker font.
            data.FontName = null;
            if (message.FieldCount(23) != 1
                || message.HasUnexpectedWireKind(23, IWorkWireKind.Bytes)
                || !TryDecodeUtf8(message.GetBytes(23)!, projectionBudget, out string fontName)
                || string.IsNullOrWhiteSpace(fontName)) {
                evidence.Record(message, 23); complete = false;
            } else data.FontName = fontName;
        } else if (message.HasField(22) && !clearFont.HasValue) {
            data.FontName = null;
        }

        IReadOnlyList<float> indents = message.GetRepeatedFloat(13);
        if (message.HasUnexpectedWireKind(13, IWorkWireKind.Fixed32) || indents.Any(indent => !IsFinite(indent))) {
            evidence.Record(message, 13); complete = false;
        } else if (indents.Count > 0) data.LeftIndents = indents;
    }

    private static int ResolveListLevel(ListStyleData data, double? paragraphLeftIndentPoints,
        ref bool complete) {
        if (data.LabelTypes.Count <= 1 || data.LabelTypes.All(type => type == 0)) return 0;
        if (!paragraphLeftIndentPoints.HasValue
            || data.LeftIndents.Count != data.LabelTypes.Count
            || data.LeftIndents.Any(indent => float.IsNaN(indent) || float.IsInfinity(indent))) {
            complete = false;
            return 0;
        }
        int bestLevel = 0;
        double bestDistance = double.MaxValue;
        for (int level = 0; level < data.LeftIndents.Count; level++) {
            double distance = Math.Abs(data.LeftIndents[level] - paragraphLeftIndentPoints.Value);
            if (distance < bestDistance) {
                bestDistance = distance;
                bestLevel = level;
            }
        }
        if (bestDistance > 0.05d) complete = false;
        return bestLevel;
    }

    // TSWP.ListStyleArchive.NumberType: decimal, upper/lower Roman, upper/lower
    // alphabetic, each in dot, double-parenthesis and right-parenthesis variants.
    // Other scripts and circled numbering remain unassessed.
    private static string? NumberMarker(ulong numberType) => numberType switch {
        0 => "1.", 1 => "(1)", 2 => "1)",
        3 => "I.", 4 => "(I)", 5 => "I)",
        6 => "i.", 7 => "(i)", 8 => "i)",
        9 => "A.", 10 => "(A)", 11 => "A)",
        12 => "a.", 13 => "(a)", 14 => "a)",
        _ => null
    };

    private sealed class ListStyleData {
        internal string? Name;
        internal string? FontName;
        internal IWorkWireMessage? LabelDeclaration;
        internal StylePropertyEvidence LabelEvidence;
        internal IReadOnlyList<ulong> LabelTypes = Array.Empty<ulong>();
        internal IReadOnlyList<ulong> NumberTypes = Array.Empty<ulong>();
        internal IReadOnlyList<string> Labels = Array.Empty<string>();
        internal IReadOnlyList<float> LeftIndents = Array.Empty<float>();
    }

}

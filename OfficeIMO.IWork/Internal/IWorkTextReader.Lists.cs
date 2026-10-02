namespace OfficeIMO.IWork.Internal;

internal static partial class IWorkTextReader {
    private static (int Level, string? Label) ResolveList(IWorkObjectIndex index,
        ulong? identifier, double? paragraphLeftIndentPoints, int? explicitLevel,
        IWorkProjectionBudget projectionBudget,
        Dictionary<(ulong Identifier, double? LeftIndentPoints, int? ExplicitLevel), Cached<(int Level, string? Label)>> cache,
        bool tolerateStyleDepth,
        IWorkSourceReferenceIssueCollector references,
        ref bool complete) {
        if (!identifier.HasValue) return (-1, null);
        var cacheKey = (identifier.Value, paragraphLeftIndentPoints, explicitLevel);
        if (cache.TryGetValue(cacheKey, out Cached<(int Level, string? Label)> cached)) {
            if (!cached.IsComplete) complete = false;
            return cached.Value;
        }
        bool resolvedCompletely = true;
        var data = new ListStyleData();
        var chain = IWorkStyleReader.ReadChain(index, identifier.Value,
            projectionBudget.MaximumTextStyleInheritanceDepth,
            type => type == ListStyleArchive, tolerateStyleDepth, references, ref resolvedCompletely);
        for (int styleIndex = chain.Count - 1; styleIndex >= 0; styleIndex--) {
            IWorkWireMessage message = chain[styleIndex].Message;
            ApplyStyleName(message, value => data.Name = value, projectionBudget, chain[styleIndex].Record, references, ref resolvedCompletely);
            OverlayList(message, data, projectionBudget,
                new StylePropertyEvidence(chain[styleIndex].Record, "", references.Declarations),
                ref resolvedCompletely);
        }
        int level = explicitLevel ?? ResolveListLevel(data, paragraphLeftIndentPoints, ref resolvedCompletely);
        if (level >= data.LabelTypes.Count) resolvedCompletely = false;
        ulong labelType = level >= 0 && level < data.LabelTypes.Count
            ? data.LabelTypes[level]
            : 0;
        string? selectedLabel = level >= 0 && level < data.Labels.Count
            ? data.Labels[level]
            : null;
        if (labelType == 3 && level >= 0 && level < data.NumberTypes.Count) {
            selectedLabel = NumberMarker(data.NumberTypes[level]);
            // Number kind is recoverable; paragraph-level starts/continuations and
            // tiered numbering are not yet qualified. Keep strict conversion gated.
            resolvedCompletely = false;
        }
        if (labelType != 0 && selectedLabel == null) resolvedCompletely = false;
        (int Level, string? Label) result = labelType == 0
            || string.Equals(data.Name, "None", StringComparison.OrdinalIgnoreCase)
            ? (-1, null)
            : (level, selectedLabel);
        cache.Add(cacheKey, new Cached<(int Level, string? Label)>(result, resolvedCompletely));
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
        else if (types.Count > 0) data.LabelTypes = types;

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
        internal IReadOnlyList<ulong> LabelTypes = Array.Empty<ulong>();
        internal IReadOnlyList<ulong> NumberTypes = Array.Empty<ulong>();
        internal IReadOnlyList<string> Labels = Array.Empty<string>();
        internal IReadOnlyList<float> LeftIndents = Array.Empty<float>();
    }

}

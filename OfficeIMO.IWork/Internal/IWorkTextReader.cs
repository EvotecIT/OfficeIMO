namespace OfficeIMO.IWork.Internal;

internal static class IWorkTextReader {
    private static readonly System.Text.UTF8Encoding StrictUtf8 = new(false, true);
    private const uint CharacterStyleArchive = 2021;
    private const uint ParagraphStyleArchive = 2022;
    private const uint ListStyleArchive = 2023;
    private const uint HyperlinkArchive = 2032;

    internal static IWorkTextContent Read(IWorkObjectIndex index, IWorkArchiveRecord storage,
        IWorkProjectionBudget projectionBudget) {
        IWorkWireMessage message = index.Message(storage);
        bool complete = true;
        string text = ReadText(message, projectionBudget, ref complete);
        IReadOnlyList<AttributeBoundary> paragraphStyles = ReadObjectTable(message, 5, text.Length,
            projectionBudget, ref complete);
        IReadOnlyList<AttributeBoundary> listStyles = ReadObjectTable(message, 7, text.Length,
            projectionBudget, ref complete);
        IReadOnlyList<AttributeBoundary> characterStyles = ReadObjectTable(message, 8, text.Length,
            projectionBudget, ref complete);
        IReadOnlyList<AttributeBoundary> hyperlinks = ReadObjectTable(message, 11, text.Length,
            projectionBudget, ref complete);
        var paragraphStyleCache = new Dictionary<ulong, Cached<IWorkParagraphStyle>>();
        var listStyleCache = new Dictionary<(ulong Identifier, double? LeftIndentPoints),
            Cached<(int Level, string? Label)>>();
        var textStyleCache = new Dictionary<TextStyleCacheKey, Cached<IWorkTextStyle>>();
        var hyperlinkCache = new Dictionary<ulong, Cached<string?>>();
        var paragraphs = new List<IWorkTextParagraph>();
        foreach (TextSpan paragraph in ParagraphSpans(text)) {
            projectionBudget.AddTextItem();
            ulong? paragraphStyleId = ObjectAt(paragraphStyles, paragraph.Start, carryMissing: true);
            ulong? listStyleId = ObjectAt(listStyles, paragraph.Start, carryMissing: true);
            IWorkParagraphStyle paragraphStyle = ResolveParagraphStyle(index, paragraphStyleId,
                projectionBudget, paragraphStyleCache, ref complete);
            (int listLevel, string? listLabel) = ResolveList(index, listStyleId,
                paragraphStyle.LeftIndentPoints,
                projectionBudget, listStyleCache, ref complete);
            if (listLabel != null) projectionBudget.AddTextCharacters(listLabel.Length);
            var boundaries = new SortedSet<int> { paragraph.Start, paragraph.End };
            AddBoundaries(boundaries, characterStyles, paragraph.Start, paragraph.End);
            AddBoundaries(boundaries, hyperlinks, paragraph.Start, paragraph.End);
            int[] ordered = boundaries
                .Where(boundary => !SplitsSurrogatePair(text, boundary))
                .ToArray();
            if (ordered.Length != boundaries.Count) complete = false;
            var runs = new List<IWorkTextRun>();
            for (int runIndex = 0; runIndex + 1 < ordered.Length; runIndex++) {
                int start = ordered[runIndex];
                int end = ordered[runIndex + 1];
                if (end <= start) continue;
                string runText = NormalizeInlineText(text.Substring(start, end - start),
                    projectionBudget, ref complete);
                if (runText.Length == 0) continue;
                projectionBudget.AddTextItem();
                ulong? characterStyleId = ObjectAt(characterStyles, start, carryMissing: false);
                IWorkTextStyle characterStyle = ResolveTextStyle(index, characterStyleId,
                    paragraphStyle.TextStyle, projectionBudget,
                    textStyleCache, ref complete);
                if (characterStyle.FontName != null) {
                    projectionBudget.AddTextCharacters(characterStyle.FontName.Length);
                }
                string? hyperlink = ResolveHyperlink(index,
                    ObjectAt(hyperlinks, start, carryMissing: false), projectionBudget,
                    hyperlinkCache, ref complete);
                runs.Add(new IWorkTextRun(runText, characterStyle, hyperlink));
            }
            paragraphs.Add(new IWorkTextParagraph(runs, paragraphStyle, listStyleId,
                listLevel, listLabel, paragraph.BreakKind));
        }
        return new IWorkTextContent(paragraphs, complete);
    }

    private static string ReadText(IWorkWireMessage message, IWorkProjectionBudget projectionBudget,
        ref bool complete) {
        var parts = new List<string>();
        if (message.HasUnexpectedWireKind(3, IWorkWireKind.Bytes)) complete = false;
        foreach (byte[] bytes in message.EnumerateRepeatedBytes(3)) {
            if (TryDecodeUtf8(bytes, projectionBudget, out string part)) parts.Add(part);
            else complete = false;
        }
        return string.Concat(parts);
    }

    private static IReadOnlyList<AttributeBoundary> ReadObjectTable(IWorkWireMessage storage,
        int field, int textLength, IWorkProjectionBudget projectionBudget, ref bool complete) {
        if (!storage.HasField(field)) return Array.Empty<AttributeBoundary>();
        if (storage.HasUnexpectedWireKind(field, IWorkWireKind.Bytes)) {
            complete = false;
            return Array.Empty<AttributeBoundary>();
        }
        byte[] tableBytes = storage.GetBytes(field)!;
        int boundaryCount;
        int totalTableFieldCount;
        try {
            boundaryCount = storage.CountNestedFields(tableBytes, 1,
                out totalTableFieldCount);
        } catch (InvalidDataException) {
            complete = false;
            return Array.Empty<AttributeBoundary>();
        }
        projectionBudget.AddTextBoundaries(boundaryCount);
        if (storage.FieldCount(field) != 1 || totalTableFieldCount != boundaryCount) {
            complete = false;
            return Array.Empty<AttributeBoundary>();
        }
        IWorkWireMessage table;
        try {
            table = storage.ParseNestedMessage(tableBytes);
        } catch (InvalidDataException) {
            complete = false;
            return Array.Empty<AttributeBoundary>();
        }
        var result = new List<AttributeBoundary>();
        if (table.HasUnexpectedWireKind(1, IWorkWireKind.Bytes)) complete = false;
        foreach (byte[] entryBytes in table.EnumerateRepeatedBytes(1)) {
            IWorkWireMessage entry;
            try {
                entry = table.ParseNestedMessage(entryBytes);
            } catch (InvalidDataException) {
                complete = false;
                continue;
            }
            ulong? rawIndex = entry.GetUnsigned(1);
            if (entry.FieldCount(1) != 1
                || entry.HasUnexpectedWireKind(1, IWorkWireKind.Varint)
                || !rawIndex.HasValue || rawIndex.Value > int.MaxValue
                || rawIndex.Value > (ulong)textLength) {
                complete = false;
                continue;
            }
            bool hasObject = entry.HasField(2);
            bool malformedReference = false;
            IWorkWireMessage? reference = hasObject
                ? IWorkObjectIndex.TryGetMessage(entry, 2, out malformedReference)
                : null;
            if (hasObject && (entry.HasUnexpectedWireKind(2, IWorkWireKind.Bytes)
                    || malformedReference || reference?.FieldCount(1) != 1
                    || reference?.GetUnsigned(1) == null
                    || reference.HasUnexpectedWireKind(1, IWorkWireKind.Varint))) {
                complete = false;
                continue;
            }
            result.Add(new AttributeBoundary((int)rawIndex.Value,
                reference?.GetUnsigned(1), hasObject));
        }
        AttributeBoundary[] ordered = result.OrderBy(boundary => boundary.Index).ToArray();
        for (int index = 1; index < ordered.Length; index++) {
            if (ordered[index - 1].Index == ordered[index].Index) complete = false;
        }
        ulong? carried = null;
        foreach (AttributeBoundary boundary in ordered) {
            if (boundary.HasObject) carried = boundary.Identifier;
            boundary.CarriedIdentifier = carried;
        }
        return ordered;
    }

    private static ulong? ObjectAt(IReadOnlyList<AttributeBoundary> boundaries, int offset,
        bool carryMissing) {
        int upper = UpperBound(boundaries, offset);
        if (upper == 0) return null;
        AttributeBoundary boundary = boundaries[upper - 1];
        return carryMissing ? boundary.CarriedIdentifier
            : boundary.HasObject ? boundary.Identifier : null;
    }

    private static void AddBoundaries(SortedSet<int> destination,
        IReadOnlyList<AttributeBoundary> source, int start, int end) {
        int index = UpperBound(source, start);
        while (index < source.Count && source[index].Index < end) {
            destination.Add(source[index].Index);
            index++;
        }
    }

    private static int UpperBound(IReadOnlyList<AttributeBoundary> boundaries, int offset) {
        int low = 0;
        int high = boundaries.Count;
        while (low < high) {
            int middle = low + (high - low) / 2;
            if (boundaries[middle].Index <= offset) low = middle + 1;
            else high = middle;
        }
        return low;
    }

    private static bool SplitsSurrogatePair(string text, int offset) =>
        offset > 0 && offset < text.Length
        && char.IsHighSurrogate(text[offset - 1])
        && char.IsLowSurrogate(text[offset]);

    private static IWorkParagraphStyle ResolveParagraphStyle(IWorkObjectIndex index,
        ulong? identifier, IWorkProjectionBudget projectionBudget,
        Dictionary<ulong, Cached<IWorkParagraphStyle>> cache,
        ref bool complete) {
        if (!identifier.HasValue) return new ParagraphStyleData().ToPublic();
        if (cache.TryGetValue(identifier.Value, out Cached<IWorkParagraphStyle> cached)) {
            if (!cached.IsComplete) complete = false;
            return cached.Value;
        }
        bool resolvedCompletely = true;
        var data = new ParagraphStyleData();
        IReadOnlyList<IWorkWireMessage> chain = ReadStyleChain(index, identifier.Value,
            projectionBudget.MaximumTextStyleInheritanceDepth,
            type => type == ParagraphStyleArchive, ref resolvedCompletely);
        for (int styleIndex = chain.Count - 1; styleIndex >= 0; styleIndex--) {
            IWorkWireMessage message = chain[styleIndex];
            ApplyStyleName(message, value => data.Name = value, projectionBudget, ref resolvedCompletely);
            IWorkWireMessage? character = IWorkObjectIndex.TryGetMessage(message, 11, out bool malformedCharacter);
            if (malformedCharacter || message.HasUnexpectedWireKind(11, IWorkWireKind.Bytes)
                || message.HasField(11) && character == null) resolvedCompletely = false;
            if (character != null) OverlayText(character, data.Text, projectionBudget, ref resolvedCompletely);
            IWorkWireMessage? paragraph = IWorkObjectIndex.TryGetMessage(message, 12, out bool malformedParagraph);
            if (malformedParagraph || message.HasUnexpectedWireKind(12, IWorkWireKind.Bytes)
                || message.HasField(12) && paragraph == null) resolvedCompletely = false;
            if (paragraph != null) OverlayParagraph(paragraph, data, ref resolvedCompletely);
        }
        IWorkParagraphStyle result = data.ToPublic();
        cache.Add(identifier.Value, new Cached<IWorkParagraphStyle>(result, resolvedCompletely));
        if (!resolvedCompletely) complete = false;
        return result;
    }

    private static IWorkTextStyle ResolveTextStyle(IWorkObjectIndex index, ulong? identifier,
        IWorkTextStyle inherited, IWorkProjectionBudget projectionBudget,
        Dictionary<TextStyleCacheKey, Cached<IWorkTextStyle>> cache,
        ref bool complete) {
        if (!identifier.HasValue) return inherited;
        var key = new TextStyleCacheKey(identifier.Value, inherited);
        if (cache.TryGetValue(key, out Cached<IWorkTextStyle> cached)) {
            if (!cached.IsComplete) complete = false;
            return cached.Value;
        }
        bool resolvedCompletely = true;
        var data = TextStyleData.From(inherited);
        IReadOnlyList<IWorkWireMessage> chain = ReadStyleChain(index, identifier.Value,
            projectionBudget.MaximumTextStyleInheritanceDepth,
            type => type is CharacterStyleArchive or ParagraphStyleArchive,
            ref resolvedCompletely);
        for (int styleIndex = chain.Count - 1; styleIndex >= 0; styleIndex--) {
            IWorkWireMessage message = chain[styleIndex];
            ApplyStyleName(message, value => data.Name = value, projectionBudget, ref resolvedCompletely);
            IWorkWireMessage? character = IWorkObjectIndex.TryGetMessage(message, 11, out bool malformedCharacter);
            if (malformedCharacter || message.HasUnexpectedWireKind(11, IWorkWireKind.Bytes)
                || message.HasField(11) && character == null) resolvedCompletely = false;
            if (character != null) OverlayText(character, data, projectionBudget, ref resolvedCompletely);
        }
        IWorkTextStyle result = data.ToPublic();
        cache.Add(key, new Cached<IWorkTextStyle>(result, resolvedCompletely));
        if (!resolvedCompletely) complete = false;
        return result;
    }

    private static IReadOnlyList<IWorkWireMessage> ReadStyleChain(IWorkObjectIndex index,
        ulong identifier, int maximumDepth, Func<uint, bool> allowedType, ref bool complete) {
        var chain = new List<IWorkWireMessage>();
        var seen = new HashSet<ulong>();
        ulong current = identifier;
        while (true) {
            if (chain.Count >= maximumDepth) {
                throw new InvalidDataException(
                    $"iWork text style inheritance exceeds the configured depth of {maximumDepth}.");
            }
            if (!seen.Add(current)) {
                complete = false;
                break;
            }
            IWorkArchiveRecord? record = index.Find(current);
            if (record == null || !allowedType(record.MessageType)) {
                complete = false;
                break;
            }
            IWorkWireMessage message = index.Message(record);
            chain.Add(message);
            IWorkWireMessage? super = IWorkObjectIndex.TryGetMessage(message, 1, out bool malformedSuper);
            if (malformedSuper || message.HasUnexpectedWireKind(1, IWorkWireKind.Bytes)
                || message.HasField(1) && super == null) {
                complete = false;
                break;
            }
            if (super == null) break;
            IWorkArchiveRecord? parent = index.Dereference(super, 3);
            if (super.HasUnexpectedWireKind(3, IWorkWireKind.Bytes)
                || super.HasField(3) && parent == null) {
                complete = false;
                break;
            }
            if (parent == null) break;
            current = parent.Identifier;
        }
        return chain;
    }

    private static void ApplyStyleName(IWorkWireMessage message, Action<string> apply,
        IWorkProjectionBudget projectionBudget, ref bool complete) {
        IWorkWireMessage? super = IWorkObjectIndex.TryGetMessage(message, 1, out bool malformedSuper);
        if (malformedSuper || message.HasUnexpectedWireKind(1, IWorkWireKind.Bytes)
            || message.HasField(1) && super == null) {
            complete = false;
            return;
        }
        if (super == null || !super.HasField(1)) return;
        if (super.FieldCount(1) != 1
            || super.HasUnexpectedWireKind(1, IWorkWireKind.Bytes)
            || !TryDecodeUtf8(super.GetBytes(1)!, projectionBudget, out string name)) complete = false;
        else apply(name);
    }

    private static void OverlayText(IWorkWireMessage message, TextStyleData data,
        IWorkProjectionBudget projectionBudget, ref bool complete) {
        OverlayFlag(message, 1, value => data.Bold = value, ref complete);
        OverlayFlag(message, 2, value => data.Italic = value, ref complete);
        OverlayFlag(message, 11, value => data.Underline = value, ref complete);
        OverlayFlag(message, 12, value => data.Strikethrough = value, ref complete);
        float? size = message.GetFloat(3);
        if (message.FieldCount(3) > 1
            || message.HasUnexpectedWireKind(3, IWorkWireKind.Fixed32)) complete = false;
        else if (size.HasValue && IsFinitePositive(size.Value)) data.FontSizePoints = size.Value;
        else if (message.HasField(3)) complete = false;
        bool? clearFont = ReadBoolean(message, 4, ref complete);
        if (clearFont == true) {
            if (message.HasField(5)) complete = false;
            data.FontName = null;
        }
        else if (message.HasField(5)) {
            if (message.FieldCount(5) != 1
                || message.HasUnexpectedWireKind(5, IWorkWireKind.Bytes)
                || !TryDecodeUtf8(message.GetBytes(5)!, projectionBudget, out string fontName)) complete = false;
            else data.FontName = fontName;
        }
        bool? clearColor = ReadBoolean(message, 6, ref complete);
        if (clearColor == true) {
            if (message.HasField(7)) complete = false;
            data.Color = null;
        }
        else if (TryColor(message, 7, out IWorkColor? color, ref complete)) data.Color = color;
        bool? clearBackground = ReadBoolean(message, 25, ref complete);
        if (clearBackground == true) {
            if (message.HasField(26)) complete = false;
            data.BackgroundColor = null;
        }
        else if (TryColor(message, 26, out IWorkColor? background, ref complete)) data.BackgroundColor = background;
    }

    private static void OverlayParagraph(IWorkWireMessage message, ParagraphStyleData data,
        ref bool complete) {
        ulong? alignment = ReadUnsigned(message, 1, ref complete);
        if (alignment.HasValue) {
            if (alignment.Value > 4) complete = false;
            else data.Alignment = alignment.Value switch {
                0 => IWorkTextAlignment.Left,
                1 => IWorkTextAlignment.Right,
                2 => IWorkTextAlignment.Center,
                3 => IWorkTextAlignment.Justified,
                _ => IWorkTextAlignment.Natural
            };
        }
        OverlayFinite(message, 7, value => data.FirstLineIndentPoints = value, ref complete);
        OverlayFinite(message, 11, value => data.LeftIndentPoints = value, ref complete);
        OverlayFinite(message, 19, value => data.RightIndentPoints = value, ref complete);
        OverlayFinite(message, 20, value => data.SpaceAfterPoints = value, ref complete);
        OverlayFinite(message, 21, value => data.SpaceBeforePoints = value, ref complete);
        OverlayFlag(message, 14, value => data.PageBreakBefore = value, ref complete);
        OverlayFlag(message, 9, value => data.KeepLinesTogether = value, ref complete);
        OverlayFlag(message, 10, value => data.KeepWithNext = value, ref complete);
    }

    private static (int Level, string? Label) ResolveList(IWorkObjectIndex index,
        ulong? identifier, double? paragraphLeftIndentPoints,
        IWorkProjectionBudget projectionBudget,
        Dictionary<(ulong Identifier, double? LeftIndentPoints), Cached<(int Level, string? Label)>> cache,
        ref bool complete) {
        if (!identifier.HasValue) return (-1, null);
        var cacheKey = (identifier.Value, paragraphLeftIndentPoints);
        if (cache.TryGetValue(cacheKey, out Cached<(int Level, string? Label)> cached)) {
            if (!cached.IsComplete) complete = false;
            return cached.Value;
        }
        bool resolvedCompletely = true;
        var data = new ListStyleData();
        IReadOnlyList<IWorkWireMessage> chain = ReadStyleChain(index, identifier.Value,
            projectionBudget.MaximumTextStyleInheritanceDepth,
            type => type == ListStyleArchive, ref resolvedCompletely);
        for (int styleIndex = chain.Count - 1; styleIndex >= 0; styleIndex--) {
            IWorkWireMessage message = chain[styleIndex];
            ApplyStyleName(message, value => data.Name = value, projectionBudget, ref resolvedCompletely);
            IReadOnlyList<ulong> labelTypes;
            if (message.HasUnexpectedWireKind(11, IWorkWireKind.Varint, IWorkWireKind.Bytes)) {
                resolvedCompletely = false;
                labelTypes = Array.Empty<ulong>();
            } else {
                try {
                    labelTypes = message.GetRepeatedUnsigned(11, packed: true);
                } catch (InvalidDataException) {
                    resolvedCompletely = false;
                    labelTypes = Array.Empty<ulong>();
                }
            }
            if (labelTypes.Count > 0) data.LabelTypes = labelTypes;
            if (message.HasUnexpectedWireKind(16, IWorkWireKind.Bytes)) resolvedCompletely = false;
            var labels = new List<string>();
            foreach (byte[] bytes in message.EnumerateRepeatedBytes(16)) {
                if (!TryDecodeUtf8(bytes, projectionBudget, out string decodedLabel)) resolvedCompletely = false;
                else labels.Add(decodedLabel);
            }
            if (labels.Count > 0) data.Labels = labels;
            if (message.HasUnexpectedWireKind(13, IWorkWireKind.Fixed32)) resolvedCompletely = false;
            IReadOnlyList<float> indents = message.GetRepeatedFloat(13);
            if (indents.Count > 0) data.LeftIndents = indents;
        }
        int level = ResolveListLevel(data, paragraphLeftIndentPoints, ref resolvedCompletely);
        ulong labelType = level >= 0 && level < data.LabelTypes.Count
            ? data.LabelTypes[level]
            : 0;
        string? selectedLabel = level >= 0 && level < data.Labels.Count
            ? data.Labels[level]
            : null;
        if (labelType != 0 && selectedLabel == null) resolvedCompletely = false;
        (int Level, string? Label) result = labelType == 0
            || string.Equals(data.Name, "None", StringComparison.OrdinalIgnoreCase)
            ? (-1, null)
            : (level, selectedLabel);
        cache.Add(cacheKey, new Cached<(int Level, string? Label)>(result, resolvedCompletely));
        if (!resolvedCompletely) complete = false;
        return result;
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

    private static string? ResolveHyperlink(IWorkObjectIndex index, ulong? identifier,
        IWorkProjectionBudget projectionBudget, Dictionary<ulong, Cached<string?>> cache,
        ref bool complete) {
        if (!identifier.HasValue) return null;
        if (cache.TryGetValue(identifier.Value, out Cached<string?> cached)) {
            if (!cached.IsComplete) complete = false;
            if (cached.Value != null) projectionBudget.AddTextCharacters(cached.Value.Length);
            return cached.Value;
        }
        bool resolvedCompletely = true;
        string? result = null;
        IWorkArchiveRecord? record = index.Find(identifier.Value);
        if (record == null || record.MessageType != HyperlinkArchive) {
            resolvedCompletely = false;
        } else {
            IWorkWireMessage message = index.Message(record);
            if (message.FieldCount(2) != 1
                || message.HasUnexpectedWireKind(2, IWorkWireKind.Bytes)
                || !TryDecodeUtf8(message.GetBytes(2)!, projectionBudget, out string value)) {
                resolvedCompletely = false;
            } else {
                result = value;
            }
        }
        cache.Add(identifier.Value, new Cached<string?>(result, resolvedCompletely));
        if (!resolvedCompletely) complete = false;
        return result;
    }

    private static bool TryColor(IWorkWireMessage owner, int field, out IWorkColor? color,
        ref bool complete) {
        color = null;
        bool hasColor = owner.HasField(field);
        IWorkWireMessage? message = IWorkObjectIndex.TryGetMessage(owner, field, out bool malformedColor);
        if (owner.HasUnexpectedWireKind(field, IWorkWireKind.Bytes)
            || malformedColor || hasColor && message == null) {
            complete = false;
            return false;
        }
        if (message == null) return false;
        bool hasWhite = message.HasField(11);
        bool hasAnyRgb = message.HasField(3) || message.HasField(4) || message.HasField(5);
        bool hasCompleteRgb = message.HasField(3) && message.HasField(4) && message.HasField(5);
        float? white = message.GetFloat(11);
        float red = white ?? message.GetFloat(3) ?? 0;
        float green = white ?? message.GetFloat(4) ?? 0;
        float blue = white ?? message.GetFloat(5) ?? 0;
        float alpha = message.GetFloat(6) ?? 1;
        if (new[] { 3, 4, 5, 6, 11 }.Any(component =>
                message.FieldCount(component) > 1
                || message.HasUnexpectedWireKind(component, IWorkWireKind.Fixed32)
                || message.HasField(component) && !message.GetFloat(component).HasValue)
            || hasWhite == hasAnyRgb
            || hasAnyRgb && !hasCompleteRgb
            || !new[] { red, green, blue, alpha }.All(IsNormalizedColorComponent)) {
            complete = false;
            return false;
        }
        color = new IWorkColor(Component(red), Component(green), Component(blue), AlphaComponent(alpha));
        return true;
    }

    private static byte AlphaComponent(float value) => value >= 1f
        ? byte.MaxValue
        : (byte)Math.Floor(Math.Max(0f, value) * byte.MaxValue);

    private static void OverlayFinite(IWorkWireMessage message, int field, Action<double> apply,
        ref bool complete) {
        float? value = message.GetFloat(field);
        if (message.FieldCount(field) > 1
            || message.HasUnexpectedWireKind(field, IWorkWireKind.Fixed32)) complete = false;
        else if (value.HasValue && IsFinite(value.Value)) apply(value.Value);
        else if (message.HasField(field)) complete = false;
    }

    private static ulong? ReadUnsigned(IWorkWireMessage message, int field, ref bool complete) {
        if (message.FieldCount(field) > 1
            || message.HasUnexpectedWireKind(field, IWorkWireKind.Varint)) complete = false;
        return message.GetUnsigned(field);
    }

    private static void OverlayFlag(IWorkWireMessage message, int field, Action<bool> apply,
        ref bool complete) {
        bool? value = ReadBoolean(message, field, ref complete);
        if (value.HasValue) apply(value.Value);
    }

    private static bool? ReadBoolean(IWorkWireMessage message, int field, ref bool complete) {
        ulong? value = ReadUnsigned(message, field, ref complete);
        if (value > 1) {
            complete = false;
            return null;
        }
        return value.HasValue ? value.Value == 1 : null;
    }

    private static IEnumerable<TextSpan> ParagraphSpans(string text) {
        if (text.Length == 0) yield break;
        int start = 0;
        for (int index = 0; index < text.Length; index++) {
            IWorkParagraphBreakKind kind = BreakKind(text[index]);
            if (kind == IWorkParagraphBreakKind.None) continue;
            yield return new TextSpan(start, index, kind);
            if (text[index] == '\r' && index + 1 < text.Length && text[index + 1] == '\n') index++;
            start = index + 1;
        }
        if (start <= text.Length) yield return new TextSpan(start, text.Length, IWorkParagraphBreakKind.None);
    }

    private static IWorkParagraphBreakKind BreakKind(char value) => value switch {
        '\n' => IWorkParagraphBreakKind.Paragraph,
        '\r' => IWorkParagraphBreakKind.Paragraph,
        '\u2029' => IWorkParagraphBreakKind.Paragraph,
        '\u0004' => IWorkParagraphBreakKind.Section,
        '\u0005' => IWorkParagraphBreakKind.Layout,
        '\u000c' => IWorkParagraphBreakKind.Page,
        _ => IWorkParagraphBreakKind.None
    };

    private static bool IsNormalizedColorComponent(float value) =>
        IsFinite(value) && value >= 0f && value <= 1f;

    private static string NormalizeInlineText(string value, IWorkProjectionBudget projectionBudget,
        ref bool complete) {
        if (value.IndexOf('\ufffc') >= 0 || value.IndexOf('\ufffb') >= 0) complete = false;
        int inlineBreakCount = 0;
        foreach (char character in value) {
            if (character == '\u2028') inlineBreakCount++;
        }
        projectionBudget.AddTextItems(inlineBreakCount);
        return value.Replace('\u2028', '\n')
            .Replace("\ufffc", string.Empty)
            .Replace("\ufffb", string.Empty);
    }

    internal static bool TryDecodeUtf8(byte[] bytes, IWorkProjectionBudget projectionBudget,
        out string value) {
        try {
            int characterCount = StrictUtf8.GetCharCount(bytes);
            projectionBudget.AddTextCharacters(characterCount);
            value = StrictUtf8.GetString(bytes);
            return IWorkXmlText.IsRepresentable(value, allowIWorkBreaks: true);
        } catch (System.Text.DecoderFallbackException) {
            value = string.Empty;
            return false;
        }
    }

    private static double? Finite(float? value) => value.HasValue && IsFinite(value.Value)
        ? value.Value
        : (double?)null;
    private static bool IsFinitePositive(float value) => IsFinite(value) && value > 0;
    private static bool IsFinite(float value) => !float.IsNaN(value) && !float.IsInfinity(value);
    private static byte Component(float value) => (byte)Math.Round(Math.Max(0, Math.Min(1, value)) * 255,
        MidpointRounding.AwayFromZero);

    private sealed class AttributeBoundary {
        internal AttributeBoundary(int index, ulong? identifier, bool hasObject) {
            Index = index;
            Identifier = identifier;
            HasObject = hasObject;
        }
        internal int Index { get; }
        internal ulong? Identifier { get; }
        internal bool HasObject { get; }
        internal ulong? CarriedIdentifier { get; set; }
    }

    private readonly struct Cached<T> {
        internal Cached(T value, bool isComplete) {
            Value = value;
            IsComplete = isComplete;
        }
        internal T Value { get; }
        internal bool IsComplete { get; }
    }

    private readonly struct TextStyleCacheKey : IEquatable<TextStyleCacheKey> {
        private readonly ulong _identifier;
        private readonly string? _name;
        private readonly bool? _bold;
        private readonly bool? _italic;
        private readonly bool? _underline;
        private readonly bool? _strikethrough;
        private readonly double? _fontSizePoints;
        private readonly string? _fontName;
        private readonly uint? _color;
        private readonly uint? _backgroundColor;

        internal TextStyleCacheKey(ulong identifier, IWorkTextStyle inherited) {
            _identifier = identifier;
            _name = inherited.Name;
            _bold = inherited.Bold;
            _italic = inherited.Italic;
            _underline = inherited.Underline;
            _strikethrough = inherited.Strikethrough;
            _fontSizePoints = inherited.FontSizePoints;
            _fontName = inherited.FontName;
            _color = Pack(inherited.Color);
            _backgroundColor = Pack(inherited.BackgroundColor);
        }

        public bool Equals(TextStyleCacheKey other) => _identifier == other._identifier
            && string.Equals(_name, other._name, StringComparison.Ordinal)
            && _bold == other._bold && _italic == other._italic
            && _underline == other._underline && _strikethrough == other._strikethrough
            && _fontSizePoints == other._fontSizePoints
            && string.Equals(_fontName, other._fontName, StringComparison.Ordinal)
            && _color == other._color && _backgroundColor == other._backgroundColor;

        public override bool Equals(object? obj) => obj is TextStyleCacheKey other && Equals(other);

        public override int GetHashCode() {
            unchecked {
                int hash = _identifier.GetHashCode();
                hash = hash * 31 + (_name?.GetHashCode() ?? 0);
                hash = hash * 31 + _bold.GetHashCode();
                hash = hash * 31 + _italic.GetHashCode();
                hash = hash * 31 + _underline.GetHashCode();
                hash = hash * 31 + _strikethrough.GetHashCode();
                hash = hash * 31 + _fontSizePoints.GetHashCode();
                hash = hash * 31 + (_fontName?.GetHashCode() ?? 0);
                hash = hash * 31 + _color.GetHashCode();
                return hash * 31 + _backgroundColor.GetHashCode();
            }
        }

        private static uint? Pack(IWorkColor? color) => color == null
            ? null
            : (uint)(color.Red << 24 | color.Green << 16 | color.Blue << 8 | color.Alpha);
    }

    private sealed class TextStyleData {
        internal string? Name;
        internal bool? Bold;
        internal bool? Italic;
        internal bool? Underline;
        internal bool? Strikethrough;
        internal double? FontSizePoints;
        internal string? FontName;
        internal IWorkColor? Color;
        internal IWorkColor? BackgroundColor;

        internal static TextStyleData From(IWorkTextStyle style) => new() {
            Name = style.Name, Bold = style.Bold, Italic = style.Italic,
            Underline = style.Underline, Strikethrough = style.Strikethrough,
            FontSizePoints = style.FontSizePoints, FontName = style.FontName,
            Color = style.Color, BackgroundColor = style.BackgroundColor
        };

        internal IWorkTextStyle ToPublic() => new(Name, Bold, Italic, Underline,
            Strikethrough, FontSizePoints, FontName, Color, BackgroundColor);
    }

    private sealed class ParagraphStyleData {
        internal string? Name;
        internal IWorkTextAlignment? Alignment;
        internal double? FirstLineIndentPoints;
        internal double? LeftIndentPoints;
        internal double? RightIndentPoints;
        internal double? SpaceBeforePoints;
        internal double? SpaceAfterPoints;
        internal bool? PageBreakBefore;
        internal bool? KeepWithNext;
        internal bool? KeepLinesTogether;
        internal TextStyleData Text { get; } = new();

        internal IWorkParagraphStyle ToPublic() => new(Name, Alignment,
            FirstLineIndentPoints, LeftIndentPoints, RightIndentPoints,
            SpaceBeforePoints, SpaceAfterPoints, PageBreakBefore, KeepWithNext,
            KeepLinesTogether, Text.ToPublic());
    }

    private sealed class ListStyleData {
        internal string? Name;
        internal IReadOnlyList<ulong> LabelTypes = Array.Empty<ulong>();
        internal IReadOnlyList<string> Labels = Array.Empty<string>();
        internal IReadOnlyList<float> LeftIndents = Array.Empty<float>();
    }

    private sealed class TextSpan {
        internal TextSpan(int start, int end, IWorkParagraphBreakKind breakKind) {
            Start = start;
            End = end;
            BreakKind = breakKind;
        }
        internal int Start { get; }
        internal int End { get; }
        internal IWorkParagraphBreakKind BreakKind { get; }
    }
}

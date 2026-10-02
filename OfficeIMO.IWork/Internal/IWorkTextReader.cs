namespace OfficeIMO.IWork.Internal;

internal static partial class IWorkTextReader {
    private static readonly System.Text.UTF8Encoding StrictUtf8 = new(false, true);
    private const uint CharacterStyleArchive = 2021;
    private const uint ParagraphStyleArchive = 2022;
    private const uint ListStyleArchive = 2023;
    private const uint HyperlinkArchive = 2032;

    internal static IWorkTextContent Read(IWorkObjectIndex index, IWorkArchiveRecord storage,
        IWorkProjectionBudget projectionBudget, IWorkSourceReferenceIssueCollector references,
        bool tolerateStyleDepth = false, bool resolveInlineObjects = false) {
        IWorkWireMessage message = index.Message(storage);
        bool textComplete = true;
        string text = ReadText(message, projectionBudget, ref textComplete);
        bool hasInvalidSourceText = !textComplete;
        bool hasUnresolvedInlineObjects = text.IndexOf('\ufffc') >= 0 || text.IndexOf('\ufffb') >= 0;
        bool complete = true;
        IReadOnlyDictionary<int, IWorkInlineObject> inlineObjects = new Dictionary<int, IWorkInlineObject>();
        if (resolveInlineObjects) {
            inlineObjects = ReadInlineObjects(index, message, text, storage, projectionBudget, references, out bool attachmentsComplete);
            complete &= attachmentsComplete;
            hasUnresolvedInlineObjects = !attachmentsComplete;
            for (int offset = 0; offset < text.Length; offset++) {
                if (text[offset] == '\ufffb' || text[offset] == '\ufffc' && !inlineObjects.ContainsKey(offset)) hasUnresolvedInlineObjects = true;
            }
        }
        IReadOnlyList<AttributeBoundary> paragraphStyles = ReadObjectTable(message, 5, text.Length,
            storage, projectionBudget, references, ref complete, static type => type == ParagraphStyleArchive);
        IReadOnlyList<AttributeBoundary> listLevels = ReadListLevelTable(message, text.Length,
            storage, projectionBudget, references, ref complete);
        IReadOnlyList<AttributeBoundary> listStyles = ReadObjectTable(message, 7, text.Length,
            storage, projectionBudget, references, ref complete, static type => type == ListStyleArchive);
        IReadOnlyList<AttributeBoundary> characterStyles = ReadObjectTable(message, 8, text.Length,
            storage, projectionBudget, references, ref complete, static type => type is CharacterStyleArchive or ParagraphStyleArchive);
        IReadOnlyList<AttributeBoundary> hyperlinks = ReadObjectTable(message, 11, text.Length,
            storage, projectionBudget, references, ref complete, static type => type == HyperlinkArchive);
        var paragraphStyleCache = new Dictionary<ulong, Cached<IWorkParagraphStyle>>();
        var listStyleCache = new Dictionary<(ulong Identifier, double? LeftIndentPoints, int? ExplicitLevel),
            Cached<(int Level, string? Label)>>();
        var textStyleCache = new Dictionary<TextStyleCacheKey, Cached<IWorkTextStyle>>();
        var hyperlinkCache = new Dictionary<ulong, Cached<string?>>();
        var inlineOffsets = new SortedSet<int>(inlineObjects.Keys);
        var paragraphs = new List<IWorkTextParagraph>();
        foreach (TextSpan paragraph in ParagraphSpans(text)) {
            projectionBudget.AddTextItem();
            ulong? paragraphStyleId = ObjectAt(paragraphStyles, paragraph.Start, carryMissing: true);
            ulong? listStyleId = ObjectAt(listStyles, paragraph.Start, carryMissing: true);
            IWorkParagraphStyle paragraphStyle = ResolveParagraphStyle(index, paragraphStyleId,
                projectionBudget, paragraphStyleCache, tolerateStyleDepth, references, ref complete);
            (int listLevel, string? listLabel) = ResolveList(index, listStyleId,
                paragraphStyle.LeftIndentPoints,
                (int?)ObjectAt(listLevels, paragraph.Start, carryMissing: false),
                projectionBudget, listStyleCache, tolerateStyleDepth, references, ref complete);
            if (listLabel != null) projectionBudget.AddTextCharacters(listLabel.Length);
            var boundaries = new SortedSet<int> { paragraph.Start, paragraph.End };
            AddBoundaries(boundaries, characterStyles, paragraph.Start, paragraph.End);
            AddBoundaries(boundaries, hyperlinks, paragraph.Start, paragraph.End);
            if (paragraph.End > paragraph.Start) {
                foreach (int offset in inlineOffsets.GetViewBetween(paragraph.Start, paragraph.End - 1)) {
                    boundaries.Add(offset); boundaries.Add(offset + 1);
                }
            }
            int[] ordered = boundaries
                .Where(boundary => !SplitsSurrogatePair(text, boundary))
                .ToArray();
            if (ordered.Length != boundaries.Count) complete = false;
            var runs = new List<IWorkTextRun>();
            for (int runIndex = 0; runIndex + 1 < ordered.Length; runIndex++) {
                int start = ordered[runIndex];
                int end = ordered[runIndex + 1];
                if (end <= start) continue;
                inlineObjects.TryGetValue(start, out IWorkInlineObject? inlineObject);
                string runText = inlineObject != null ? string.Empty : NormalizeInlineText(text.Substring(start, end - start),
                    projectionBudget, ref textComplete);
                if (runText.Length == 0 && inlineObject == null) continue;
                projectionBudget.AddTextItem();
                ulong? characterStyleId = ObjectAt(characterStyles, start, carryMissing: false);
                IWorkTextStyle characterStyle = ResolveTextStyle(index, characterStyleId,
                    paragraphStyle.TextStyle, projectionBudget,
                    textStyleCache, tolerateStyleDepth, references, ref complete);
                if (characterStyle.FontName != null) {
                    projectionBudget.AddTextCharacters(characterStyle.FontName.Length);
                }
                string? hyperlink = ResolveHyperlink(index,
                    ObjectAt(hyperlinks, start, carryMissing: false), projectionBudget,
                    hyperlinkCache, references, ref complete);
                runs.Add(new IWorkTextRun(runText, characterStyle, hyperlink, inlineObject));
            }
            paragraphs.Add(new IWorkTextParagraph(runs, paragraphStyle, listStyleId,
                listLevel, listLabel, paragraph.BreakKind));
        }
        return new IWorkTextContent(paragraphs, complete && textComplete, textComplete,
            hasInvalidSourceText, hasUnresolvedInlineObjects, isFormattingComplete: complete,
            sourceIdentity: new IWorkObjectIdentity(storage));
    }

    /// <summary>Reads table-cell text without traversing formatting references.</summary>
    internal static string ReadPlainText(IWorkWireMessage storage,
        IWorkProjectionBudget projectionBudget, out bool isComplete) {
        bool complete = true;
        string text = ReadText(storage, projectionBudget, ref complete);
        var paragraphs = new List<string>();
        foreach (TextSpan paragraph in ParagraphSpans(text)) {
            projectionBudget.AddTextItem();
            string value = NormalizeInlineText(text.Substring(paragraph.Start,
                paragraph.End - paragraph.Start), projectionBudget, ref complete);
            if (value.Length > 0) projectionBudget.AddTextItem();
            paragraphs.Add(value);
        }
        isComplete = complete;
        return string.Join("\n", paragraphs);
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

    private static bool SplitsSurrogatePair(string text, int offset) =>
        offset > 0 && offset < text.Length
        && char.IsHighSurrogate(text[offset - 1])
        && char.IsLowSurrogate(text[offset]);

    private static string? ResolveHyperlink(IWorkObjectIndex index, ulong? identifier,
        IWorkProjectionBudget projectionBudget, Dictionary<ulong, Cached<string?>> cache,
        IWorkSourceReferenceIssueCollector references, ref bool complete) {
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
            IWorkWireMessage? message = null;
            try {
                message = index.Message(record);
            } catch (InvalidDataException exception) when (!IWorkProtobuf.IsLimitException(exception)) {
                references.Declarations.Record(record, "$", null);
                resolvedCompletely = false;
            }
            if (message == null || message.FieldCount(2) != 1
                || message.HasUnexpectedWireKind(2, IWorkWireKind.Bytes)
                || !TryDecodeUtf8(message.GetBytes(2)!, projectionBudget, out string value)) {
                if (message != null) new StylePropertyEvidence(record, "", references.Declarations).Record(message, 2);
                resolvedCompletely = false;
            } else {
                result = value;
            }
        }
        cache.Add(identifier.Value, new Cached<string?>(result, resolvedCompletely));
        if (!resolvedCompletely) complete = false;
        return result;
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

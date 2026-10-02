namespace OfficeIMO.IWork.Internal;

internal static partial class IWorkTextReader {
    private static IWorkParagraphStyle ResolveParagraphStyle(IWorkObjectIndex index,
        ulong? identifier, IWorkProjectionBudget projectionBudget,
        Dictionary<ulong, ParagraphStyleCacheEntry> cache,
        bool tolerateStyleDepth,
        IWorkSourceReferenceIssueCollector references,
        ref bool complete) {
        if (!identifier.HasValue) return new ParagraphStyleData().ToPublic();
        if (cache.TryGetValue(identifier.Value, out ParagraphStyleCacheEntry cached)) {
            if (!cached.IsComplete) complete = false;
            return cached.Value;
        }
        bool resolvedCompletely = true;
        bool layoutComplete = true;
        var data = new ParagraphStyleData();
        var chain = IWorkStyleReader.ReadChain(index, identifier.Value,
            projectionBudget.MaximumTextStyleInheritanceDepth,
            type => type == ParagraphStyleArchive, tolerateStyleDepth, references, ref resolvedCompletely);
        for (int styleIndex = chain.Count - 1; styleIndex >= 0; styleIndex--) {
            IWorkWireMessage message = chain[styleIndex].Message;
            ApplyStyleName(message, value => data.Name = value, projectionBudget, chain[styleIndex].Record, references, ref resolvedCompletely);
            IWorkWireMessage? character = IWorkObjectIndex.TryGetMessage(message, 11, out bool malformedCharacter);
            if (malformedCharacter || message.FieldCount(11) > 1 || message.HasUnexpectedWireKind(11, IWorkWireKind.Bytes)
                || message.HasField(11) && character == null) {
                references.Declarations.Record(chain[styleIndex].Record, "11", message.FieldCount(11));
                resolvedCompletely = false;
            }
            if (character != null) OverlayText(character, data.Text, projectionBudget, chain[styleIndex].Record, references, ref resolvedCompletely);
            IWorkWireMessage? paragraph = IWorkObjectIndex.TryGetMessage(message, 12, out bool malformedParagraph);
            if (malformedParagraph || message.FieldCount(12) > 1 || message.HasUnexpectedWireKind(12, IWorkWireKind.Bytes)
                || message.HasField(12) && paragraph == null) {
                references.Declarations.Record(chain[styleIndex].Record, "12", message.FieldCount(12));
                resolvedCompletely = false;
            }
            if (paragraph != null) {
                OverlayParagraph(paragraph, data, chain[styleIndex].Record, references, ref resolvedCompletely);
                AssessParagraphLayout(paragraph, data, projectionBudget, chain[styleIndex].Record, references, ref layoutComplete);
            }
        }
        IWorkParagraphStyle result = data.ToPublic();
        cache.Add(identifier.Value, new ParagraphStyleCacheEntry(result, resolvedCompletely && layoutComplete, resolvedCompletely));
        if (!resolvedCompletely || !layoutComplete) complete = false;
        return result;
    }

    private static IWorkTextStyle ResolveTextStyle(IWorkObjectIndex index, ulong? identifier,
        IWorkTextStyle inherited, IWorkProjectionBudget projectionBudget,
        Dictionary<TextStyleCacheKey, Cached<IWorkTextStyle>> cache,
        bool tolerateStyleDepth,
        IWorkSourceReferenceIssueCollector references,
        ref bool complete) {
        if (!identifier.HasValue) return inherited;
        var key = new TextStyleCacheKey(identifier.Value, inherited);
        if (cache.TryGetValue(key, out Cached<IWorkTextStyle> cached)) {
            if (!cached.IsComplete) complete = false;
            return cached.Value;
        }
        bool resolvedCompletely = true;
        var data = TextStyleData.From(inherited);
        var chain = IWorkStyleReader.ReadChain(index, identifier.Value,
            projectionBudget.MaximumTextStyleInheritanceDepth,
            type => type is CharacterStyleArchive or ParagraphStyleArchive,
            tolerateStyleDepth, references, ref resolvedCompletely);
        for (int styleIndex = chain.Count - 1; styleIndex >= 0; styleIndex--) {
            IWorkWireMessage message = chain[styleIndex].Message;
            ApplyStyleName(message, value => data.Name = value, projectionBudget, chain[styleIndex].Record, references, ref resolvedCompletely);
            IWorkWireMessage? character = IWorkObjectIndex.TryGetMessage(message, 11, out bool malformedCharacter);
            if (malformedCharacter || message.FieldCount(11) > 1 || message.HasUnexpectedWireKind(11, IWorkWireKind.Bytes)
                || message.HasField(11) && character == null) {
                references.Declarations.Record(chain[styleIndex].Record, "11", message.FieldCount(11));
                resolvedCompletely = false;
            }
            if (character != null) OverlayText(character, data, projectionBudget, chain[styleIndex].Record, references, ref resolvedCompletely);
        }
        IWorkTextStyle result = data.ToPublic();
        cache.Add(key, new Cached<IWorkTextStyle>(result, resolvedCompletely));
        if (!resolvedCompletely) complete = false;
        return result;
    }

    private static void ApplyStyleName(IWorkWireMessage message, Action<string> apply,
        IWorkProjectionBudget projectionBudget, IWorkArchiveRecord record,
        IWorkSourceReferenceIssueCollector references, ref bool complete) {
        IWorkWireMessage? super = IWorkObjectIndex.TryGetMessage(message, 1, out bool malformedSuper);
        if (malformedSuper || message.HasUnexpectedWireKind(1, IWorkWireKind.Bytes)
            || message.HasField(1) && super == null) {
            references.Declarations.Record(record, "1", message.FieldCount(1));
            complete = false;
            return;
        }
        if (super == null || !super.HasField(1)) return;
        if (super.FieldCount(1) != 1
            || super.HasUnexpectedWireKind(1, IWorkWireKind.Bytes)
            || !TryDecodeUtf8(super.GetBytes(1)!, projectionBudget, out string name)) {
            new StylePropertyEvidence(record, "1/", references.Declarations).Record(super, 1);
            complete = false;
        }
        else apply(name);
    }

    private static void OverlayText(IWorkWireMessage message, TextStyleData data,
        IWorkProjectionBudget projectionBudget, IWorkArchiveRecord record,
        IWorkSourceReferenceIssueCollector references, ref bool complete) {
        AssessUnmappedTypography(message, record, references.Declarations, ref complete);
        var evidence = new StylePropertyEvidence(record, "11/", references.Declarations);
        OverlayFlag(message, 1, value => data.Bold = value, evidence, ref complete);
        OverlayFlag(message, 2, value => data.Italic = value, evidence, ref complete);
        OverlayFlag(message, 11, value => data.Underline = value, evidence, ref complete);
        OverlayFlag(message, 12, value => data.Strikethrough = value, evidence, ref complete);
        float? size = message.GetFloat(3);
        if (message.FieldCount(3) > 1
            || message.HasUnexpectedWireKind(3, IWorkWireKind.Fixed32)) {
            evidence.Record(message, 3); complete = false;
        }
        else if (size.HasValue && IsFinitePositive(size.Value)) data.FontSizePoints = size.Value;
        else if (message.HasField(3)) { evidence.Record(message, 3); complete = false; }
        bool? clearFont = ReadBoolean(message, 4, evidence, ref complete);
        if (clearFont == true) {
            if (message.HasField(5)) { evidence.Record(message, 5); complete = false; }
            data.FontName = null;
        }
        else if ((!message.HasField(4) || clearFont.HasValue) && message.HasField(5)) {
            if (message.FieldCount(5) != 1
                || message.HasUnexpectedWireKind(5, IWorkWireKind.Bytes)
                || !TryDecodeUtf8(message.GetBytes(5)!, projectionBudget, out string fontName)) {
                evidence.Record(message, 5); complete = false;
            }
            else data.FontName = fontName;
        }
        bool? clearColor = ReadBoolean(message, 6, evidence, ref complete);
        if (clearColor == true) {
            if (message.HasField(7)) { evidence.Record(message, 7); complete = false; }
            data.Color = null;
        }
        else if ((!message.HasField(6) || clearColor.HasValue)
            && TryColor(message, 7, out IWorkColor? color, evidence, ref complete)) data.Color = color;
        bool? clearBackground = ReadBoolean(message, 25, evidence, ref complete);
        if (clearBackground == true) {
            if (message.HasField(26)) { evidence.Record(message, 26); complete = false; }
            data.BackgroundColor = null;
        }
        else if ((!message.HasField(25) || clearBackground.HasValue)
            && TryColor(message, 26, out IWorkColor? background, evidence, ref complete)) data.BackgroundColor = background;
    }

    private static void OverlayParagraph(IWorkWireMessage message, ParagraphStyleData data,
        IWorkArchiveRecord record, IWorkSourceReferenceIssueCollector references, ref bool complete) {
        var evidence = new StylePropertyEvidence(record, "12/", references.Declarations);
        ulong? alignment = ReadUnsigned(message, 1, evidence, ref complete);
        if (alignment.HasValue) {
            if (alignment.Value > 4) { evidence.Record(message, 1); complete = false; }
            else data.Alignment = alignment.Value switch {
                0 => IWorkTextAlignment.Left,
                1 => IWorkTextAlignment.Right,
                2 => IWorkTextAlignment.Center,
                3 => IWorkTextAlignment.Justified,
                _ => IWorkTextAlignment.Natural
            };
        }
        OverlayFinite(message, 7, value => data.FirstLineIndentPoints = value, evidence, ref complete);
        OverlayFinite(message, 11, value => data.LeftIndentPoints = value, evidence, ref complete);
        OverlayFinite(message, 19, value => data.RightIndentPoints = value, evidence, ref complete);
        OverlayFinite(message, 20, value => data.SpaceAfterPoints = value, evidence, ref complete);
        OverlayFinite(message, 21, value => data.SpaceBeforePoints = value, evidence, ref complete);
        OverlayFlag(message, 14, value => data.PageBreakBefore = value, evidence, ref complete);
        OverlayFlag(message, 9, value => data.KeepLinesTogether = value, evidence, ref complete);
        OverlayFlag(message, 10, value => data.KeepWithNext = value, evidence, ref complete);
    }


    private static bool TryColor(IWorkWireMessage message, int field, out IWorkColor? color,
        StylePropertyEvidence evidence, ref bool complete) {
        bool colorComplete = true;
        bool result = IWorkColorReader.TryRead(message, field, out color, ref colorComplete);
        if (!colorComplete) { evidence.Record(message, field); complete = false; }
        return result;
    }

    private static void OverlayFinite(IWorkWireMessage message, int field, Action<double> apply,
        StylePropertyEvidence evidence, ref bool complete) {
        float? value = message.GetFloat(field);
        if (message.FieldCount(field) > 1
            || message.HasUnexpectedWireKind(field, IWorkWireKind.Fixed32)) {
            evidence.Record(message, field); complete = false;
        } else if (value.HasValue && IsFinite(value.Value)) apply(value.Value);
        else if (message.HasField(field)) { evidence.Record(message, field); complete = false; }
    }

    private static ulong? ReadUnsigned(IWorkWireMessage message, int field,
        StylePropertyEvidence evidence, ref bool complete) {
        if (message.FieldCount(field) > 1
            || message.HasUnexpectedWireKind(field, IWorkWireKind.Varint)) {
            evidence.Record(message, field); complete = false;
            return null;
        }
        return message.GetUnsigned(field);
    }

    private static void OverlayFlag(IWorkWireMessage message, int field, Action<bool> apply,
        StylePropertyEvidence evidence, ref bool complete) {
        bool? value = ReadBoolean(message, field, evidence, ref complete);
        if (value.HasValue) apply(value.Value);
    }

    private static bool? ReadBoolean(IWorkWireMessage message, int field,
        StylePropertyEvidence evidence, ref bool complete) {
        ulong? value = ReadUnsigned(message, field, evidence, ref complete);
        if (value > 1) {
            evidence.Record(message, field); complete = false;
            return null;
        }
        return value.HasValue ? value.Value == 1 : null;
    }

    /// <summary>Identifies a selected style property without treating invalid bytes as references or content counts.</summary>
    private readonly struct StylePropertyEvidence(IWorkArchiveRecord owner, string prefix,
        IWorkSourceDeclarationIssueCollector declarations) {
        internal void Record(IWorkWireMessage message, int field) =>
            declarations.Record(owner, prefix + field.ToString(System.Globalization.CultureInfo.InvariantCulture),
                message.FieldCount(field), IWorkSourceDeclarationIssueKind.InvalidValue);
    }
}

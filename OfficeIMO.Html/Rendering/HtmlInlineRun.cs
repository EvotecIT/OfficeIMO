using AngleSharp.Dom;
using OfficeIMO.Drawing;

namespace OfficeIMO.Html;

internal sealed class HtmlInlineRun {
    internal HtmlInlineRun(
        IEnumerable<HtmlCssRunningStringAssignment> runningElementAssignments,
        HtmlRenderBoxStyle style,
        string source) {
        RunningElementAssignments = new List<HtmlCssRunningStringAssignment>(runningElementAssignments).AsReadOnly();
        RunningElementAssignment = RunningElementAssignments.FirstOrDefault();
        Text = string.Empty;
        LogicalText = string.Empty;
        Style = style;
        Source = source;
        SemanticRole = style.SemanticRole;
    }

    internal HtmlInlineRun(
        IElement runningStringElement,
        HtmlRenderBoxStyle style,
        string source) {
        RunningStringElement = runningStringElement;
        Text = string.Empty;
        LogicalText = string.Empty;
        Style = style;
        Source = source;
        SemanticRole = style.SemanticRole;
    }

    internal HtmlInlineRun(
        string text,
        HtmlRenderBoxStyle style,
        string? linkUri,
        string source,
        double paintOffsetX = 0D,
        double paintOffsetY = 0D,
        IElement? ownerElement = null,
        IElement? positionedMarkerElement = null,
        string? logicalText = null,
        bool textTransformPending = false,
        string? leaderPattern = null) {
        Text = text;
        LogicalText = logicalText ?? text;
        Style = style;
        LinkUri = linkUri;
        Source = source;
        PaintOffsetX = paintOffsetX;
        PaintOffsetY = paintOffsetY;
        OwnerElement = ownerElement;
        PositionedMarkerElement = positionedMarkerElement;
        SemanticRole = style.SemanticRole;
        TextTransformPending = textTransformPending;
        LeaderPattern = leaderPattern;
    }

    internal HtmlInlineRun(
        HtmlRenderFlowBlock atomicBlock,
        HtmlRenderBoxStyle style,
        string? linkUri,
        string source,
        double paintOffsetX = 0D,
        double paintOffsetY = 0D,
        IElement? ownerElement = null,
        bool isReplacedImage = false,
        double? atomicBaseline = null,
        bool isBlockInterruption = false, bool isFlowMarker = false) {
        AtomicBlock = atomicBlock;
        Text = string.Empty;
        LogicalText = string.Empty;
        Style = style;
        LinkUri = linkUri;
        Source = source;
        PaintOffsetX = paintOffsetX;
        PaintOffsetY = paintOffsetY;
        OwnerElement = ownerElement;
        IsReplacedImage = isReplacedImage;
        AtomicBaseline = atomicBaseline;
        IsBlockInterruption = isBlockInterruption;
        IsFlowMarker = isFlowMarker;
        SemanticRole = style.SemanticRole;
    }

    internal HtmlInlineRun(
        HtmlRenderFlowBlock floatingBlock,
        HtmlRenderBoxStyle style,
        string? linkUri,
        string source,
        string floatSide,
        string clearSide,
        IElement ownerElement) {
        FloatingBlock = floatingBlock;
        Text = string.Empty;
        LogicalText = string.Empty;
        Style = style;
        LinkUri = linkUri;
        Source = source;
        FloatSide = floatSide;
        ClearSide = clearSide;
        OwnerElement = ownerElement;
        SemanticRole = style.SemanticRole;
    }

    internal string Text { get; private set; }
    internal string LogicalText { get; private set; }
    internal HtmlRenderFlowBlock? AtomicBlock { get; }
    internal HtmlRenderFlowBlock? FloatingBlock { get; }
    internal HtmlRenderBoxStyle Style { get; }
    internal string? LinkUri { get; }
    internal string Source { get; }
    internal double PaintOffsetX { get; }
    internal double PaintOffsetY { get; }
    internal IElement? OwnerElement { get; }
    internal IElement? PositionedMarkerElement { get; }
    internal IElement? RunningStringElement { get; }
    internal HtmlCssRunningStringAssignment? RunningElementAssignment { get; }
    internal IReadOnlyList<HtmlCssRunningStringAssignment> RunningElementAssignments { get; } = Array.Empty<HtmlCssRunningStringAssignment>();
    internal bool IsReplacedImage { get; }
    internal double? AtomicBaseline { get; }
    internal bool HasEmptyInlineBlockBaseline { get; set; }
    internal bool IsBlockInterruption { get; }
    internal bool IsFlowMarker { get; }
    internal HtmlInlineEdgeBoundary? InlineEdgeBoundary { get; set; }
    internal IReadOnlyList<HtmlInlineEdgeScope> InlineEdgeScopes { get; set; } = Array.Empty<HtmlInlineEdgeScope>();
    // Inline line-box metrics also belong to wrappers without decoration edges.
    internal IReadOnlyList<HtmlRenderBoxStyle> InlineStrutStyles { get; set; } = Array.Empty<HtmlRenderBoxStyle>();
    internal double InlineClosingAdvance { get; set; }
    internal bool InlineTokenEndsRun { get; set; }
    internal bool IsEmptyInlineBox { get; set; }
    internal bool IsInlineStrutMarker { get; set; }
    // Removed zero-font source text may retain a CSS whitespace opportunity without content or advance.
    internal bool IsSuppressedWhitespaceBreak { get; set; }
    internal string SemanticRole { get; private set; }
    internal int? SemanticNodeId { get; private set; }
    internal int? SemanticFragmentOrder { get; private set; }
    internal int? LogicalTextOrder { get; private set; }
    internal HtmlRenderSemanticGroupRole? InlineSemanticGroupRole { get; private set; }
    internal string? InlineSemanticGroupKey { get; private set; }
    internal string? BookmarkAnchorText { get; private set; }
    internal string FloatSide { get; } = "none";
    internal string ClearSide { get; } = "none";
    internal bool TextTransformPending { get; private set; }
    // Suppressed text participates in casing and first-letter selection before paint.
    internal bool IsTextTransformContextOnly { get; set; }
    internal string? LeaderPattern { get; }
    internal bool IsFirstLetter { get; private set; }
    internal OfficePreparedTextShaping? PreparedShaping { get; set; }
    internal HtmlRenderBoxStyle? PreparedShapingStyle { get; set; }

    internal OfficePreparedTextShaping? GetPreparedShaping(string text, HtmlRenderBoxStyle style) =>
        PreparedShaping != null && string.Equals(text, PreparedShaping.Text, StringComparison.Ordinal)
        && PreparedShapingStyle != null && style.Font.Equals(PreparedShapingStyle.Font)
        && style.FontDescriptor == PreparedShapingStyle.FontDescriptor
        && style.TextFeatureSettings.Equals(PreparedShapingStyle.TextFeatureSettings)
        && style.LetterSpacing == PreparedShapingStyle.LetterSpacing
        && style.WordSpacing == PreparedShapingStyle.WordSpacing
        ? PreparedShaping : null;
    // A pseudo-element split retains the whole word's break decisions and logical text.
    internal HtmlRenderLayoutEngine.HyphenationToken? PreparedHyphenation { get; set; }
    internal bool EndsFirstLine { get; set; }
    internal string FirstLineHyphen { get; set; } = string.Empty;

    internal void CompleteTextTransform(string text) {
        Text = text;
        LogicalText = text;
        TextTransformPending = false;
    }

    internal void AssignSemanticNode(string role, int nodeId, string? bookmarkAnchorText = null, int? semanticFragmentOrder = null) {
        SemanticNodeId = nodeId;
        SemanticFragmentOrder = semanticFragmentOrder;
        BookmarkAnchorText = bookmarkAnchorText;
        if (!SemanticRole.StartsWith("generated-", StringComparison.Ordinal)) {
            SemanticRole = role;
        }
    }

    internal void AssignLogicalTextOrder(int order) => LogicalTextOrder = order;

    internal void AssignInlineSemanticGroup(HtmlRenderSemanticGroupRole role, string structureElementKey) {
        InlineSemanticGroupRole = role;
        InlineSemanticGroupKey = structureElementKey;
    }

    internal HtmlInlineRun CloneText(string text, string logicalText, HtmlRenderBoxStyle style, bool isFirstLetter = false) {
        var clone = new HtmlInlineRun(
            text,
            style,
            LinkUri,
            Source,
            PaintOffsetX,
            PaintOffsetY,
            OwnerElement,
            PositionedMarkerElement,
            logicalText,
            TextTransformPending,
            LeaderPattern) {
            SemanticRole = SemanticRole,
            SemanticNodeId = SemanticNodeId,
            SemanticFragmentOrder = SemanticFragmentOrder,
            LogicalTextOrder = LogicalTextOrder,
            InlineSemanticGroupRole = InlineSemanticGroupRole,
            InlineSemanticGroupKey = InlineSemanticGroupKey,
            BookmarkAnchorText = BookmarkAnchorText,
            IsTextTransformContextOnly = style.Font.Size <= 0D,
            IsFirstLetter = isFirstLetter,
            PreparedShaping = GetPreparedShaping(text, style),
            PreparedShapingStyle = PreparedShapingStyle,
            PreparedHyphenation = PreparedHyphenation,
            EndsFirstLine = EndsFirstLine,
            FirstLineHyphen = FirstLineHyphen,
            InlineEdgeScopes = InlineEdgeScopes,
            InlineStrutStyles = InlineStrutStyles,
            IsInlineStrutMarker = IsInlineStrutMarker,
            IsSuppressedWhitespaceBreak = IsSuppressedWhitespaceBreak,
            InlineClosingAdvance = InlineClosingAdvance,
            InlineTokenEndsRun = InlineTokenEndsRun
        };
        return clone;
    }
}

using OfficeIMO.Drawing;

namespace OfficeIMO.Html;

internal sealed class HtmlRenderBoxStyle {
    internal bool HasIntrinsicWidths => IntrinsicWidth.HasValue || IntrinsicMinWidth.HasValue || IntrinsicMaxWidth.HasValue;
    internal string Display = "block";
    internal bool DisplayWasSpecified;
    internal bool PaintVisible = true;
    internal string BorderCollapse = "separate";
    internal double BorderSpacingX;
    internal double BorderSpacingY;
    internal double? ContainerUnitWidth;
    internal double? ContainerUnitHeight;
    internal double MarginTop;
    internal double MarginRight;
    internal double MarginBottom;
    internal double MarginLeft;
    internal double PaddingTop;
    internal double PaddingRight;
    internal double PaddingBottom;
    internal double PaddingLeft;
    internal double BorderWidth {
        get => Borders.MaximumLayoutWidth;
        set => Borders = Borders.WithUniformWidth(value);
    }
    internal OfficeColor BorderColor {
        get => Borders.Top.Color;
        set => Borders = Borders.WithUniformColor(value);
    }
    internal string BorderStyle {
        get => Borders.IsUniform ? Borders.Top.Style : "mixed";
        set => Borders = Borders.WithUniformStyle(value);
    }
    /// <summary>Used half-width winning edges for collapsed tables; authored borders remain available for conflict resolution.</summary>
    internal HtmlRenderBorderInsets? TableCollapsedBorderInsets;
    internal double BorderTopWidth => BorderInsets.Top;
    internal double BorderRightWidth => BorderInsets.Right;
    internal double BorderBottomWidth => BorderInsets.Bottom;
    internal double BorderLeftWidth => BorderInsets.Left;
    internal bool HasBorderLayout => BorderInsets.Horizontal > 0D || BorderInsets.Vertical > 0D;
    internal HtmlRenderBorderInsets BorderInsets => TableCollapsedBorderInsets ?? Borders.Insets;
    internal OfficeFontInfo Font;
    internal OfficeFontFaceDescriptor FontDescriptor = OfficeFontFaceDescriptor.Regular;
    internal OfficeColor DecorationColor = OfficeColor.Black;
    internal OfficeColor Color = OfficeColor.Black;
    internal OfficeTextAlignment Alignment;
    internal double LineHeight;
    internal bool UsesNaturalLineHeight;
    // Keep the used alignment height separate from ordinary text paint and the
    // computed CSS value. Cloned styles retain the originating length's units.
    internal double? InlineAlignmentLineHeight;
    internal double? ExplicitWidth;
    internal bool ExplicitWidthUsesPercentage;
    internal double? ExplicitHeight;
    // Non-replaced inline/contents boxes carry their enclosing basis without
    // turning it into an authored dimension used by decoration or box layout.
    internal bool HasForwardedContainingHeight;
    internal double? ForwardedContainingHeight;
    internal double? DefiniteContentHeight => HasForwardedContainingHeight
        ? ForwardedContainingHeight
        : ExplicitHeight.HasValue
            ? BorderBox ? Math.Max(0D, ExplicitHeight.Value - VerticalInsets) : ExplicitHeight.Value
            : null;
    // Flex/grid stretching supplies a used height without changing the authored auto-height contract.
    internal bool AutoHeightLayoutStretch;
    internal bool HasPageMarginIntrinsicHeights => PageMarginIntrinsicHeight
        || PageMarginIntrinsicMinHeight || PageMarginIntrinsicMaxHeight;
    internal bool BorderBox;
    internal bool PreserveWhitespace;
    internal bool PreventTextWrapping;
    internal string Language = string.Empty;
    // Resolved box applicability survives positioning/float clones that reset
    // their original boxifying fields. Synthetic styles carry no declarations.
    internal string UsedTransform => TransformAppliesToBox ? Transform : "none";
    internal string UsedIndividualScale => TransformAppliesToBox ? IndividualScale : "none";
    internal string SemanticRole = "block";

    internal double HorizontalInsets => BorderLeftWidth + BorderRightWidth + PaddingLeft + PaddingRight;
    internal double VerticalInsets => BorderTopWidth + BorderBottomWidth + PaddingTop + PaddingBottom;

    internal HtmlRenderBoxStyle Clone() {
        _shared = -1;
        return (HtmlRenderBoxStyle)MemberwiseClone();
    }

    private static bool SameValue(double left, double right) => BitConverter.DoubleToInt64Bits(left) == BitConverter.DoubleToInt64Bits(right);

    private static bool SameValue(double? left, double? right) => left.HasValue
        ? right.HasValue && SameValue(left.Value, right.Value)
        : !right.HasValue;

    private static bool SameValue(HtmlRenderBorderEdges left, HtmlRenderBorderEdges right) =>
        SameValue(left.Top, right.Top) && SameValue(left.Right, right.Right)
        && SameValue(left.Bottom, right.Bottom) && SameValue(left.Left, right.Left);

    private static bool SameValue(HtmlRenderBorderSide left, HtmlRenderBorderSide right) =>
        SameValue(left.Width, right.Width)
        && string.Equals(left.Style, right.Style, StringComparison.Ordinal)
        && left.Color.Equals(right.Color);

    private static bool SameValue(HtmlFontSizeContext? left, HtmlFontSizeContext? right) => left.HasValue
        ? right.HasValue && SameValue(left.Value.Pixels, right.Value.Pixels)
            && left.Value.IsAbsolute == right.Value.IsAbsolute
            && string.Equals(left.Value.AbsoluteKeyword, right.Value.AbsoluteKeyword, StringComparison.Ordinal)
            && left.Value.SoleGenericMonospace == right.Value.SoleGenericMonospace
        : !right.HasValue;

    // Rarely set computed values live in groups that copies share until one of them writes.
    // A bit in _shared means the group may be referenced by another style and is copied on write.
    private int _shared = -1;
    private PositionValues _position = PositionValues.Default;
    private PositionValues MutablePosition() {
        if ((_shared & 256) == 0) return _position;
        _shared &= ~256;
        return _position = _position.Copy();
    }
    internal string Position { get => _position.Position; set { if (!(string.Equals(_position.Position, value, StringComparison.Ordinal))) MutablePosition().Position = value; } }
    internal string FloatSide { get => _position.FloatSide; set { if (!(string.Equals(_position.FloatSide, value, StringComparison.Ordinal))) MutablePosition().FloatSide = value; } }
    internal string ClearSide { get => _position.ClearSide; set { if (!(string.Equals(_position.ClearSide, value, StringComparison.Ordinal))) MutablePosition().ClearSide = value; } }
    internal string Top { get => _position.Top; set { if (!(string.Equals(_position.Top, value, StringComparison.Ordinal))) MutablePosition().Top = value; } }
    internal string Right { get => _position.Right; set { if (!(string.Equals(_position.Right, value, StringComparison.Ordinal))) MutablePosition().Right = value; } }
    internal string Bottom { get => _position.Bottom; set { if (!(string.Equals(_position.Bottom, value, StringComparison.Ordinal))) MutablePosition().Bottom = value; } }
    internal string Left { get => _position.Left; set { if (!(string.Equals(_position.Left, value, StringComparison.Ordinal))) MutablePosition().Left = value; } }
    internal string ZIndex { get => _position.ZIndex; set { if (!(string.Equals(_position.ZIndex, value, StringComparison.Ordinal))) MutablePosition().ZIndex = value; } }
    internal string ContainerType { get => _position.ContainerType; set { if (!(string.Equals(_position.ContainerType, value, StringComparison.Ordinal))) MutablePosition().ContainerType = value; } }
    internal string OverflowX { get => _position.OverflowX; set { if (!(string.Equals(_position.OverflowX, value, StringComparison.Ordinal))) MutablePosition().OverflowX = value; } }
    internal string OverflowY { get => _position.OverflowY; set { if (!(string.Equals(_position.OverflowY, value, StringComparison.Ordinal))) MutablePosition().OverflowY = value; } }
    internal bool MarginTopAuto { get => _position.MarginTopAuto; set { if (!(_position.MarginTopAuto == value)) MutablePosition().MarginTopAuto = value; } }
    internal bool MarginRightAuto { get => _position.MarginRightAuto; set { if (!(_position.MarginRightAuto == value)) MutablePosition().MarginRightAuto = value; } }
    internal bool MarginBottomAuto { get => _position.MarginBottomAuto; set { if (!(_position.MarginBottomAuto == value)) MutablePosition().MarginBottomAuto = value; } }
    internal bool MarginLeftAuto { get => _position.MarginLeftAuto; set { if (!(_position.MarginLeftAuto == value)) MutablePosition().MarginLeftAuto = value; } }
    internal double HorizontalAutoMarginLeftShare { get => _position.HorizontalAutoMarginLeftShare; set { if (!(SameValue(_position.HorizontalAutoMarginLeftShare, value))) MutablePosition().HorizontalAutoMarginLeftShare = value; } }
    internal double HorizontalAutoMarginRightShare { get => _position.HorizontalAutoMarginRightShare; set { if (!(SameValue(_position.HorizontalAutoMarginRightShare, value))) MutablePosition().HorizontalAutoMarginRightShare = value; } }
    private SizingValues _sizing = SizingValues.Default;
    private SizingValues MutableSizing() {
        if ((_shared & 1024) == 0) return _sizing;
        _shared &= ~1024;
        return _sizing = _sizing.Copy();
    }
    internal bool IntrinsicWidthsResolved { get => _sizing.IntrinsicWidthsResolved; set { if (!(_sizing.IntrinsicWidthsResolved == value)) MutableSizing().IntrinsicWidthsResolved = value; } }
    internal HtmlRenderIntrinsicWidth? IntrinsicWidth { get => _sizing.IntrinsicWidth; set { if (!(_sizing.IntrinsicWidth == value)) MutableSizing().IntrinsicWidth = value; } }
    internal HtmlRenderIntrinsicWidth? IntrinsicMinWidth { get => _sizing.IntrinsicMinWidth; set { if (!(_sizing.IntrinsicMinWidth == value)) MutableSizing().IntrinsicMinWidth = value; } }
    internal HtmlRenderIntrinsicWidth? IntrinsicMaxWidth { get => _sizing.IntrinsicMaxWidth; set { if (!(_sizing.IntrinsicMaxWidth == value)) MutableSizing().IntrinsicMaxWidth = value; } }
    internal double? MinWidth { get => _sizing.MinWidth; set { if (!(SameValue(_sizing.MinWidth, value))) MutableSizing().MinWidth = value; } }
    internal double? MinWidthWithIndefiniteReference { get => _sizing.MinWidthWithIndefiniteReference; set { if (!(SameValue(_sizing.MinWidthWithIndefiniteReference, value))) MutableSizing().MinWidthWithIndefiniteReference = value; } }
    internal double? MaxWidth { get => _sizing.MaxWidth; set { if (!(SameValue(_sizing.MaxWidth, value))) MutableSizing().MaxWidth = value; } }
    internal bool MaxWidthUsesPercentage { get => _sizing.MaxWidthUsesPercentage; set { if (!(_sizing.MaxWidthUsesPercentage == value)) MutableSizing().MaxWidthUsesPercentage = value; } }
    internal double? MinHeight { get => _sizing.MinHeight; set { if (!(SameValue(_sizing.MinHeight, value))) MutableSizing().MinHeight = value; } }
    internal double? MaxHeight { get => _sizing.MaxHeight; set { if (!(SameValue(_sizing.MaxHeight, value))) MutableSizing().MaxHeight = value; } }
    internal bool PageMarginIntrinsicHeight { get => _sizing.PageMarginIntrinsicHeight; set { if (!(_sizing.PageMarginIntrinsicHeight == value)) MutableSizing().PageMarginIntrinsicHeight = value; } }
    internal bool PageMarginIntrinsicMinHeight { get => _sizing.PageMarginIntrinsicMinHeight; set { if (!(_sizing.PageMarginIntrinsicMinHeight == value)) MutableSizing().PageMarginIntrinsicMinHeight = value; } }
    internal bool PageMarginIntrinsicMaxHeight { get => _sizing.PageMarginIntrinsicMaxHeight; set { if (!(_sizing.PageMarginIntrinsicMaxHeight == value)) MutableSizing().PageMarginIntrinsicMaxHeight = value; } }
    internal bool IntrinsicColumnHeightAllocation { get => _sizing.IntrinsicColumnHeightAllocation; set { if (!(_sizing.IntrinsicColumnHeightAllocation == value)) MutableSizing().IntrinsicColumnHeightAllocation = value; } }
    private TextValues _text = TextValues.Default;
    private TextValues MutableText() {
        if ((_shared & 2048) == 0) return _text;
        _shared &= ~2048;
        return _text = _text.Copy();
    }
    internal HtmlFontSizeContext? FontSizeContext { get => _text.FontSizeContext; set { if (!(SameValue(_text.FontSizeContext, value))) MutableText().FontSizeContext = value; } }
    internal OfficeTextDecorationStyle UnderlineStyle { get => _text.UnderlineStyle; set { if (!(_text.UnderlineStyle == value)) MutableText().UnderlineStyle = value; } }
    internal OfficeTextDecorationStyle StrikethroughStyle { get => _text.StrikethroughStyle; set { if (!(_text.StrikethroughStyle == value)) MutableText().StrikethroughStyle = value; } }
    internal OfficeTextBaseline Baseline { get => _text.Baseline; set { if (!(_text.Baseline == value)) MutableText().Baseline = value; } }
    internal int BaselineLevel { get => _text.BaselineLevel; set { if (!(_text.BaselineLevel == value)) MutableText().BaselineLevel = value; } }
    internal double BaselineScale { get => _text.BaselineScale; set { if (!(SameValue(_text.BaselineScale, value))) MutableText().BaselineScale = value; } }
    internal double BaselineOffset { get => _text.BaselineOffset; set { if (!(SameValue(_text.BaselineOffset, value))) MutableText().BaselineOffset = value; } }
    internal HtmlRenderTextIndent? TextIndent { get => _text.TextIndent; set { if (!(ReferenceEquals(_text.TextIndent, value))) MutableText().TextIndent = value; } }
    internal double CharacterAdvance { get => _text.CharacterAdvance; set { if (!(SameValue(_text.CharacterAdvance, value))) MutableText().CharacterAdvance = value; } }
    internal double LetterSpacing { get => _text.LetterSpacing; set { if (!(SameValue(_text.LetterSpacing, value))) MutableText().LetterSpacing = value; } }
    internal double WordSpacing { get => _text.WordSpacing; set { if (!(SameValue(_text.WordSpacing, value))) MutableText().WordSpacing = value; } }
    internal OfficeTextFeatureSettings TextFeatureSettings { get => _text.TextFeatureSettings; set { if (!(ReferenceEquals(_text.TextFeatureSettings, value))) MutableText().TextFeatureSettings = value; } }
    internal string Direction { get => _text.Direction; set { if (!(string.Equals(_text.Direction, value, StringComparison.Ordinal))) MutableText().Direction = value; } }
    internal string WritingMode { get => _text.WritingMode; set { if (!(string.Equals(_text.WritingMode, value, StringComparison.Ordinal))) MutableText().WritingMode = value; } }
    private TableValues _table = TableValues.Default;
    private TableValues MutableTable() {
        if ((_shared & 512) == 0) return _table;
        _shared &= ~512;
        return _table = _table.Copy();
    }
    internal string TableVerticalAlignment { get => _table.TableVerticalAlignment; set { if (!(string.Equals(_table.TableVerticalAlignment, value, StringComparison.Ordinal))) MutableTable().TableVerticalAlignment = value; } }
    internal string CaptionSide { get => _table.CaptionSide; set { if (!(string.Equals(_table.CaptionSide, value, StringComparison.Ordinal))) MutableTable().CaptionSide = value; } }
    internal string TableLayout { get => _table.TableLayout; set { if (!(string.Equals(_table.TableLayout, value, StringComparison.Ordinal))) MutableTable().TableLayout = value; } }
    internal string TablePercentageHeight { get => _table.TablePercentageHeight; set { if (!(string.Equals(_table.TablePercentageHeight, value, StringComparison.Ordinal))) MutableTable().TablePercentageHeight = value; } }
    private BoxPaintValues _boxPaint = BoxPaintValues.Default;
    private BoxPaintValues MutableBoxPaint() {
        if ((_shared & 4096) == 0) return _boxPaint;
        _shared &= ~4096;
        return _boxPaint = _boxPaint.Copy();
    }
    internal HtmlRenderBorderEdges Borders { get => _boxPaint.Borders; set { if (!(SameValue(_boxPaint.Borders, value))) MutableBoxPaint().Borders = value; } }
    internal bool BorderDeclared { get => _boxPaint.BorderDeclared; set { if (!(_boxPaint.BorderDeclared == value)) MutableBoxPaint().BorderDeclared = value; } }
    internal OfficeColor? BackgroundColor { get => _boxPaint.BackgroundColor; set { if (!(Nullable.Equals(_boxPaint.BackgroundColor, value))) MutableBoxPaint().BackgroundColor = value; } }
    private TypographyValues _typography = TypographyValues.Default;
    private TypographyValues MutableTypography() {
        if ((_shared & 1) == 0) return _typography;
        _shared &= ~1;
        return _typography = _typography.Copy();
    }
    internal bool BreakSpaces { get => _typography.BreakSpaces; set { if (!(_typography.BreakSpaces == value)) MutableTypography().BreakSpaces = value; } }
    internal double TabSize { get => _typography.TabSize; set { if (!(SameValue(_typography.TabSize, value))) MutableTypography().TabSize = value; } }
    internal bool TabSizeIsLength { get => _typography.TabSizeIsLength; set { if (!(_typography.TabSizeIsLength == value)) MutableTypography().TabSizeIsLength = value; } }
    internal string TextOverflow { get => _typography.TextOverflow; set { if (!(string.Equals(_typography.TextOverflow, value, StringComparison.Ordinal))) MutableTypography().TextOverflow = value; } }
    internal int? LineClamp { get => _typography.LineClamp; set { if (!(_typography.LineClamp == value)) MutableTypography().LineClamp = value; } }
    internal string ListStyleType { get => _typography.ListStyleType; set { if (!(string.Equals(_typography.ListStyleType, value, StringComparison.Ordinal))) MutableTypography().ListStyleType = value; } }
    internal string ListStylePosition { get => _typography.ListStylePosition; set { if (!(string.Equals(_typography.ListStylePosition, value, StringComparison.Ordinal))) MutableTypography().ListStylePosition = value; } }
    internal string ListStyleImage { get => _typography.ListStyleImage; set { if (!(string.Equals(_typography.ListStyleImage, value, StringComparison.Ordinal))) MutableTypography().ListStyleImage = value; } }
    internal string FontVariant { get => _typography.FontVariant; set { if (!(string.Equals(_typography.FontVariant, value, StringComparison.Ordinal))) MutableTypography().FontVariant = value; } }
    internal string FontPalette { get => _typography.FontPalette; set { if (!(string.Equals(_typography.FontPalette, value, StringComparison.Ordinal))) MutableTypography().FontPalette = value; } }
    internal string TextTransform { get => _typography.TextTransform; set { if (!(string.Equals(_typography.TextTransform, value, StringComparison.Ordinal))) MutableTypography().TextTransform = value; } }
    internal bool ApproximateSmallCaps { get => _typography.ApproximateSmallCaps; set { if (!(_typography.ApproximateSmallCaps == value)) MutableTypography().ApproximateSmallCaps = value; } }
    internal string UnicodeBidi { get => _typography.UnicodeBidi; set { if (!(string.Equals(_typography.UnicodeBidi, value, StringComparison.Ordinal))) MutableTypography().UnicodeBidi = value; } }
    internal string TextOrientation { get => _typography.TextOrientation; set { if (!(string.Equals(_typography.TextOrientation, value, StringComparison.Ordinal))) MutableTypography().TextOrientation = value; } }
    internal string RubyPosition { get => _typography.RubyPosition; set { if (!(string.Equals(_typography.RubyPosition, value, StringComparison.Ordinal))) MutableTypography().RubyPosition = value; } }
    internal string RubyAlign { get => _typography.RubyAlign; set { if (!(string.Equals(_typography.RubyAlign, value, StringComparison.Ordinal))) MutableTypography().RubyAlign = value; } }
    internal string OverflowWrap { get => _typography.OverflowWrap; set { if (!(string.Equals(_typography.OverflowWrap, value, StringComparison.Ordinal))) MutableTypography().OverflowWrap = value; } }
    internal string WordBreak { get => _typography.WordBreak; set { if (!(string.Equals(_typography.WordBreak, value, StringComparison.Ordinal))) MutableTypography().WordBreak = value; } }
    internal string Hyphens { get => _typography.Hyphens; set { if (!(string.Equals(_typography.Hyphens, value, StringComparison.Ordinal))) MutableTypography().Hyphens = value; } }
    internal string HyphenateCharacter { get => _typography.HyphenateCharacter; set { if (!(string.Equals(_typography.HyphenateCharacter, value, StringComparison.Ordinal))) MutableTypography().HyphenateCharacter = value; } }
    internal int HyphenateMinimumWordLength { get => _typography.HyphenateMinimumWordLength; set { if (!(_typography.HyphenateMinimumWordLength == value)) MutableTypography().HyphenateMinimumWordLength = value; } }
    internal int HyphenateMinimumPrefixLength { get => _typography.HyphenateMinimumPrefixLength; set { if (!(_typography.HyphenateMinimumPrefixLength == value)) MutableTypography().HyphenateMinimumPrefixLength = value; } }
    internal int HyphenateMinimumSuffixLength { get => _typography.HyphenateMinimumSuffixLength; set { if (!(_typography.HyphenateMinimumSuffixLength == value)) MutableTypography().HyphenateMinimumSuffixLength = value; } }
    internal int? HyphenateLimitLines { get => _typography.HyphenateLimitLines; set { if (!(_typography.HyphenateLimitLines == value)) MutableTypography().HyphenateLimitLines = value; } }
    internal string HyphenateLimitLast { get => _typography.HyphenateLimitLast; set { if (!(string.Equals(_typography.HyphenateLimitLast, value, StringComparison.Ordinal))) MutableTypography().HyphenateLimitLast = value; } }
    internal double HyphenateLimitZone { get => _typography.HyphenateLimitZone; set { if (!(SameValue(_typography.HyphenateLimitZone, value))) MutableTypography().HyphenateLimitZone = value; } }
    private FlexValues _flex = FlexValues.Default;
    private FlexValues MutableFlex() {
        if ((_shared & 2) == 0) return _flex;
        _shared &= ~2;
        return _flex = _flex.Copy();
    }
    internal string FlexDirection { get => _flex.FlexDirection; set { if (!(string.Equals(_flex.FlexDirection, value, StringComparison.Ordinal))) MutableFlex().FlexDirection = value; } }
    internal string FlexWrap { get => _flex.FlexWrap; set { if (!(string.Equals(_flex.FlexWrap, value, StringComparison.Ordinal))) MutableFlex().FlexWrap = value; } }
    internal string JustifyContent { get => _flex.JustifyContent; set { if (!(string.Equals(_flex.JustifyContent, value, StringComparison.Ordinal))) MutableFlex().JustifyContent = value; } }
    internal string AlignItems { get => _flex.AlignItems; set { if (!(string.Equals(_flex.AlignItems, value, StringComparison.Ordinal))) MutableFlex().AlignItems = value; } }
    internal string AlignContent { get => _flex.AlignContent; set { if (!(string.Equals(_flex.AlignContent, value, StringComparison.Ordinal))) MutableFlex().AlignContent = value; } }
    internal string AlignSelf { get => _flex.AlignSelf; set { if (!(string.Equals(_flex.AlignSelf, value, StringComparison.Ordinal))) MutableFlex().AlignSelf = value; } }
    internal double FlexGrow { get => _flex.FlexGrow; set { if (!(SameValue(_flex.FlexGrow, value))) MutableFlex().FlexGrow = value; } }
    internal double FlexShrink { get => _flex.FlexShrink; set { if (!(SameValue(_flex.FlexShrink, value))) MutableFlex().FlexShrink = value; } }
    internal string FlexBasis { get => _flex.FlexBasis; set { if (!(string.Equals(_flex.FlexBasis, value, StringComparison.Ordinal))) MutableFlex().FlexBasis = value; } }
    internal int Order { get => _flex.Order; set { if (!(_flex.Order == value)) MutableFlex().Order = value; } }
    internal double RowGap { get => _flex.RowGap; set { if (!(SameValue(_flex.RowGap, value))) MutableFlex().RowGap = value; } }
    internal double ColumnGap { get => _flex.ColumnGap; set { if (!(SameValue(_flex.ColumnGap, value))) MutableFlex().ColumnGap = value; } }
    internal bool ColumnGapWasSpecified { get => _flex.ColumnGapWasSpecified; set { if (!(_flex.ColumnGapWasSpecified == value)) MutableFlex().ColumnGapWasSpecified = value; } }
    internal bool RowGapWasSpecified { get => _flex.RowGapWasSpecified; set { if (!(_flex.RowGapWasSpecified == value)) MutableFlex().RowGapWasSpecified = value; } }
    private GridValues _grid = GridValues.Default;
    private GridValues MutableGrid() {
        if ((_shared & 4) == 0) return _grid;
        _shared &= ~4;
        return _grid = _grid.Copy();
    }
    internal string GridTemplateColumns { get => _grid.GridTemplateColumns; set { if (!(string.Equals(_grid.GridTemplateColumns, value, StringComparison.Ordinal))) MutableGrid().GridTemplateColumns = value; } }
    internal string GridTemplateRows { get => _grid.GridTemplateRows; set { if (!(string.Equals(_grid.GridTemplateRows, value, StringComparison.Ordinal))) MutableGrid().GridTemplateRows = value; } }
    internal string GridTemplateAreas { get => _grid.GridTemplateAreas; set { if (!(string.Equals(_grid.GridTemplateAreas, value, StringComparison.Ordinal))) MutableGrid().GridTemplateAreas = value; } }
    internal string GridAutoColumns { get => _grid.GridAutoColumns; set { if (!(string.Equals(_grid.GridAutoColumns, value, StringComparison.Ordinal))) MutableGrid().GridAutoColumns = value; } }
    internal string GridAutoRows { get => _grid.GridAutoRows; set { if (!(string.Equals(_grid.GridAutoRows, value, StringComparison.Ordinal))) MutableGrid().GridAutoRows = value; } }
    internal string GridAutoFlow { get => _grid.GridAutoFlow; set { if (!(string.Equals(_grid.GridAutoFlow, value, StringComparison.Ordinal))) MutableGrid().GridAutoFlow = value; } }
    internal string GridColumnStart { get => _grid.GridColumnStart; set { if (!(string.Equals(_grid.GridColumnStart, value, StringComparison.Ordinal))) MutableGrid().GridColumnStart = value; } }
    internal string GridColumnEnd { get => _grid.GridColumnEnd; set { if (!(string.Equals(_grid.GridColumnEnd, value, StringComparison.Ordinal))) MutableGrid().GridColumnEnd = value; } }
    internal string GridRowStart { get => _grid.GridRowStart; set { if (!(string.Equals(_grid.GridRowStart, value, StringComparison.Ordinal))) MutableGrid().GridRowStart = value; } }
    internal string GridRowEnd { get => _grid.GridRowEnd; set { if (!(string.Equals(_grid.GridRowEnd, value, StringComparison.Ordinal))) MutableGrid().GridRowEnd = value; } }
    internal string JustifyItems { get => _grid.JustifyItems; set { if (!(string.Equals(_grid.JustifyItems, value, StringComparison.Ordinal))) MutableGrid().JustifyItems = value; } }
    internal string JustifySelf { get => _grid.JustifySelf; set { if (!(string.Equals(_grid.JustifySelf, value, StringComparison.Ordinal))) MutableGrid().JustifySelf = value; } }
    private ColumnsValues _columns = ColumnsValues.Default;
    private ColumnsValues MutableColumns() {
        if ((_shared & 8) == 0) return _columns;
        _shared &= ~8;
        return _columns = _columns.Copy();
    }
    internal int? ColumnCount { get => _columns.ColumnCount; set { if (!(_columns.ColumnCount == value)) MutableColumns().ColumnCount = value; } }
    internal double? ColumnWidth { get => _columns.ColumnWidth; set { if (!(SameValue(_columns.ColumnWidth, value))) MutableColumns().ColumnWidth = value; } }
    internal string ColumnFill { get => _columns.ColumnFill; set { if (!(string.Equals(_columns.ColumnFill, value, StringComparison.Ordinal))) MutableColumns().ColumnFill = value; } }
    internal string ColumnSpan { get => _columns.ColumnSpan; set { if (!(string.Equals(_columns.ColumnSpan, value, StringComparison.Ordinal))) MutableColumns().ColumnSpan = value; } }
    internal double ColumnRuleWidth { get => _columns.ColumnRuleWidth; set { if (!(SameValue(_columns.ColumnRuleWidth, value))) MutableColumns().ColumnRuleWidth = value; } }
    internal string ColumnRuleStyle { get => _columns.ColumnRuleStyle; set { if (!(string.Equals(_columns.ColumnRuleStyle, value, StringComparison.Ordinal))) MutableColumns().ColumnRuleStyle = value; } }
    internal OfficeColor? ColumnRuleColor { get => _columns.ColumnRuleColor; set { if (!(Nullable.Equals(_columns.ColumnRuleColor, value))) MutableColumns().ColumnRuleColor = value; } }
    private FragmentationValues _fragmentation = FragmentationValues.Default;
    private FragmentationValues MutableFragmentation() {
        if ((_shared & 16) == 0) return _fragmentation;
        _shared &= ~16;
        return _fragmentation = _fragmentation.Copy();
    }
    internal string StringSet { get => _fragmentation.StringSet; set { if (!(string.Equals(_fragmentation.StringSet, value, StringComparison.Ordinal))) MutableFragmentation().StringSet = value; } }
    internal string BoxDecorationBreak { get => _fragmentation.BoxDecorationBreak; set { if (!(string.Equals(_fragmentation.BoxDecorationBreak, value, StringComparison.Ordinal))) MutableFragmentation().BoxDecorationBreak = value; } }
    internal HtmlPageBreakTarget BreakBefore { get => _fragmentation.BreakBefore; set { if (!(_fragmentation.BreakBefore == value)) MutableFragmentation().BreakBefore = value; } }
    internal HtmlPageBreakTarget BreakAfter { get => _fragmentation.BreakAfter; set { if (!(_fragmentation.BreakAfter == value)) MutableFragmentation().BreakAfter = value; } }
    internal bool AvoidBreakInside { get => _fragmentation.AvoidBreakInside; set { if (!(_fragmentation.AvoidBreakInside == value)) MutableFragmentation().AvoidBreakInside = value; } }
    internal bool AvoidBreakBefore { get => _fragmentation.AvoidBreakBefore; set { if (!(_fragmentation.AvoidBreakBefore == value)) MutableFragmentation().AvoidBreakBefore = value; } }
    internal bool AvoidBreakAfter { get => _fragmentation.AvoidBreakAfter; set { if (!(_fragmentation.AvoidBreakAfter == value)) MutableFragmentation().AvoidBreakAfter = value; } }
    internal HtmlFragmentationScope AvoidBreakInsideScope { get => _fragmentation.AvoidBreakInsideScope; set { if (!(_fragmentation.AvoidBreakInsideScope == value)) MutableFragmentation().AvoidBreakInsideScope = value; } }
    internal HtmlFragmentationScope AvoidBreakBeforeScope { get => _fragmentation.AvoidBreakBeforeScope; set { if (!(_fragmentation.AvoidBreakBeforeScope == value)) MutableFragmentation().AvoidBreakBeforeScope = value; } }
    internal HtmlFragmentationScope AvoidBreakAfterScope { get => _fragmentation.AvoidBreakAfterScope; set { if (!(_fragmentation.AvoidBreakAfterScope == value)) MutableFragmentation().AvoidBreakAfterScope = value; } }
    internal int Orphans { get => _fragmentation.Orphans; set { if (!(_fragmentation.Orphans == value)) MutableFragmentation().Orphans = value; } }
    internal int Widows { get => _fragmentation.Widows; set { if (!(_fragmentation.Widows == value)) MutableFragmentation().Widows = value; } }
    internal string? PageName { get => _fragmentation.PageName; set { if (!(string.Equals(_fragmentation.PageName, value, StringComparison.Ordinal))) MutableFragmentation().PageName = value; } }
    internal HtmlRenderSemanticGroupRole? SemanticGroupRoleOverride { get => _fragmentation.SemanticGroupRoleOverride; set { if (!(_fragmentation.SemanticGroupRoleOverride == value)) MutableFragmentation().SemanticGroupRoleOverride = value; } }
    internal bool SemanticArtifact { get => _fragmentation.SemanticArtifact; set { if (!(_fragmentation.SemanticArtifact == value)) MutableFragmentation().SemanticArtifact = value; } }
    internal bool BookmarkLevelSpecified { get => _fragmentation.BookmarkLevelSpecified; set { if (!(_fragmentation.BookmarkLevelSpecified == value)) MutableFragmentation().BookmarkLevelSpecified = value; } }
    internal bool BookmarkSuppressed { get => _fragmentation.BookmarkSuppressed; set { if (!(_fragmentation.BookmarkSuppressed == value)) MutableFragmentation().BookmarkSuppressed = value; } }
    internal int? BookmarkLevel { get => _fragmentation.BookmarkLevel; set { if (!(_fragmentation.BookmarkLevel == value)) MutableFragmentation().BookmarkLevel = value; } }
    internal string? BookmarkLabel { get => _fragmentation.BookmarkLabel; set { if (!(string.Equals(_fragmentation.BookmarkLabel, value, StringComparison.Ordinal))) MutableFragmentation().BookmarkLabel = value; } }
    internal HtmlRenderBookmarkState BookmarkState { get => _fragmentation.BookmarkState; set { if (!(_fragmentation.BookmarkState == value)) MutableFragmentation().BookmarkState = value; } }
    private OutlineValues _outline = OutlineValues.Default;
    private OutlineValues MutableOutline() {
        if ((_shared & 32) == 0) return _outline;
        _shared &= ~32;
        return _outline = _outline.Copy();
    }
    internal double OutlineWidth { get => _outline.OutlineWidth; set { if (!(SameValue(_outline.OutlineWidth, value))) MutableOutline().OutlineWidth = value; } }
    internal double OutlineOffset { get => _outline.OutlineOffset; set { if (!(SameValue(_outline.OutlineOffset, value))) MutableOutline().OutlineOffset = value; } }
    internal OfficeColor OutlineColor { get => _outline.OutlineColor; set { if (!(_outline.OutlineColor.Equals(value))) MutableOutline().OutlineColor = value; } }
    internal bool OutlineColorInvert { get => _outline.OutlineColorInvert; set { if (!(_outline.OutlineColorInvert == value)) MutableOutline().OutlineColorInvert = value; } }
    internal string OutlineStyle { get => _outline.OutlineStyle; set { if (!(string.Equals(_outline.OutlineStyle, value, StringComparison.Ordinal))) MutableOutline().OutlineStyle = value; } }
    private EffectsValues _effects = EffectsValues.Default;
    private EffectsValues MutableEffects() {
        if ((_shared & 64) == 0) return _effects;
        _shared &= ~64;
        return _effects = _effects.Copy();
    }
    internal HtmlRenderBackgroundLayer? BorderImageLayer { get => _effects.BorderImageLayer; set { if (!(ReferenceEquals(_effects.BorderImageLayer, value))) MutableEffects().BorderImageLayer = value; } }
    internal string BorderRadius { get => _effects.BorderRadius; set { if (!(string.Equals(_effects.BorderRadius, value, StringComparison.Ordinal))) MutableEffects().BorderRadius = value; } }
    internal string BorderTopLeftRadius { get => _effects.BorderTopLeftRadius; set { if (!(string.Equals(_effects.BorderTopLeftRadius, value, StringComparison.Ordinal))) MutableEffects().BorderTopLeftRadius = value; } }
    internal string BorderTopRightRadius { get => _effects.BorderTopRightRadius; set { if (!(string.Equals(_effects.BorderTopRightRadius, value, StringComparison.Ordinal))) MutableEffects().BorderTopRightRadius = value; } }
    internal string BorderBottomRightRadius { get => _effects.BorderBottomRightRadius; set { if (!(string.Equals(_effects.BorderBottomRightRadius, value, StringComparison.Ordinal))) MutableEffects().BorderBottomRightRadius = value; } }
    internal string BorderBottomLeftRadius { get => _effects.BorderBottomLeftRadius; set { if (!(string.Equals(_effects.BorderBottomLeftRadius, value, StringComparison.Ordinal))) MutableEffects().BorderBottomLeftRadius = value; } }
    internal string BackgroundColorClip { get => _effects.BackgroundColorClip; set { if (!(string.Equals(_effects.BackgroundColorClip, value, StringComparison.Ordinal))) MutableEffects().BackgroundColorClip = value; } }
    internal int BackgroundImageLayerCount { get => _effects.BackgroundImageLayerCount; set { if (!(_effects.BackgroundImageLayerCount == value)) MutableEffects().BackgroundImageLayerCount = value; } }
    internal bool HasDeclaredBackgroundImage { get => _effects.HasDeclaredBackgroundImage; set { if (!(_effects.HasDeclaredBackgroundImage == value)) MutableEffects().HasDeclaredBackgroundImage = value; } }
    internal int GradientStopLimitExceededCount { get => _effects.GradientStopLimitExceededCount; set { if (!(_effects.GradientStopLimitExceededCount == value)) MutableEffects().GradientStopLimitExceededCount = value; } }
    internal IReadOnlyList<HtmlRenderBackgroundLayer> BackgroundImageLayers { get => _effects.BackgroundImageLayers; set { if (!(ReferenceEquals(_effects.BackgroundImageLayers, value))) MutableEffects().BackgroundImageLayers = value; } }
    internal OfficeTextDecorationStyle OverlineStyle { get => _effects.OverlineStyle; set { if (!(_effects.OverlineStyle == value)) MutableEffects().OverlineStyle = value; } }
    internal double? DecorationThickness { get => _effects.DecorationThickness; set { if (!(SameValue(_effects.DecorationThickness, value))) MutableEffects().DecorationThickness = value; } }
    internal double? UnderlineOffset { get => _effects.UnderlineOffset; set { if (!(SameValue(_effects.UnderlineOffset, value))) MutableEffects().UnderlineOffset = value; } }
    internal bool VectorTextDecoration { get => _effects.VectorTextDecoration; set { if (!(_effects.VectorTextDecoration == value)) MutableEffects().VectorTextDecoration = value; } }
    internal string ObjectFit { get => _effects.ObjectFit; set { if (!(string.Equals(_effects.ObjectFit, value, StringComparison.Ordinal))) MutableEffects().ObjectFit = value; } }
    internal string ObjectPosition { get => _effects.ObjectPosition; set { if (!(string.Equals(_effects.ObjectPosition, value, StringComparison.Ordinal))) MutableEffects().ObjectPosition = value; } }
    internal bool ApplyEmbeddedImageOrientation { get => _effects.ApplyEmbeddedImageOrientation; set { if (!(_effects.ApplyEmbeddedImageOrientation == value)) MutableEffects().ApplyEmbeddedImageOrientation = value; } }
    internal double? ImageResolutionDpi { get; set; }
    internal double? AspectRatio { get => _effects.AspectRatio; set { if (!(SameValue(_effects.AspectRatio, value))) MutableEffects().AspectRatio = value; } }
    internal bool AspectRatioPrefersIntrinsic { get => _effects.AspectRatioPrefersIntrinsic; set { if (!(_effects.AspectRatioPrefersIntrinsic == value)) MutableEffects().AspectRatioPrefersIntrinsic = value; } }
    internal string OverflowClipMarginBox { get => _effects.OverflowClipMarginBox; set { if (!(string.Equals(_effects.OverflowClipMarginBox, value, StringComparison.Ordinal))) MutableEffects().OverflowClipMarginBox = value; } }
    internal double OverflowClipMargin { get => _effects.OverflowClipMargin; set { if (!(SameValue(_effects.OverflowClipMargin, value))) MutableEffects().OverflowClipMargin = value; } }
    internal string Transform { get => _effects.Transform; set { if (!(string.Equals(_effects.Transform, value, StringComparison.Ordinal))) MutableEffects().Transform = value; } }
    internal string IndividualScale { get => _effects.IndividualScale; set { if (!(string.Equals(_effects.IndividualScale, value, StringComparison.Ordinal))) MutableEffects().IndividualScale = value; } }
    internal bool TransformAppliesToBox { get; set; }
    internal string TransformOrigin { get => _effects.TransformOrigin; set { if (!(string.Equals(_effects.TransformOrigin, value, StringComparison.Ordinal))) MutableEffects().TransformOrigin = value; } }
    internal string ClipPath { get => _effects.ClipPath; set { if (!(string.Equals(_effects.ClipPath, value, StringComparison.Ordinal))) MutableEffects().ClipPath = value; } }
    internal string Clip { get => _effects.Clip; set { if (!(string.Equals(_effects.Clip, value, StringComparison.Ordinal))) MutableEffects().Clip = value; } }
    internal bool OpacityWasSpecified { get => _effects.OpacityWasSpecified; set { if (!(_effects.OpacityWasSpecified == value)) MutableEffects().OpacityWasSpecified = value; } }
    internal int BoxShadowLayerCount { get => _effects.BoxShadowLayerCount; set { if (!(_effects.BoxShadowLayerCount == value)) MutableEffects().BoxShadowLayerCount = value; } }
    internal IReadOnlyList<HtmlCssBoxShadow> BoxShadows { get => _effects.BoxShadows; set { if (!(ReferenceEquals(_effects.BoxShadows, value))) MutableEffects().BoxShadows = value; } }
    internal int TextShadowLayerCount { get => _effects.TextShadowLayerCount; set { if (!(_effects.TextShadowLayerCount == value)) MutableEffects().TextShadowLayerCount = value; } }
    internal IReadOnlyList<HtmlCssTextShadow> TextShadows { get => _effects.TextShadows; set { if (!(ReferenceEquals(_effects.TextShadows, value))) MutableEffects().TextShadows = value; } }
    internal double Opacity { get => _effects.Opacity; set { if (!(SameValue(_effects.Opacity, value))) MutableEffects().Opacity = value; } }
    private DiagnosticsValues _diagnostics = DiagnosticsValues.Default;
    private DiagnosticsValues MutableDiagnostics() {
        if ((_shared & 128) == 0) return _diagnostics;
        _shared &= ~128;
        return _diagnostics = _diagnostics.Copy();
    }
    internal string UnsupportedFloat { get => _diagnostics.UnsupportedFloat; set { if (!(string.Equals(_diagnostics.UnsupportedFloat, value, StringComparison.Ordinal))) MutableDiagnostics().UnsupportedFloat = value; } }
    internal string UnsupportedClear { get => _diagnostics.UnsupportedClear; set { if (!(string.Equals(_diagnostics.UnsupportedClear, value, StringComparison.Ordinal))) MutableDiagnostics().UnsupportedClear = value; } }
    internal string UnsupportedCaptionSide { get => _diagnostics.UnsupportedCaptionSide; set { if (!(string.Equals(_diagnostics.UnsupportedCaptionSide, value, StringComparison.Ordinal))) MutableDiagnostics().UnsupportedCaptionSide = value; } }
    internal string UnsupportedTableLayout { get => _diagnostics.UnsupportedTableLayout; set { if (!(string.Equals(_diagnostics.UnsupportedTableLayout, value, StringComparison.Ordinal))) MutableDiagnostics().UnsupportedTableLayout = value; } }
    internal string UnsupportedBorderCollapse { get => _diagnostics.UnsupportedBorderCollapse; set { if (!(string.Equals(_diagnostics.UnsupportedBorderCollapse, value, StringComparison.Ordinal))) MutableDiagnostics().UnsupportedBorderCollapse = value; } }
    internal string UnsupportedBorderSpacing { get => _diagnostics.UnsupportedBorderSpacing; set { if (!(string.Equals(_diagnostics.UnsupportedBorderSpacing, value, StringComparison.Ordinal))) MutableDiagnostics().UnsupportedBorderSpacing = value; } }
    internal string UnsupportedRowGap { get => _diagnostics.UnsupportedRowGap; set { if (!(string.Equals(_diagnostics.UnsupportedRowGap, value, StringComparison.Ordinal))) MutableDiagnostics().UnsupportedRowGap = value; } }
    internal string UnsupportedColumnGap { get => _diagnostics.UnsupportedColumnGap; set { if (!(string.Equals(_diagnostics.UnsupportedColumnGap, value, StringComparison.Ordinal))) MutableDiagnostics().UnsupportedColumnGap = value; } }
    internal string UnsupportedColumns { get => _diagnostics.UnsupportedColumns; set { if (!(string.Equals(_diagnostics.UnsupportedColumns, value, StringComparison.Ordinal))) MutableDiagnostics().UnsupportedColumns = value; } }
    internal string UnsupportedColumnFill { get => _diagnostics.UnsupportedColumnFill; set { if (!(string.Equals(_diagnostics.UnsupportedColumnFill, value, StringComparison.Ordinal))) MutableDiagnostics().UnsupportedColumnFill = value; } }
    internal string UnsupportedColumnSpan { get => _diagnostics.UnsupportedColumnSpan; set { if (!(string.Equals(_diagnostics.UnsupportedColumnSpan, value, StringComparison.Ordinal))) MutableDiagnostics().UnsupportedColumnSpan = value; } }
    internal string UnsupportedColumnRule { get => _diagnostics.UnsupportedColumnRule; set { if (!(string.Equals(_diagnostics.UnsupportedColumnRule, value, StringComparison.Ordinal))) MutableDiagnostics().UnsupportedColumnRule = value; } }
    internal string UnsupportedBorderPaint { get => _diagnostics.UnsupportedBorderPaint; set { if (!(string.Equals(_diagnostics.UnsupportedBorderPaint, value, StringComparison.Ordinal))) MutableDiagnostics().UnsupportedBorderPaint = value; } }
    internal int UnsupportedBackgroundImageLayerCount { get => _diagnostics.UnsupportedBackgroundImageLayerCount; set { if (!(_diagnostics.UnsupportedBackgroundImageLayerCount == value)) MutableDiagnostics().UnsupportedBackgroundImageLayerCount = value; } }
    internal string UnsupportedReplacedElementLayout { get => _diagnostics.UnsupportedReplacedElementLayout; set { if (!(string.Equals(_diagnostics.UnsupportedReplacedElementLayout, value, StringComparison.Ordinal))) MutableDiagnostics().UnsupportedReplacedElementLayout = value; } }
    internal string UnsupportedOverflowX { get => _diagnostics.UnsupportedOverflowX; set { if (!(string.Equals(_diagnostics.UnsupportedOverflowX, value, StringComparison.Ordinal))) MutableDiagnostics().UnsupportedOverflowX = value; } }
    internal string UnsupportedOverflowY { get => _diagnostics.UnsupportedOverflowY; set { if (!(string.Equals(_diagnostics.UnsupportedOverflowY, value, StringComparison.Ordinal))) MutableDiagnostics().UnsupportedOverflowY = value; } }
    internal string UnsupportedOverflowClipMargin { get => _diagnostics.UnsupportedOverflowClipMargin; set { if (!(string.Equals(_diagnostics.UnsupportedOverflowClipMargin, value, StringComparison.Ordinal))) MutableDiagnostics().UnsupportedOverflowClipMargin = value; } }
    internal string UnsupportedOpacity { get => _diagnostics.UnsupportedOpacity; set { if (!(string.Equals(_diagnostics.UnsupportedOpacity, value, StringComparison.Ordinal))) MutableDiagnostics().UnsupportedOpacity = value; } }
    internal string UnsupportedBoxShadow { get => _diagnostics.UnsupportedBoxShadow; set { if (!(string.Equals(_diagnostics.UnsupportedBoxShadow, value, StringComparison.Ordinal))) MutableDiagnostics().UnsupportedBoxShadow = value; } }
    internal string UnsupportedTextShadow { get => _diagnostics.UnsupportedTextShadow; set { if (!(string.Equals(_diagnostics.UnsupportedTextShadow, value, StringComparison.Ordinal))) MutableDiagnostics().UnsupportedTextShadow = value; } }
    internal string UnsupportedOutlinePaint { get => _diagnostics.UnsupportedOutlinePaint; set { if (!(string.Equals(_diagnostics.UnsupportedOutlinePaint, value, StringComparison.Ordinal))) MutableDiagnostics().UnsupportedOutlinePaint = value; } }
    internal string UnsupportedColumnBreaks { get => _diagnostics.UnsupportedColumnBreaks; set { if (!(string.Equals(_diagnostics.UnsupportedColumnBreaks, value, StringComparison.Ordinal))) MutableDiagnostics().UnsupportedColumnBreaks = value; } }
    internal string UnsupportedSemanticTag { get => _diagnostics.UnsupportedSemanticTag; set { if (!(string.Equals(_diagnostics.UnsupportedSemanticTag, value, StringComparison.Ordinal))) MutableDiagnostics().UnsupportedSemanticTag = value; } }
    internal string UnsupportedBookmark { get => _diagnostics.UnsupportedBookmark; set { if (!(string.Equals(_diagnostics.UnsupportedBookmark, value, StringComparison.Ordinal))) MutableDiagnostics().UnsupportedBookmark = value; } }

    private sealed class PositionValues {
        internal static readonly PositionValues Default = new PositionValues();
        internal string Position = "static";
        internal string FloatSide = "none";
        internal string ClearSide = "none";
        internal string Top = "auto";
        internal string Right = "auto";
        internal string Bottom = "auto";
        internal string Left = "auto";
        internal string ZIndex = "auto";
        internal string ContainerType = "normal";
        internal string OverflowX = "visible";
        internal string OverflowY = "visible";
        internal bool MarginTopAuto;
        internal bool MarginRightAuto;
        internal bool MarginBottomAuto;
        internal bool MarginLeftAuto;
        // Keep computed auto-margin space separate from float-avoidance offsets
        // so table measurement can replace that space after resolving its used width.
        internal double HorizontalAutoMarginLeftShare;
        internal double HorizontalAutoMarginRightShare;
        internal PositionValues Copy() => (PositionValues)MemberwiseClone();
    }

    private sealed class SizingValues {
        internal static readonly SizingValues Default = new SizingValues();
        internal bool IntrinsicWidthsResolved;
        internal HtmlRenderIntrinsicWidth? IntrinsicWidth;
        internal HtmlRenderIntrinsicWidth? IntrinsicMinWidth;
        internal HtmlRenderIntrinsicWidth? IntrinsicMaxWidth;
        internal double? MinWidth;
        internal double? MinWidthWithIndefiniteReference;
        internal double? MaxWidth;
        internal bool MaxWidthUsesPercentage;
        internal double? MinHeight;
        internal double? MaxHeight;
        // Horizontal page-margin block-axis keywords use the laid-out content height.
        // Ordinary element sizing retains its separate intrinsic-height contract.
        internal bool PageMarginIntrinsicHeight;
        internal bool PageMarginIntrinsicMinHeight;
        internal bool PageMarginIntrinsicMaxHeight;
        // A flex/grid allocation equal to the intrinsic column height still needs
        // its definite percentage basis, but may fragment at physical page bands.
        internal bool IntrinsicColumnHeightAllocation;
        internal SizingValues Copy() => (SizingValues)MemberwiseClone();
    }

    private sealed class TextValues {
        internal static readonly TextValues Default = new TextValues();
        internal HtmlFontSizeContext? FontSizeContext;
        internal OfficeTextDecorationStyle UnderlineStyle;
        internal OfficeTextDecorationStyle StrikethroughStyle;
        internal OfficeTextBaseline Baseline;
        internal int BaselineLevel;
        internal double BaselineScale = 1D;
        internal double BaselineOffset;
        internal HtmlRenderTextIndent? TextIndent;
        internal double CharacterAdvance = double.NaN;
        internal double LetterSpacing;
        internal double WordSpacing;
        internal OfficeTextFeatureSettings TextFeatureSettings = OfficeTextFeatureSettings.Default;
        internal string Direction = "ltr";
        internal string WritingMode = "horizontal-tb";
        internal TextValues Copy() => (TextValues)MemberwiseClone();
    }

    private sealed class TableValues {
        internal static readonly TableValues Default = new TableValues();
        internal string TableVerticalAlignment = "baseline";
        internal string CaptionSide = "top";
        internal string TableLayout = "auto";
        // Table tracks resolve percentages after the grid's content minimum is known.
        internal string TablePercentageHeight = string.Empty;
        internal TableValues Copy() => (TableValues)MemberwiseClone();
    }

    private sealed class BoxPaintValues {
        internal static readonly BoxPaintValues Default = new BoxPaintValues();
        internal HtmlRenderBorderEdges Borders = HtmlRenderBorderEdges.Uniform(0D, "none", OfficeColor.Black);
        internal bool BorderDeclared;
        internal OfficeColor? BackgroundColor;
        internal BoxPaintValues Copy() => (BoxPaintValues)MemberwiseClone();
    }

    private sealed class TypographyValues {
        internal static readonly TypographyValues Default = new TypographyValues();
        internal bool BreakSpaces;
        internal double TabSize = 8D;
        internal bool TabSizeIsLength;
        internal string TextOverflow = "clip";
        internal int? LineClamp;
        internal string ListStyleType = string.Empty;
        internal string ListStylePosition = "outside";
        internal string ListStyleImage = "none";
        internal string FontVariant = "normal";
        internal string FontPalette = "normal";
        internal string TextTransform = "none";
        internal bool ApproximateSmallCaps;
        internal string UnicodeBidi = "normal";
        internal string TextOrientation = "mixed";
        internal string RubyPosition = "over";
        internal string RubyAlign = "space-around";
        internal string OverflowWrap = "normal";
        internal string WordBreak = "normal";
        internal string Hyphens = "manual";
        internal string HyphenateCharacter = "-";
        internal int HyphenateMinimumWordLength = 5;
        internal int HyphenateMinimumPrefixLength = 2;
        internal int HyphenateMinimumSuffixLength = 2;
        internal int? HyphenateLimitLines;
        internal string HyphenateLimitLast = "none";
        internal double HyphenateLimitZone;
        internal TypographyValues Copy() => (TypographyValues)MemberwiseClone();
    }

    private sealed class FlexValues {
        internal static readonly FlexValues Default = new FlexValues();
        internal string FlexDirection = "row";
        internal string FlexWrap = "nowrap";
        internal string JustifyContent = "normal";
        internal string AlignItems = "normal";
        internal string AlignContent = "normal";
        internal string AlignSelf = "auto";
        internal double FlexGrow;
        internal double FlexShrink = 1D;
        internal string FlexBasis = "auto";
        internal int Order;
        internal double RowGap;
        internal double ColumnGap;
        internal bool ColumnGapWasSpecified;
        internal bool RowGapWasSpecified;
        internal FlexValues Copy() => (FlexValues)MemberwiseClone();
    }

    private sealed class GridValues {
        internal static readonly GridValues Default = new GridValues();
        internal string GridTemplateColumns = "none";
        internal string GridTemplateRows = "none";
        internal string GridTemplateAreas = "none";
        internal string GridAutoColumns = "auto";
        internal string GridAutoRows = "auto";
        internal string GridAutoFlow = "row";
        internal string GridColumnStart = "auto";
        internal string GridColumnEnd = "auto";
        internal string GridRowStart = "auto";
        internal string GridRowEnd = "auto";
        internal string JustifyItems = "normal";
        internal string JustifySelf = "auto";
        internal GridValues Copy() => (GridValues)MemberwiseClone();
    }

    private sealed class ColumnsValues {
        internal static readonly ColumnsValues Default = new ColumnsValues();
        internal int? ColumnCount;
        internal double? ColumnWidth;
        internal string ColumnFill = "balance";
        internal string ColumnSpan = "none";
        internal double ColumnRuleWidth = 3D;
        internal string ColumnRuleStyle = "none";
        internal OfficeColor? ColumnRuleColor;
        internal ColumnsValues Copy() => (ColumnsValues)MemberwiseClone();
    }

    private sealed class FragmentationValues {
        internal static readonly FragmentationValues Default = new FragmentationValues();
        internal string StringSet = string.Empty;
        internal string BoxDecorationBreak = "slice";
        internal HtmlPageBreakTarget BreakBefore;
        internal HtmlPageBreakTarget BreakAfter;
        internal bool AvoidBreakInside;
        internal bool AvoidBreakBefore;
        internal bool AvoidBreakAfter;
        internal HtmlFragmentationScope AvoidBreakInsideScope;
        internal HtmlFragmentationScope AvoidBreakBeforeScope;
        internal HtmlFragmentationScope AvoidBreakAfterScope;
        internal int Orphans = 2;
        internal int Widows = 2;
        internal string? PageName;
        internal HtmlRenderSemanticGroupRole? SemanticGroupRoleOverride;
        internal bool SemanticArtifact;
        internal bool BookmarkLevelSpecified;
        internal bool BookmarkSuppressed;
        internal int? BookmarkLevel;
        internal string? BookmarkLabel;
        internal HtmlRenderBookmarkState BookmarkState;
        internal FragmentationValues Copy() => (FragmentationValues)MemberwiseClone();
    }

    private sealed class OutlineValues {
        internal static readonly OutlineValues Default = new OutlineValues();
        internal double OutlineWidth;
        internal double OutlineOffset;
        internal OfficeColor OutlineColor = OfficeColor.Black;
        internal bool OutlineColorInvert;
        internal string OutlineStyle = "none";
        internal OutlineValues Copy() => (OutlineValues)MemberwiseClone();
    }

    private sealed class EffectsValues {
        internal static readonly EffectsValues Default = new EffectsValues();
        internal HtmlRenderBackgroundLayer? BorderImageLayer;
        internal string BorderRadius = "0";
        internal string BorderTopLeftRadius = string.Empty;
        internal string BorderTopRightRadius = string.Empty;
        internal string BorderBottomRightRadius = string.Empty;
        internal string BorderBottomLeftRadius = string.Empty;
        internal string BackgroundColorClip = "border-box";
        internal int BackgroundImageLayerCount;
        internal bool HasDeclaredBackgroundImage;
        internal int GradientStopLimitExceededCount;
        internal IReadOnlyList<HtmlRenderBackgroundLayer> BackgroundImageLayers = Array.Empty<HtmlRenderBackgroundLayer>();
        internal OfficeTextDecorationStyle OverlineStyle;
        internal double? DecorationThickness;
        internal double? UnderlineOffset;
        internal bool VectorTextDecoration;
        internal string ObjectFit = "fill";
        internal string ObjectPosition = "50% 50%";
        internal bool ApplyEmbeddedImageOrientation = true;
        internal double? AspectRatio;
        internal bool AspectRatioPrefersIntrinsic = true;
        internal string OverflowClipMarginBox = "padding-box";
        internal double OverflowClipMargin;
        internal string Transform = "none";
        internal string IndividualScale = "none";
        internal string TransformOrigin = "50% 50%";
        internal string ClipPath = "none";
        internal string Clip = "auto";
        internal bool OpacityWasSpecified;
        internal int BoxShadowLayerCount;
        internal IReadOnlyList<HtmlCssBoxShadow> BoxShadows = Array.Empty<HtmlCssBoxShadow>();
        internal int TextShadowLayerCount;
        internal IReadOnlyList<HtmlCssTextShadow> TextShadows = Array.Empty<HtmlCssTextShadow>();
        internal double Opacity = 1D;
        internal EffectsValues Copy() => (EffectsValues)MemberwiseClone();
    }

    private sealed class DiagnosticsValues {
        internal static readonly DiagnosticsValues Default = new DiagnosticsValues();
        internal string UnsupportedFloat = string.Empty;
        internal string UnsupportedClear = string.Empty;
        internal string UnsupportedCaptionSide = string.Empty;
        internal string UnsupportedTableLayout = string.Empty;
        internal string UnsupportedBorderCollapse = string.Empty;
        internal string UnsupportedBorderSpacing = string.Empty;
        internal string UnsupportedRowGap = string.Empty;
        internal string UnsupportedColumnGap = string.Empty;
        internal string UnsupportedColumns = string.Empty;
        internal string UnsupportedColumnFill = string.Empty;
        internal string UnsupportedColumnSpan = string.Empty;
        internal string UnsupportedColumnRule = string.Empty;
        internal string UnsupportedBorderPaint = string.Empty;
        internal int UnsupportedBackgroundImageLayerCount;
        internal string UnsupportedReplacedElementLayout = string.Empty;
        internal string UnsupportedOverflowX = string.Empty;
        internal string UnsupportedOverflowY = string.Empty;
        internal string UnsupportedOverflowClipMargin = string.Empty;
        internal string UnsupportedOpacity = string.Empty;
        internal string UnsupportedBoxShadow = string.Empty;
        internal string UnsupportedTextShadow = string.Empty;
        internal string UnsupportedOutlinePaint = string.Empty;
        internal string UnsupportedColumnBreaks = string.Empty;
        internal string UnsupportedSemanticTag = string.Empty;
        internal string UnsupportedBookmark = string.Empty;
        internal DiagnosticsValues Copy() => (DiagnosticsValues)MemberwiseClone();
    }
}

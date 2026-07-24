using System;
using OfficeIMO.Drawing;

namespace OfficeIMO.Excel {
    /// <summary>
    /// Format-neutral visual snapshot for a worksheet range.
    /// </summary>
    public sealed class ExcelRangeVisualSnapshot {
        internal ExcelRangeVisualSnapshot(
            string sheetName,
            string range,
            int firstRow,
            int firstColumn,
            int lastRow,
            int lastColumn,
            IReadOnlyList<ExcelVisualColumn> columns,
            IReadOnlyList<ExcelVisualRow> rows,
            IReadOnlyList<ExcelVisualCell> cells,
            IReadOnlyList<ExcelVisualConditionalDataBar> conditionalDataBars,
            IReadOnlyList<ExcelVisualConditionalIcon> conditionalIcons,
            IReadOnlyList<ExcelVisualCommentIndicator> commentIndicators,
            IReadOnlyList<ExcelVisualCommentBody> commentBodies,
            IReadOnlyList<ExcelVisualSparkline> sparklines,
            IReadOnlyList<ExcelVisualDrawingLayer> drawingLayers,
            IReadOnlyList<ExcelVisualDrawingObject> drawingObjects,
            IReadOnlyList<ExcelVisualImage> images,
            IReadOnlyList<ExcelVisualChart> charts,
            IReadOnlyList<OfficeImageExportDiagnostic> diagnostics) {
            SheetName = sheetName;
            Range = range;
            FirstRow = firstRow;
            FirstColumn = firstColumn;
            LastRow = lastRow;
            LastColumn = lastColumn;
            Columns = columns;
            Rows = rows;
            Cells = cells;
            ConditionalDataBars = conditionalDataBars;
            ConditionalIcons = conditionalIcons;
            CommentIndicators = commentIndicators;
            CommentBodies = commentBodies;
            Sparklines = sparklines;
            DrawingLayers = drawingLayers;
            DrawingObjects = drawingObjects;
            Images = images;
            Charts = charts;
            Diagnostics = diagnostics;
        }

        /// <summary>Worksheet name.</summary>
        public string SheetName { get; }

        /// <summary>A1 range represented by this snapshot.</summary>
        public string Range { get; }

        /// <summary>First source row.</summary>
        public int FirstRow { get; }

        /// <summary>First source column.</summary>
        public int FirstColumn { get; }

        /// <summary>Last source row.</summary>
        public int LastRow { get; }

        /// <summary>Last source column.</summary>
        public int LastColumn { get; }

        /// <summary>Visible columns included in the snapshot.</summary>
        public IReadOnlyList<ExcelVisualColumn> Columns { get; }

        /// <summary>Visible rows included in the snapshot.</summary>
        public IReadOnlyList<ExcelVisualRow> Rows { get; }

        /// <summary>Cell visuals included in the snapshot.</summary>
        public IReadOnlyList<ExcelVisualCell> Cells { get; }

        /// <summary>Conditional-formatting data bars included in the snapshot.</summary>
        public IReadOnlyList<ExcelVisualConditionalDataBar> ConditionalDataBars { get; }

        /// <summary>Conditional-formatting icons included in the snapshot.</summary>
        public IReadOnlyList<ExcelVisualConditionalIcon> ConditionalIcons { get; }

        /// <summary>Cell comment indicators included in the snapshot.</summary>
        public IReadOnlyList<ExcelVisualCommentIndicator> CommentIndicators { get; }

        /// <summary>Visible comment body callouts included in the snapshot.</summary>
        public IReadOnlyList<ExcelVisualCommentBody> CommentBodies { get; }

        /// <summary>Worksheet sparklines included in the snapshot.</summary>
        public IReadOnlyList<ExcelVisualSparkline> Sparklines { get; }

        /// <summary>Worksheet drawing-layer overlays in source paint order.</summary>
        public IReadOnlyList<ExcelVisualDrawingLayer> DrawingLayers { get; }

        /// <summary>Worksheet drawing objects included in the snapshot.</summary>
        public IReadOnlyList<ExcelVisualDrawingObject> DrawingObjects { get; }

        /// <summary>Worksheet images included in the snapshot.</summary>
        public IReadOnlyList<ExcelVisualImage> Images { get; }

        /// <summary>Worksheet charts included in the snapshot.</summary>
        public IReadOnlyList<ExcelVisualChart> Charts { get; }

        /// <summary>Snapshot diagnostics.</summary>
        public IReadOnlyList<OfficeImageExportDiagnostic> Diagnostics { get; }

        /// <summary>Total snapshot width in CSS pixels before export scaling.</summary>
        public double Width => Columns.Sum(column => column.Width);

        /// <summary>Total snapshot height in CSS pixels before export scaling.</summary>
        public double Height => Rows.Sum(row => row.Height);
    }

    /// <summary>
    /// Visual column metadata in an Excel range snapshot.
    /// </summary>
    public sealed class ExcelVisualColumn {
        internal ExcelVisualColumn(int index, double x, double width) {
            Index = index;
            X = x;
            Width = width;
        }

        /// <summary>One-based source column index.</summary>
        public int Index { get; }

        /// <summary>X position in CSS pixels.</summary>
        public double X { get; }

        /// <summary>Column width in CSS pixels.</summary>
        public double Width { get; }
    }

    /// <summary>
    /// Visual row metadata in an Excel range snapshot.
    /// </summary>
    public sealed class ExcelVisualRow {
        internal ExcelVisualRow(int index, double y, double height) {
            Index = index;
            Y = y;
            Height = height;
        }

        /// <summary>One-based source row index.</summary>
        public int Index { get; }

        /// <summary>Y position in CSS pixels.</summary>
        public double Y { get; }

        /// <summary>Row height in CSS pixels.</summary>
        public double Height { get; }
    }

    /// <summary>
    /// Visual cell metadata in an Excel range snapshot.
    /// </summary>
    public sealed class ExcelVisualCell {
        internal ExcelVisualCell(
            int row,
            int column,
            double x,
            double y,
            double width,
            double height,
            string text,
            ExcelCellStyleSnapshot style,
            bool coveredByMerge,
            ExcelHyperlinkSnapshot? hyperlink = null,
            IReadOnlyList<ExcelVisualTextRun>? richTextRuns = null,
            ExcelVisualCellValueKind valueKind = ExcelVisualCellValueKind.Text) {
            Row = row;
            Column = column;
            X = x;
            Y = y;
            Width = width;
            Height = height;
            Text = text;
            Style = style;
            CoveredByMerge = coveredByMerge;
            Hyperlink = hyperlink;
            RichTextRuns = richTextRuns ?? Array.Empty<ExcelVisualTextRun>();
            ValueKind = valueKind;
        }

        /// <summary>One-based source row.</summary>
        public int Row { get; }

        /// <summary>One-based source column.</summary>
        public int Column { get; }

        /// <summary>X position in CSS pixels.</summary>
        public double X { get; }

        /// <summary>Y position in CSS pixels.</summary>
        public double Y { get; }

        /// <summary>Cell width in CSS pixels.</summary>
        public double Width { get; }

        /// <summary>Cell height in CSS pixels.</summary>
        public double Height { get; }

        /// <summary>Formatted display text.</summary>
        public string Text { get; }

        /// <summary>Resolved cell style.</summary>
        public ExcelCellStyleSnapshot Style { get; }

        /// <summary>Whether this cell is hidden by another merged cell in the snapshot.</summary>
        public bool CoveredByMerge { get; }

        /// <summary>Hyperlink metadata attached to this cell, when available.</summary>
        public ExcelHyperlinkSnapshot? Hyperlink { get; }

        /// <summary>Rich text runs attached to this cell, when available.</summary>
        public IReadOnlyList<ExcelVisualTextRun> RichTextRuns { get; }

        /// <summary>Value kind used for Excel visual policies such as default General alignment.</summary>
        public ExcelVisualCellValueKind ValueKind { get; }
    }

    /// <summary>
    /// Kind of value represented by an Excel visual cell.
    /// </summary>
    public enum ExcelVisualCellValueKind {
        /// <summary>The cell has no value.</summary>
        Blank,
        /// <summary>The cell displays text.</summary>
        Text,
        /// <summary>The cell displays a number.</summary>
        Number,
        /// <summary>The cell displays a date or time serial.</summary>
        Date,
        /// <summary>The cell displays a Boolean value.</summary>
        Boolean,
        /// <summary>The cell displays an error value.</summary>
        Error
    }

    /// <summary>
    /// Visual rich text run metadata in an Excel cell snapshot.
    /// </summary>
    public sealed class ExcelVisualTextRun {
        internal ExcelVisualTextRun(string text, bool bold, bool italic, bool underline, bool strikethrough, string? fontColorArgb, string? fontName, double? fontSize) {
            Text = text ?? string.Empty;
            Bold = bold;
            Italic = italic;
            Underline = underline;
            Strikethrough = strikethrough;
            FontColorArgb = fontColorArgb;
            FontName = fontName;
            FontSize = fontSize;
        }

        /// <summary>Run text.</summary>
        public string Text { get; }

        /// <summary>Whether the run is bold.</summary>
        public bool Bold { get; }

        /// <summary>Whether the run is italic.</summary>
        public bool Italic { get; }

        /// <summary>Whether the run is underlined.</summary>
        public bool Underline { get; }

        /// <summary>Whether the run is struck through.</summary>
        public bool Strikethrough { get; }

        /// <summary>Run font color in ARGB hexadecimal form, when specified.</summary>
        public string? FontColorArgb { get; }

        /// <summary>Run font family, when specified.</summary>
        public string? FontName { get; }

        /// <summary>Run font size in points, when specified.</summary>
        public double? FontSize { get; }
    }

    /// <summary>
    /// Conditional-formatting data bar overlay in an Excel range snapshot.
    /// </summary>
    public sealed class ExcelVisualConditionalDataBar {
        internal ExcelVisualConditionalDataBar(int row, int column, double x, double y, double width, double height, string colorArgb, double startRatio, double ratio, bool showValue) {
            Row = row;
            Column = column;
            X = x;
            Y = y;
            Width = width;
            Height = height;
            ColorArgb = colorArgb ?? string.Empty;
            StartRatio = startRatio;
            Ratio = ratio;
            ShowValue = showValue;
        }

        /// <summary>One-based source row.</summary>
        public int Row { get; }

        /// <summary>One-based source column.</summary>
        public int Column { get; }

        /// <summary>X position in CSS pixels.</summary>
        public double X { get; }

        /// <summary>Y position in CSS pixels.</summary>
        public double Y { get; }

        /// <summary>Cell width in CSS pixels.</summary>
        public double Width { get; }

        /// <summary>Cell height in CSS pixels.</summary>
        public double Height { get; }

        /// <summary>Bar fill color in ARGB hexadecimal form.</summary>
        public string ColorArgb { get; }

        /// <summary>Start position inside the cell, where 0 is the left edge and 1 is the right edge.</summary>
        public double StartRatio { get; }

        /// <summary>Bar width ratio inside the cell.</summary>
        public double Ratio { get; }

        /// <summary>Whether the underlying cell value should be drawn over the bar.</summary>
        public bool ShowValue { get; }
    }

    /// <summary>
    /// Conditional-formatting icon overlay in an Excel range snapshot.
    /// </summary>
    public sealed class ExcelVisualConditionalIcon {
        internal ExcelVisualConditionalIcon(int row, int column, double x, double y, double width, double height, ExcelConditionalIconKind kind, bool showValue) {
            Row = row;
            Column = column;
            X = x;
            Y = y;
            Width = width;
            Height = height;
            Kind = kind;
            ShowValue = showValue;
        }

        /// <summary>One-based source row.</summary>
        public int Row { get; }

        /// <summary>One-based source column.</summary>
        public int Column { get; }

        /// <summary>X position in CSS pixels.</summary>
        public double X { get; }

        /// <summary>Y position in CSS pixels.</summary>
        public double Y { get; }

        /// <summary>Cell width in CSS pixels.</summary>
        public double Width { get; }

        /// <summary>Cell height in CSS pixels.</summary>
        public double Height { get; }

        /// <summary>Rendered icon kind.</summary>
        public ExcelConditionalIconKind Kind { get; }

        /// <summary>Whether the formatted cell value should be rendered beside the icon.</summary>
        public bool ShowValue { get; }
    }

    /// <summary>
    /// Dependency-free icon shapes used for conditional-formatting icon sets.
    /// </summary>
    public enum ExcelConditionalIconKind {
        /// <summary>Green upward arrow.</summary>
        GreenUpArrow,
        /// <summary>Yellow upward arrow.</summary>
        YellowUpArrow,
        /// <summary>Yellow sideways arrow.</summary>
        YellowSideArrow,
        /// <summary>Yellow downward arrow.</summary>
        YellowDownArrow,
        /// <summary>Red downward arrow.</summary>
        RedDownArrow,
        /// <summary>Green check mark.</summary>
        GreenCheck,
        /// <summary>Yellow exclamation mark.</summary>
        YellowExclamation,
        /// <summary>Red cross.</summary>
        RedCross,
        /// <summary>Green circle.</summary>
        GreenCircle,
        /// <summary>Light green circle.</summary>
        LightGreenCircle,
        /// <summary>Yellow circle.</summary>
        YellowCircle,
        /// <summary>Orange circle.</summary>
        OrangeCircle,
        /// <summary>Red circle.</summary>
        RedCircle,
        /// <summary>One filled rating bar.</summary>
        RatingOne,
        /// <summary>Two filled rating bars.</summary>
        RatingTwo,
        /// <summary>Three filled rating bars.</summary>
        RatingThree,
        /// <summary>Four filled rating bars.</summary>
        RatingFour,
        /// <summary>Five filled rating bars.</summary>
        RatingFive,
        /// <summary>Empty quarter-pie indicator.</summary>
        QuarterEmpty,
        /// <summary>One-quarter filled pie indicator.</summary>
        QuarterOne,
        /// <summary>Half-filled pie indicator.</summary>
        QuarterTwo,
        /// <summary>Three-quarter filled pie indicator.</summary>
        QuarterThree,
        /// <summary>Fully filled pie indicator.</summary>
        QuarterFull,
        /// <summary>Green flag indicator.</summary>
        GreenFlag,
        /// <summary>Yellow flag indicator.</summary>
        YellowFlag,
        /// <summary>Red flag indicator.</summary>
        RedFlag
    }

    /// <summary>
    /// Visual indicator for an Excel cell comment or threaded comment.
    /// </summary>
    public sealed class ExcelVisualCommentIndicator {
        internal ExcelVisualCommentIndicator(int row, int column, double x, double y, double width, double height, bool threaded, string source) {
            Row = row;
            Column = column;
            X = x;
            Y = y;
            Width = width;
            Height = height;
            Threaded = threaded;
            Source = source ?? string.Empty;
        }

        /// <summary>One-based source row.</summary>
        public int Row { get; }

        /// <summary>One-based source column.</summary>
        public int Column { get; }

        /// <summary>X position in CSS pixels.</summary>
        public double X { get; }

        /// <summary>Y position in CSS pixels.</summary>
        public double Y { get; }

        /// <summary>Cell width in CSS pixels.</summary>
        public double Width { get; }

        /// <summary>Cell height in CSS pixels.</summary>
        public double Height { get; }

        /// <summary>Whether the indicator represents a threaded comment.</summary>
        public bool Threaded { get; }

        /// <summary>Source reference used by export diagnostics.</summary>
        public string Source { get; }
    }

    /// <summary>
    /// Visual callout for an Excel cell comment or threaded comment body.
    /// </summary>
    public sealed class ExcelVisualCommentBody {
        internal ExcelVisualCommentBody(
            int row,
            int column,
            double x,
            double y,
            double width,
            double height,
            double anchorX,
            double anchorY,
            bool threaded,
            string title,
            string text,
            IReadOnlyList<ExcelVisualTextRun>? richTextRuns,
            string source) {
            Row = row;
            Column = column;
            X = x;
            Y = y;
            Width = width;
            Height = height;
            AnchorX = anchorX;
            AnchorY = anchorY;
            Threaded = threaded;
            Title = title ?? string.Empty;
            Text = text ?? string.Empty;
            RichTextRuns = richTextRuns ?? Array.Empty<ExcelVisualTextRun>();
            Source = source ?? string.Empty;
        }

        /// <summary>One-based source row.</summary>
        public int Row { get; }

        /// <summary>One-based source column.</summary>
        public int Column { get; }

        /// <summary>X position in CSS pixels.</summary>
        public double X { get; }

        /// <summary>Y position in CSS pixels.</summary>
        public double Y { get; }

        /// <summary>Callout width in CSS pixels.</summary>
        public double Width { get; }

        /// <summary>Callout height in CSS pixels.</summary>
        public double Height { get; }

        /// <summary>X coordinate of the cell-side anchor point in CSS pixels.</summary>
        public double AnchorX { get; }

        /// <summary>Y coordinate of the cell-side anchor point in CSS pixels.</summary>
        public double AnchorY { get; }

        /// <summary>Whether the callout represents threaded comments.</summary>
        public bool Threaded { get; }

        /// <summary>Callout title, usually the author or comment kind.</summary>
        public string Title { get; }

        /// <summary>Plain text body rendered inside the callout.</summary>
        public string Text { get; }

        /// <summary>Optional rich body text runs rendered inside the callout.</summary>
        public IReadOnlyList<ExcelVisualTextRun> RichTextRuns { get; }

        /// <summary>Source reference used by export diagnostics.</summary>
        public string Source { get; }
    }

    /// <summary>
    /// Positioned worksheet sparkline metadata in an Excel range snapshot.
    /// </summary>
    public sealed class ExcelVisualSparkline {
        internal ExcelVisualSparkline(
            int row,
            int column,
            double x,
            double y,
            double width,
            double height,
            string kind,
            IReadOnlyList<double> values,
            bool displayMarkers,
            bool displayHigh,
            bool displayLow,
            bool displayFirst,
            bool displayLast,
            bool displayNegative,
            bool displayAxis,
            string? seriesColorArgb,
            string? axisColorArgb,
            string? negativeColorArgb,
            string? markersColorArgb,
            string? highColorArgb,
            string? lowColorArgb,
            string? firstColorArgb,
            string? lastColorArgb,
            double? scaleMinimum,
            double? scaleMaximum,
            string source) {
            Row = row;
            Column = column;
            X = x;
            Y = y;
            Width = width;
            Height = height;
            Kind = kind ?? string.Empty;
            Values = values ?? Array.Empty<double>();
            DisplayMarkers = displayMarkers;
            DisplayHigh = displayHigh;
            DisplayLow = displayLow;
            DisplayFirst = displayFirst;
            DisplayLast = displayLast;
            DisplayNegative = displayNegative;
            DisplayAxis = displayAxis;
            SeriesColorArgb = seriesColorArgb;
            AxisColorArgb = axisColorArgb;
            NegativeColorArgb = negativeColorArgb;
            MarkersColorArgb = markersColorArgb;
            HighColorArgb = highColorArgb;
            LowColorArgb = lowColorArgb;
            FirstColorArgb = firstColorArgb;
            LastColorArgb = lastColorArgb;
            ScaleMinimum = scaleMinimum;
            ScaleMaximum = scaleMaximum;
            Source = source ?? string.Empty;
        }

        /// <summary>One-based source row.</summary>
        public int Row { get; }

        /// <summary>One-based source column.</summary>
        public int Column { get; }

        /// <summary>X position in CSS pixels.</summary>
        public double X { get; }

        /// <summary>Y position in CSS pixels.</summary>
        public double Y { get; }

        /// <summary>Cell width in CSS pixels.</summary>
        public double Width { get; }

        /// <summary>Cell height in CSS pixels.</summary>
        public double Height { get; }

        /// <summary>Sparkline kind as authored in the worksheet metadata.</summary>
        public string Kind { get; }

        /// <summary>Numeric values plotted by the sparkline.</summary>
        public IReadOnlyList<double> Values { get; }

        /// <summary>Whether point markers should be displayed.</summary>
        public bool DisplayMarkers { get; }

        /// <summary>Whether the high point should be highlighted.</summary>
        public bool DisplayHigh { get; }

        /// <summary>Whether the low point should be highlighted.</summary>
        public bool DisplayLow { get; }

        /// <summary>Whether the first point should be highlighted.</summary>
        public bool DisplayFirst { get; }

        /// <summary>Whether the last point should be highlighted.</summary>
        public bool DisplayLast { get; }

        /// <summary>Whether negative points should be highlighted.</summary>
        public bool DisplayNegative { get; }

        /// <summary>Whether a zero axis should be displayed when the data crosses zero.</summary>
        public bool DisplayAxis { get; }

        /// <summary>Series color in ARGB hexadecimal form, when specified.</summary>
        public string? SeriesColorArgb { get; }

        /// <summary>Axis color in ARGB hexadecimal form, when specified.</summary>
        public string? AxisColorArgb { get; }

        /// <summary>Negative point color in ARGB hexadecimal form, when specified.</summary>
        public string? NegativeColorArgb { get; }

        /// <summary>Regular marker color in ARGB hexadecimal form, when specified.</summary>
        public string? MarkersColorArgb { get; }

        /// <summary>High point color in ARGB hexadecimal form, when specified.</summary>
        public string? HighColorArgb { get; }

        /// <summary>Low point color in ARGB hexadecimal form, when specified.</summary>
        public string? LowColorArgb { get; }

        /// <summary>First point color in ARGB hexadecimal form, when specified.</summary>
        public string? FirstColorArgb { get; }

        /// <summary>Last point color in ARGB hexadecimal form, when specified.</summary>
        public string? LastColorArgb { get; }

        /// <summary>Minimum value used to scale the rendered sparkline, usually resolved from its Excel group.</summary>
        public double? ScaleMinimum { get; }

        /// <summary>Maximum value used to scale the rendered sparkline, usually resolved from its Excel group.</summary>
        public double? ScaleMaximum { get; }

        /// <summary>Source reference used by export diagnostics.</summary>
        public string Source { get; }
    }

    /// <summary>
    /// Positioned worksheet drawing object metadata in an Excel range snapshot.
    /// </summary>
    public sealed class ExcelVisualDrawingObject {
        internal ExcelVisualDrawingObject(
            string name,
            int order,
            string shapePresetName,
            OfficeShapeKind shapeKind,
            bool horizontalFlip,
            bool verticalFlip,
            double rotationDegrees,
            double x,
            double y,
            double width,
            double height,
            string? fillColorArgb,
            string? strokeColorArgb,
            double strokeWidth,
            OfficeStrokeDashStyle strokeDashStyle,
            OfficeStrokeLineCap? strokeLineCap,
            OfficeStrokeLineJoin? strokeLineJoin,
            string text,
            OfficeTextAlignment textAlignment,
            OfficeTextVerticalAlignment textVerticalAlignment,
            string? textColorArgb,
            string? textFontFamily,
            double? textFontSize,
            OfficeFontStyle textFontStyle,
            bool textWrap,
            bool textShrinkToFit,
            bool textResizeShapeToFit,
            ExcelDrawingTextOrientation textOrientation,
            double textInsetLeft,
            double textInsetTop,
            double textInsetRight,
            double textInsetBottom,
            OfficeGlow? glow,
            OfficeShadow? shadow,
            string source) {
            Name = name ?? string.Empty;
            Order = order;
            ShapePresetName = shapePresetName ?? string.Empty;
            ShapeKind = shapeKind;
            HorizontalFlip = horizontalFlip;
            VerticalFlip = verticalFlip;
            RotationDegrees = rotationDegrees;
            X = x;
            Y = y;
            Width = width;
            Height = height;
            FillColorArgb = fillColorArgb;
            StrokeColorArgb = strokeColorArgb;
            StrokeWidth = strokeWidth;
            StrokeDashStyle = strokeDashStyle;
            StrokeLineCap = strokeLineCap;
            StrokeLineJoin = strokeLineJoin;
            Text = text ?? string.Empty;
            TextAlignment = textAlignment;
            TextVerticalAlignment = textVerticalAlignment;
            TextColorArgb = textColorArgb;
            TextFontFamily = textFontFamily;
            TextFontSize = textFontSize;
            TextFontStyle = textFontStyle;
            TextWrap = textWrap;
            TextShrinkToFit = textShrinkToFit;
            TextResizeShapeToFit = textResizeShapeToFit;
            TextOrientation = textOrientation;
            TextInsetLeft = textInsetLeft;
            TextInsetTop = textInsetTop;
            TextInsetRight = textInsetRight;
            TextInsetBottom = textInsetBottom;
            Glow = glow;
            Shadow = shadow;
            Source = source ?? string.Empty;
        }

        /// <summary>Drawing object name.</summary>
        public string Name { get; }

        /// <summary>Zero-based source drawing layer order.</summary>
        public int Order { get; }

        /// <summary>Serialized DrawingML preset geometry name used to create the shared OfficeIMO.Drawing shape.</summary>
        public string ShapePresetName { get; }

        /// <summary>Shared OfficeIMO.Drawing shape kind.</summary>
        public OfficeShapeKind ShapeKind { get; }

        /// <summary>Whether the DrawingML geometry is mirrored horizontally.</summary>
        public bool HorizontalFlip { get; }

        /// <summary>Whether the DrawingML geometry is mirrored vertically.</summary>
        public bool VerticalFlip { get; }

        /// <summary>Clockwise DrawingML rotation in degrees.</summary>
        public double RotationDegrees { get; }

        /// <summary>Whether the drawing object has any authored rotation.</summary>
        public bool HasRotation => Math.Abs(RotationDegrees) > 0.0001D;

        /// <summary>X position in CSS pixels.</summary>
        public double X { get; }

        /// <summary>Y position in CSS pixels.</summary>
        public double Y { get; }

        /// <summary>Width in CSS pixels.</summary>
        public double Width { get; }

        /// <summary>Height in CSS pixels.</summary>
        public double Height { get; }

        /// <summary>Fill color in ARGB hexadecimal form, when supported.</summary>
        public string? FillColorArgb { get; }

        /// <summary>Stroke color in ARGB hexadecimal form, when supported.</summary>
        public string? StrokeColorArgb { get; }

        /// <summary>Stroke width in CSS pixels.</summary>
        public double StrokeWidth { get; }

        /// <summary>Stroke dash style resolved from the DrawingML outline.</summary>
        public OfficeStrokeDashStyle StrokeDashStyle { get; }

        /// <summary>Stroke line cap resolved from the DrawingML outline, when specified.</summary>
        public OfficeStrokeLineCap? StrokeLineCap { get; }

        /// <summary>Stroke line join resolved from the DrawingML outline, when specified.</summary>
        public OfficeStrokeLineJoin? StrokeLineJoin { get; }

        /// <summary>Plain text extracted from the drawing object.</summary>
        public string Text { get; }

        /// <summary>Horizontal text alignment extracted from the drawing object's paragraph properties.</summary>
        public OfficeTextAlignment TextAlignment { get; }

        /// <summary>Vertical text alignment extracted from the drawing object's body properties.</summary>
        public OfficeTextVerticalAlignment TextVerticalAlignment { get; }

        /// <summary>Text color in ARGB hexadecimal form, when supported.</summary>
        public string? TextColorArgb { get; }

        /// <summary>Text font family, when specified on the drawing text run.</summary>
        public string? TextFontFamily { get; }

        /// <summary>Text font size in points, when specified on the drawing text run.</summary>
        public double? TextFontSize { get; }

        /// <summary>Text font style flags extracted from the drawing text run.</summary>
        public OfficeFontStyle TextFontStyle { get; }

        /// <summary>Whether text should wrap inside the drawing object's text box.</summary>
        public bool TextWrap { get; }

        /// <summary>Whether DrawingML normalAutoFit should shrink overflowing text inside the text box.</summary>
        public bool TextShrinkToFit { get; }

        /// <summary>Whether DrawingML shapeAutoFit requested resizing the shape to fit text.</summary>
        public bool TextResizeShapeToFit { get; }

        /// <summary>Text orientation requested by DrawingML body properties.</summary>
        public ExcelDrawingTextOrientation TextOrientation { get; }

        /// <summary>Left text inset in CSS pixels after DrawingML EMU conversion.</summary>
        public double TextInsetLeft { get; }

        /// <summary>Top text inset in CSS pixels after DrawingML EMU conversion.</summary>
        public double TextInsetTop { get; }

        /// <summary>Right text inset in CSS pixels after DrawingML EMU conversion.</summary>
        public double TextInsetRight { get; }

        /// <summary>Bottom text inset in CSS pixels after DrawingML EMU conversion.</summary>
        public double TextInsetBottom { get; }

        /// <summary>DrawingML glow effect mapped to the shared dependency-free Drawing renderer, when supported.</summary>
        public OfficeGlow? Glow { get; }

        /// <summary>DrawingML outer shadow effect mapped to the shared dependency-free Drawing renderer, when supported.</summary>
        public OfficeShadow? Shadow { get; }

        /// <summary>Source reference used by export diagnostics.</summary>
        public string Source { get; }
    }

    /// <summary>
    /// Kind of worksheet drawing-layer overlay in an Excel range snapshot.
    /// </summary>
    public enum ExcelVisualDrawingLayerKind {
        /// <summary>A supported worksheet shape or text box.</summary>
        DrawingObject,

        /// <summary>A worksheet picture.</summary>
        Image,

        /// <summary>A worksheet chart.</summary>
        Chart,

        /// <summary>An opt-in worksheet comment or threaded-comment body callout.</summary>
        CommentBody
    }

    /// <summary>
    /// Ordered worksheet drawing-layer overlay in an Excel range snapshot.
    /// </summary>
    public sealed class ExcelVisualDrawingLayer {
        private ExcelVisualDrawingLayer(
            int order,
            ExcelVisualDrawingLayerKind kind,
            ExcelVisualDrawingObject? drawingObject,
            ExcelVisualImage? image,
            ExcelVisualChart? chart,
            ExcelVisualCommentBody? commentBody) {
            Order = order;
            Kind = kind;
            DrawingObject = drawingObject;
            Image = image;
            Chart = chart;
            CommentBody = commentBody;
        }

        internal static ExcelVisualDrawingLayer FromDrawingObject(ExcelVisualDrawingObject drawingObject) =>
            new ExcelVisualDrawingLayer(drawingObject.Order, ExcelVisualDrawingLayerKind.DrawingObject, drawingObject, null, null, null);

        internal static ExcelVisualDrawingLayer FromImage(ExcelVisualImage image) =>
            new ExcelVisualDrawingLayer(image.Order, ExcelVisualDrawingLayerKind.Image, null, image, null, null);

        internal static ExcelVisualDrawingLayer FromChart(ExcelVisualChart chart) =>
            new ExcelVisualDrawingLayer(chart.Order, ExcelVisualDrawingLayerKind.Chart, null, null, chart, null);

        internal static ExcelVisualDrawingLayer FromCommentBody(ExcelVisualCommentBody commentBody, int order) =>
            new ExcelVisualDrawingLayer(order, ExcelVisualDrawingLayerKind.CommentBody, null, null, null, commentBody);

        /// <summary>Zero-based source drawing layer order.</summary>
        public int Order { get; }

        /// <summary>Overlay kind.</summary>
        public ExcelVisualDrawingLayerKind Kind { get; }

        /// <summary>Drawing object payload when <see cref="Kind"/> is <see cref="ExcelVisualDrawingLayerKind.DrawingObject"/>.</summary>
        public ExcelVisualDrawingObject? DrawingObject { get; }

        /// <summary>Image payload when <see cref="Kind"/> is <see cref="ExcelVisualDrawingLayerKind.Image"/>.</summary>
        public ExcelVisualImage? Image { get; }

        /// <summary>Chart payload when <see cref="Kind"/> is <see cref="ExcelVisualDrawingLayerKind.Chart"/>.</summary>
        public ExcelVisualChart? Chart { get; }

        /// <summary>Comment body payload when <see cref="Kind"/> is <see cref="ExcelVisualDrawingLayerKind.CommentBody"/>.</summary>
        public ExcelVisualCommentBody? CommentBody { get; }
    }

    /// <summary>
    /// Positioned worksheet image metadata in an Excel range snapshot.
    /// </summary>
    public sealed class ExcelVisualImage {
        private readonly byte[] _bytes;

        internal ExcelVisualImage(
            string name,
            int order,
            string contentType,
            OfficeImageFormat detectedFormat,
            byte[] bytes,
            double sourceWidth,
            double sourceHeight,
            double x,
            double y,
            double width,
            double height,
            double cropLeftRatio,
            double cropTopRatio,
            double cropRightRatio,
            double cropBottomRatio,
            double rotationDegrees,
            bool flipHorizontal,
            bool flipVertical,
            bool isFullyOpaque,
            string source) {
            Name = name ?? string.Empty;
            Order = order;
            ContentType = contentType ?? string.Empty;
            DetectedFormat = detectedFormat;
            _bytes = bytes == null ? Array.Empty<byte>() : (byte[])bytes.Clone();
            SourceWidth = sourceWidth;
            SourceHeight = sourceHeight;
            X = x;
            Y = y;
            Width = width;
            Height = height;
            SourceCrop = OfficeImageSourceCrop.FromClampedFractions(cropLeftRatio, cropTopRatio, cropRightRatio, cropBottomRatio);
            RotationDegrees = rotationDegrees;
            FlipHorizontal = flipHorizontal;
            FlipVertical = flipVertical;
            IsFullyOpaque = isFullyOpaque;
            Source = source ?? string.Empty;
        }

        /// <summary>Drawing name.</summary>
        public string Name { get; }

        /// <summary>Zero-based source drawing layer order.</summary>
        public int Order { get; }

        /// <summary>Image content type.</summary>
        public string ContentType { get; }

        /// <summary>Detected image format based on the image bytes.</summary>
        public OfficeImageFormat DetectedFormat { get; }

        /// <summary>Image bytes.</summary>
        public byte[] Bytes => (byte[])_bytes.Clone();

        /// <summary>Intrinsic source image width in pixels, when it could be identified.</summary>
        public double SourceWidth { get; }

        /// <summary>Intrinsic source image height in pixels, when it could be identified.</summary>
        public double SourceHeight { get; }

        /// <summary>X position in CSS pixels.</summary>
        public double X { get; }

        /// <summary>Y position in CSS pixels.</summary>
        public double Y { get; }

        /// <summary>Image width in CSS pixels.</summary>
        public double Width { get; }

        /// <summary>Image height in CSS pixels.</summary>
        public double Height { get; }

        /// <summary>Normalized source-image crop from the authored image edges.</summary>
        public OfficeImageSourceCrop SourceCrop { get; }

        /// <summary>Normalized crop from the source image left edge.</summary>
        public double CropLeftRatio => SourceCrop.Left;

        /// <summary>Normalized crop from the source image top edge.</summary>
        public double CropTopRatio => SourceCrop.Top;

        /// <summary>Normalized crop from the source image right edge.</summary>
        public double CropRightRatio => SourceCrop.Right;

        /// <summary>Normalized crop from the source image bottom edge.</summary>
        public double CropBottomRatio => SourceCrop.Bottom;

        /// <summary>Whether the image has any authored crop rectangle.</summary>
        public bool HasCrop => SourceCrop.HasCrop;

        /// <summary>Clockwise rotation in degrees.</summary>
        public double RotationDegrees { get; }

        /// <summary>Whether the image has an authored horizontal flip transform.</summary>
        public bool FlipHorizontal { get; }

        /// <summary>Whether the image has an authored vertical flip transform.</summary>
        public bool FlipVertical { get; }

        /// <summary>Whether the decoded source image contains no transparent pixels.</summary>
        public bool IsFullyOpaque { get; }

        /// <summary>Whether the image has any authored rotation transform.</summary>
        public bool HasRotation => Math.Abs(RotationDegrees) > 0.0001D;

        /// <summary>Source reference used by export diagnostics.</summary>
        public string Source { get; }
    }

    /// <summary>
    /// Positioned worksheet chart metadata in an Excel range snapshot.
    /// </summary>
    public sealed class ExcelVisualChart {
        internal ExcelVisualChart(ExcelChartSnapshot snapshot, int order, double x, double y, double width, double height) {
            Snapshot = snapshot ?? throw new ArgumentNullException(nameof(snapshot));
            Order = order;
            X = x;
            Y = y;
            Width = width;
            Height = height;
        }

        /// <summary>Chart snapshot.</summary>
        public ExcelChartSnapshot Snapshot { get; }

        /// <summary>Zero-based source drawing layer order.</summary>
        public int Order { get; }

        /// <summary>X position in CSS pixels.</summary>
        public double X { get; }

        /// <summary>Y position in CSS pixels.</summary>
        public double Y { get; }

        /// <summary>Chart width in CSS pixels.</summary>
        public double Width { get; }

        /// <summary>Chart height in CSS pixels.</summary>
        public double Height { get; }
    }
}

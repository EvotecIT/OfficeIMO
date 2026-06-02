using System.Globalization;
using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private const double TableCellClipBleed = 2D;
    private const double TableCellCheckBoxGap = 2D;

    // Helper shapes for column pagination
    private abstract class ColItem { public string Kind = string.Empty; }
    private sealed class ColPar : ColItem { public RichParagraphBlock Block = null!; public System.Collections.Generic.List<System.Collections.Generic.List<RichSeg>> Lines = null!; public System.Collections.Generic.List<double> Heights = null!; public double Leading; public double Size; public double XOffset; public double TextWidth; public double FirstLineXOffset; public double FirstLineTextWidth; public ColPar() { Kind = "P"; } }
    private sealed class ColHead : ColItem { public HeadingBlock Block = null!; public System.Collections.Generic.List<string> Lines = null!; public double Leading; public double Size; public double SpacingBefore; public double SpacingAfter; public bool Bold; public bool ApplySpacingBeforeAtTop; public bool KeepWithNext; public PdfColor? Color; public ColHead() { Kind = "H"; } }
    private sealed class ColRule : ColItem { public HorizontalRuleBlock Block = null!; public ColRule() { Kind = "R"; } }
    private sealed class ColImg : ColItem { public ImageBlock Block = null!; public ColImg() { Kind = "I"; } }
    private sealed class ColShape : ColItem { public ShapeBlock Block = null!; public ColShape() { Kind = "S"; } }
    private sealed class ColDrawing : ColItem { public DrawingBlock Block = null!; public ColDrawing() { Kind = "D"; } }
    private sealed class ColForm : ColItem { public IPdfBlock Block = null!; public ColForm() { Kind = "FORM"; } }
    private sealed class ColBookmark : ColItem { public BookmarkBlock Block = null!; public ColBookmark() { Kind = "B"; } }
    private sealed class ColSpacer : ColItem { public SpacerBlock Block = null!; public ColSpacer() { Kind = "SPACE"; } }
    private sealed class ColListItem : ColItem { public System.Collections.Generic.IReadOnlyList<TextRun> Runs = null!; public System.Collections.Generic.List<System.Collections.Generic.List<RichSeg>> Lines = null!; public System.Collections.Generic.List<double> Heights = null!; public string Marker = string.Empty; public double MarkerXOffset; public double MarkerWidth; public PdfAlign MarkerAlign; public double TextXOffset; public double TextWidth; public PdfAlign TextAlign; public PdfColor? Color; public double Leading; public double Size; public double SpacingBefore; public double SpacingAfter; public string? BookmarkName; public bool KeepTogether; public bool IsFirstInKeepGroup; public double KeepGroupHeight; public bool KeepWithNext; public bool IsFirstInKeepWithNextGroup; public int KeepWithNextGroupItemCount; public double KeepWithNextGroupHeight; public ColListItem() { Kind = "L"; } }
    private sealed class ColPanel : ColItem { public PanelParagraphBlock Block = null!; public PanelStyle Style = null!; public System.Collections.Generic.List<System.Collections.Generic.List<RichSeg>> Lines = null!; public System.Collections.Generic.List<double> Heights = null!; public double Leading; public double Size; public double FirstBaselineOffset; public double XOffset; public double PanelWidth; public double TextWidth; public ColPanel() { Kind = "PANEL"; } }
    private sealed class ColTable : ColItem { public TableBlock Block = null!; public PdfTableStyle Style = null!; public int Columns; public double[] ColumnWidths = null!; public TableCellTextLayout[][] RowLines = null!; public int[] RowLineCounts = null!; public double[] RowHeights = null!; public double[] RowLeadings = null!; public double[] RowSizes = null!; public bool[] RowBold = null!; public double Width; public double Size; public int HeaderRowCount; public int RepeatHeaderRowCount; public int FooterStartRowIndex; public System.Collections.Generic.List<string>? CaptionLines; public double CaptionLeading; public double CaptionHeight; public ColTable() { Kind = "T"; } }
    private sealed class TableColumnLayout { public double[] Widths = null!; public double Width; }
    private sealed class TableCellTextLayout {
        public TableCellTextLayout(System.Collections.Generic.List<System.Collections.Generic.List<RichSeg>> lines, System.Collections.Generic.List<double> lineHeights) {
            Lines = lines;
            LineHeights = lineHeights;
        }

        public System.Collections.Generic.List<System.Collections.Generic.List<RichSeg>> Lines { get; }
        public System.Collections.Generic.List<double> LineHeights { get; }
        public int LineCount => System.Math.Max(1, Lines.Count);
    }
    private readonly struct TableCellLayout {
        public TableCellLayout(int column, int columnSpan, int rowSpan, string text, System.Collections.Generic.IReadOnlyList<TextRun> runs, string? linkUri, string? linkDestinationName, string? linkContents, string? namedDestinationName, System.Collections.Generic.IReadOnlyList<PdfTableCellCheckBox> checkBoxes, System.Collections.Generic.IReadOnlyList<PdfTableCellFormField> formFields, System.Collections.Generic.IReadOnlyList<PdfTableCellImage> images) {
            Column = column;
            ColumnSpan = columnSpan;
            RowSpan = rowSpan;
            Text = text;
            Runs = runs;
            LinkUri = linkUri;
            LinkDestinationName = linkDestinationName;
            LinkContents = linkContents;
            NamedDestinationName = namedDestinationName;
            CheckBoxes = checkBoxes;
            FormFields = formFields;
            Images = images;
        }

        public int Column { get; }
        public int ColumnSpan { get; }
        public int RowSpan { get; }
        public string Text { get; }
        public System.Collections.Generic.IReadOnlyList<TextRun> Runs { get; }
        public string? LinkUri { get; }
        public string? LinkDestinationName { get; }
        public string? LinkContents { get; }
        public string? NamedDestinationName { get; }
        public System.Collections.Generic.IReadOnlyList<PdfTableCellCheckBox> CheckBoxes { get; }
        public System.Collections.Generic.IReadOnlyList<PdfTableCellFormField> FormFields { get; }
        public System.Collections.Generic.IReadOnlyList<PdfTableCellImage> Images { get; }
    }

    private static LayoutResult LayoutBlocks(IEnumerable<IPdfBlock> blocks, PdfOptions opts) {
        var sb = new StringBuilder();
        var pages = new System.Collections.Generic.List<LayoutResult.Page>();
        var optionsStack = new System.Collections.Generic.Stack<PdfOptions>();
        optionsStack.Push(opts);
        var pageGroupStack = new System.Collections.Generic.Stack<int>();
        pageGroupStack.Push(0);
        PdfOptions currentOpts = opts;
        int currentPageGroupId = 0;
        int nextPageGroupId = 1;

        LayoutResult.Page? currentPage = null;
        double width = 0;
        double yStart = 0;
        double y = 0;
        bool pageDirty = false;

        bool usedBold = false;
        bool usedItalic = false;
        bool usedBoldItalic = false;
        var emittedTableCellNamedDestinations = new System.Collections.Generic.HashSet<string>(System.StringComparer.Ordinal);

        void StartPage(PdfOptions options) {
            options.Validate();
            currentOpts = options;
            width = options.PageWidth - options.MarginLeft - options.MarginRight;
            yStart = options.PageHeight - options.MarginTop;
            y = yStart;
            currentPage = new LayoutResult.Page { Options = options, PageGroupId = currentPageGroupId };
            sb.Clear();
            pageDirty = false;
        }

        void EnsurePage() {
            if (currentPage == null) StartPage(currentOpts);
        }

        bool HasCurrentPageNonContentObjects() =>
            currentPage != null &&
            (currentPage.Images.Count > 0 ||
            currentPage.Annotations.Count > 0 ||
            currentPage.FormFields.Count > 0 ||
            currentPage.GraphicsStates.Count > 0 ||
            currentPage.Shadings.Count > 0 ||
            currentPage.NamedDestinations.Count > 0);

        void FlushPage(bool force = false) {
            if (currentPage == null) return;
            if (!force && !pageDirty && !HasCurrentPageNonContentObjects()) {
                currentPage = null;
                sb.Clear();
                pageDirty = false;
                return;
            }
            currentPage.Content = sb.ToString();
            pages.Add(currentPage);
            currentPage = null;
            sb.Clear();
            pageDirty = false;
        }

        void NewPage() {
            FlushPage(pageDirty || HasCurrentPageNonContentObjects());
            StartPage(currentOpts);
        }

        double ResolveTopLevelSpacingBefore(double spacingBefore) {
            return y < yStart - 0.001 ? spacingBefore : 0D;
        }

        static double ResolveColumnSpacingBefore(double spacingBefore, double consumed) {
            return consumed > 0.001 ? spacingBefore : 0D;
        }

        void WriteLinesInternal(string fontRes, double fontSize, double lineHeight, double x, double widthUsed, double startY, System.Collections.Generic.IReadOnlyList<string> lines, PdfAlign align, PdfColor? color = null, bool applyBaselineTweak = false) {
            EnsurePage();
            pageDirty = true;
            var content = new ContentStreamBuilder(sb)
                .BeginText()
                .Font(fontRes, fontSize)
                .TextLeading(lineHeight);
            var lineFont = fontRes == "F2" ? ChooseBold(ChooseNormal(currentOpts.DefaultFont)) : ChooseNormal(currentOpts.DefaultFont);
            double yStart2 = startY;
            if (applyBaselineTweak) {
                yStart2 -= GetDescender(lineFont, fontSize) * 0.0;
            }
            content.TextMatrix(x, yStart2);
            var effectiveColor = color ?? currentOpts.DefaultTextColor ?? PdfColor.Black;
            content.FillColor(effectiveColor);
            for (int i = 0; i < lines.Count; i++) {
                string line = lines[i];
                double dx = 0;
                double estWidth = EstimateSimpleTextWidth(line, lineFont, fontSize);
                if (align == PdfAlign.Center) dx = Math.Max(0, (widthUsed - estWidth) / 2);
                else if (align == PdfAlign.Right) dx = Math.Max(0, (widthUsed - estWidth));
                if (Math.Abs(dx) > 0.0001) content.MoveText(dx, 0);
                content.ShowHexText(EncodeWinAnsiHex(line));
                if (Math.Abs(dx) > 0.0001) content.MoveText(-dx, 0);
                if (i != lines.Count - 1) content.NextTextLine();
            }
            content.EndText();
        }

        void WriteLines(string fontRes, double fontSize, double lineHeight, double x, double startY, System.Collections.Generic.IReadOnlyList<string> lines, PdfAlign align, PdfColor? color = null, bool applyBaselineTweak = false)
            => WriteLinesInternal(fontRes, fontSize, lineHeight, x, width, startY, lines, align, color, applyBaselineTweak);

        void AddHeadingLinkAnnotations(HeadingBlock heading, System.Collections.Generic.IReadOnlyList<string> lines, PdfStandardFont font, double fontSize, double lineHeight, double x, double widthUsed, double startBaselineY) {
            if (string.IsNullOrEmpty(heading.LinkUri) && string.IsNullOrEmpty(heading.LinkDestinationName)) {
                return;
            }

            double asc = GetAscender(font, fontSize);
            double desc = GetDescender(font, fontSize);
            for (int i = 0; i < lines.Count; i++) {
                string line = lines[i];
                double lineWidth = EstimateSimpleTextWidth(line, font, fontSize);
                if (lineWidth <= 0.001D) {
                    continue;
                }

                double dx = 0D;
                if (heading.Align == PdfAlign.Center) dx = Math.Max(0, (widthUsed - lineWidth) / 2);
                else if (heading.Align == PdfAlign.Right) dx = Math.Max(0, widthUsed - lineWidth);
                double baselineY = startBaselineY - i * lineHeight;
                double x1 = x + dx;
                double x2 = x1 + Math.Min(widthUsed, lineWidth);
                double y1 = baselineY - desc;
                double y2 = baselineY + asc;
                currentPage!.Annotations.Add(new LinkAnnotation { X1 = x1, Y1 = y1, X2 = x2, Y2 = y2, Uri = heading.LinkUri, DestinationName = heading.LinkDestinationName, Contents = heading.LinkContents });
            }
        }

        void AddImageLinkAnnotation(ImageBlock image, PdfImageStyle style, PageImage pageImage, double targetX, double targetBottomY) {
            if (string.IsNullOrEmpty(image.LinkUri)) {
                return;
            }

            double x1 = pageImage.X;
            double y1 = pageImage.Y;
            double x2 = pageImage.X + pageImage.W;
            double y2 = pageImage.Y + pageImage.H;
            if (style.Fit == OfficeImageFit.Cover || style.ClipPath != null) {
                x1 = targetX;
                y1 = targetBottomY;
                x2 = targetX + image.Width;
                y2 = targetBottomY + image.Height;
            }

            currentPage!.Annotations.Add(new LinkAnnotation { X1 = x1, Y1 = y1, X2 = x2, Y2 = y2, Uri = image.LinkUri!, Contents = image.LinkContents });
        }

        double GetAlignedObjectX(double containerX, double containerWidth, double objectWidth, PdfAlign align) {
            if (align == PdfAlign.Center) return containerX + Math.Max(0, (containerWidth - objectWidth) / 2);
            if (align == PdfAlign.Right) return containerX + Math.Max(0, containerWidth - objectWidth);
            return containerX;
        }

        void AddShapeLinkAnnotation(ShapeBlock shape, PdfDrawingStyle style, double containerX, double containerWidth, double topY) {
            if (string.IsNullOrEmpty(shape.LinkUri)) {
                return;
            }

            double x = GetAlignedObjectX(containerX, containerWidth, shape.Shape.Width, style.Align);
            currentPage!.Annotations.Add(new LinkAnnotation { X1 = x, Y1 = topY - shape.Shape.Height, X2 = x + shape.Shape.Width, Y2 = topY, Uri = shape.LinkUri!, Contents = shape.LinkContents });
        }

        void AddDrawingLinkAnnotation(DrawingBlock drawing, PdfDrawingStyle style, double containerX, double containerWidth, double topY) {
            if (string.IsNullOrEmpty(drawing.LinkUri)) {
                return;
            }

            double x = GetAlignedObjectX(containerX, containerWidth, drawing.Drawing.Width, style.Align);
            currentPage!.Annotations.Add(new LinkAnnotation { X1 = x, Y1 = topY - drawing.Drawing.Height, X2 = x + drawing.Drawing.Width, Y2 = topY, Uri = drawing.LinkUri!, Contents = drawing.LinkContents });
        }

        void AddNamedDestination(BookmarkBlock bookmark, double topY) {
            AddNamedDestinationName(bookmark.Name, topY);
        }

        void AddNamedDestinationName(string name, double topY) {
            EnsurePage();
            currentPage!.NamedDestinations.Add(new PageNamedDestination { Name = name, Y = topY });
        }

        void AddTableCellNamedDestinationName(string? name, double topY) {
            if (string.IsNullOrWhiteSpace(name) || !emittedTableCellNamedDestinations.Add(name!)) {
                return;
            }

            AddNamedDestinationName(name!, topY);
        }

        double FirstTextBaselineFromTop(PdfStandardFont font, double fontSize, double topY) =>
            topY - GetAscender(font, fontSize);

        PdfColor? ToPdfColor(OfficeIMO.Drawing.OfficeColor? color) =>
            color.HasValue ? PdfColor.FromOfficeColorOrNull(color.Value) : null;

        string? EnsureGraphicsState(double fillOpacity, double strokeOpacity) {
            if (fillOpacity >= 1D && strokeOpacity >= 1D) {
                return null;
            }

            EnsurePage();
            for (int i = 0; i < currentPage!.GraphicsStates.Count; i++) {
                var existing = currentPage.GraphicsStates[i];
                if (existing.FillOpacity.Equals(fillOpacity) && existing.StrokeOpacity.Equals(strokeOpacity)) {
                    return existing.Name;
                }
            }

            string name = "GS" + (currentPage.GraphicsStates.Count + 1).ToString(CultureInfo.InvariantCulture);
            currentPage.GraphicsStates.Add(new PageGraphicsState {
                Name = name,
                FillOpacity = fillOpacity,
                StrokeOpacity = strokeOpacity
            });
            return name;
        }

        string? EnsureOpacityState(OfficeIMO.Drawing.OfficeShape shape) {
            bool hasFill = (shape.FillColor.HasValue || shape.FillGradient != null) && shape.Kind != OfficeIMO.Drawing.OfficeShapeKind.Line;
            bool hasStroke = shape.StrokeColor.HasValue && shape.StrokeWidth > 0;
            double fillOpacity = hasFill ? shape.FillOpacity ?? 1D : 1D;
            double strokeOpacity = hasStroke ? shape.StrokeOpacity ?? 1D : 1D;
            return EnsureGraphicsState(fillOpacity, strokeOpacity);
        }

        string? EnsureLinearGradient(OfficeIMO.Drawing.OfficeShape shape, double xShape, double bottomY, bool localCoordinates) {
            var gradient = shape.FillGradient;
            if (gradient == null || shape.Kind == OfficeIMO.Drawing.OfficeShapeKind.Line) {
                return null;
            }

            var start = gradient.Stops[0].Color;
            var end = gradient.Stops[1].Color;
            double originX = localCoordinates ? 0D : xShape;
            double originY = localCoordinates ? 0D : bottomY;
            double x0 = originX + gradient.StartX * shape.Width;
            double y0 = originY + shape.Height - gradient.StartY * shape.Height;
            double x1 = originX + gradient.EndX * shape.Width;
            double y1 = originY + shape.Height - gradient.EndY * shape.Height;

            EnsurePage();
            for (int i = 0; i < currentPage!.Shadings.Count; i++) {
                var existing = currentPage.Shadings[i];
                if (existing.StartColor.Equals(start) &&
                    existing.EndColor.Equals(end) &&
                    existing.X0.Equals(x0) &&
                    existing.Y0.Equals(y0) &&
                    existing.X1.Equals(x1) &&
                    existing.Y1.Equals(y1)) {
                    return existing.Name;
                }
            }

            string name = "SH" + (currentPage.Shadings.Count + 1).ToString(CultureInfo.InvariantCulture);
            currentPage.Shadings.Add(new PageShading {
                Name = name,
                StartColor = start,
                EndColor = end,
                X0 = x0,
                Y0 = y0,
                X1 = x1,
                Y1 = y1
            });
            return name;
        }

        void DrawShapeShadowAt(OfficeIMO.Drawing.OfficeShape shape, double xShape, double bottomY) {
            var shadow = shape.Shadow;
            if (shadow == null || shadow.Opacity <= 0D) {
                return;
            }

            PdfColor shadowColor = PdfColor.FromOfficeColor(shadow.Color);
            double shadowX = xShape + shadow.OffsetX;
            double shadowBottomY = bottomY - shadow.OffsetY;
            string? shadowState = EnsureGraphicsState(shadow.Opacity, shadow.Opacity);

            var content = new ContentStreamBuilder(sb)
                .SaveState();
            if (shadowState != null) {
                content.GraphicsState(shadowState);
            }

            if (shape.Transform.HasValue) {
                DrawTransformedShape(
                    sb,
                    shape,
                    shape.Kind == OfficeIMO.Drawing.OfficeShapeKind.Line ? null : shadowColor,
                    shape.Kind == OfficeIMO.Drawing.OfficeShapeKind.Line ? shadowColor : null,
                    null,
                    shadowX,
                    shadowBottomY);
            } else if (shape.Kind == OfficeIMO.Drawing.OfficeShapeKind.Line) {
                DrawLine(sb, shadowColor, shape.StrokeWidth, shape.StrokeDashStyle, shape.StrokeLineCap, shape.StrokeLineJoin, shape.Points, shadowX, shadowBottomY, shape.Height);
            } else if (shape.Kind == OfficeIMO.Drawing.OfficeShapeKind.RoundedRectangle) {
                DrawRoundedRectangle(sb, shadowColor, null, 0, shape.StrokeDashStyle, shape.StrokeLineCap, shape.StrokeLineJoin, shadowX, shadowBottomY, shape.Width, shape.Height, shape.CornerRadius);
            } else if (shape.Kind == OfficeIMO.Drawing.OfficeShapeKind.Rectangle) {
                DrawRectangle(sb, shadowColor, null, 0, shape.StrokeDashStyle, shape.StrokeLineCap, shape.StrokeLineJoin, shadowX, shadowBottomY, shape.Width, shape.Height);
            } else if (shape.Kind == OfficeIMO.Drawing.OfficeShapeKind.Ellipse) {
                DrawEllipse(sb, shadowColor, null, 0, shape.StrokeDashStyle, shape.StrokeLineCap, shape.StrokeLineJoin, shadowX, shadowBottomY, shape.Width, shape.Height);
            } else if (shape.Kind == OfficeIMO.Drawing.OfficeShapeKind.Polygon) {
                DrawPolygon(sb, shadowColor, null, 0, shape.StrokeDashStyle, shape.StrokeLineCap, shape.StrokeLineJoin, shape.Points, shadowX, shadowBottomY, shape.Height);
            } else if (shape.Kind == OfficeIMO.Drawing.OfficeShapeKind.Path) {
                DrawPath(sb, shadowColor, null, 0, shape.StrokeDashStyle, shape.StrokeLineCap, shape.StrokeLineJoin, shape.PathCommands, shadowX, shadowBottomY, shape.Height);
            }

            content.RestoreState();
            pageDirty = true;
        }

        void DrawShapeGeometryAt(OfficeIMO.Drawing.OfficeShape shape, double xShape, double bottomY) {
            DrawShapeShadowAt(shape, xShape, bottomY);

            string? opacityState = EnsureOpacityState(shape);
            if (opacityState != null) {
                new ContentStreamBuilder(sb)
                    .SaveState()
                    .GraphicsState(opacityState);
            }

            if (shape.Transform.HasValue) {
                pageDirty = true;
                string? shadingName = EnsureLinearGradient(shape, xShape, bottomY, localCoordinates: true);
                DrawTransformedShape(sb, shape, shadingName == null ? ToPdfColor(shape.FillColor) : null, ToPdfColor(shape.StrokeColor), shadingName, xShape, bottomY);
            } else {
                if (shape.ClipPath != null) {
                    new ContentStreamBuilder(sb)
                        .SaveState();
                    AppendClipPath(sb, shape.ClipPath, xShape, bottomY, shape.Height);
                }

                string? shadingName = EnsureLinearGradient(shape, xShape, bottomY, localCoordinates: false);
                if (shadingName != null) {
                    pageDirty = true;
                    DrawGradientShape(sb, shape, shadingName, xShape, bottomY);
                }

                PdfColor? fillColor = shadingName == null ? ToPdfColor(shape.FillColor) : null;
                if (shape.Kind == OfficeIMO.Drawing.OfficeShapeKind.Line) {
                    pageDirty = true;
                    DrawLine(sb, ToPdfColor(shape.StrokeColor), shape.StrokeWidth, shape.StrokeDashStyle, shape.StrokeLineCap, shape.StrokeLineJoin, shape.Points, xShape, bottomY, shape.Height);
                } else if (shape.Kind == OfficeIMO.Drawing.OfficeShapeKind.RoundedRectangle) {
                    pageDirty = true;
                    DrawRoundedRectangle(sb, fillColor, ToPdfColor(shape.StrokeColor), shape.StrokeWidth, shape.StrokeDashStyle, shape.StrokeLineCap, shape.StrokeLineJoin, xShape, bottomY, shape.Width, shape.Height, shape.CornerRadius);
                } else if (shape.Kind == OfficeIMO.Drawing.OfficeShapeKind.Rectangle) {
                    pageDirty = true;
                    DrawRectangle(sb, fillColor, ToPdfColor(shape.StrokeColor), shape.StrokeWidth, shape.StrokeDashStyle, shape.StrokeLineCap, shape.StrokeLineJoin, xShape, bottomY, shape.Width, shape.Height);
                } else if (shape.Kind == OfficeIMO.Drawing.OfficeShapeKind.Ellipse) {
                    pageDirty = true;
                    DrawEllipse(sb, fillColor, ToPdfColor(shape.StrokeColor), shape.StrokeWidth, shape.StrokeDashStyle, shape.StrokeLineCap, shape.StrokeLineJoin, xShape, bottomY, shape.Width, shape.Height);
                } else if (shape.Kind == OfficeIMO.Drawing.OfficeShapeKind.Polygon) {
                    pageDirty = true;
                    DrawPolygon(sb, fillColor, ToPdfColor(shape.StrokeColor), shape.StrokeWidth, shape.StrokeDashStyle, shape.StrokeLineCap, shape.StrokeLineJoin, shape.Points, xShape, bottomY, shape.Height);
                } else if (shape.Kind == OfficeIMO.Drawing.OfficeShapeKind.Path) {
                    pageDirty = true;
                    DrawPath(sb, fillColor, ToPdfColor(shape.StrokeColor), shape.StrokeWidth, shape.StrokeDashStyle, shape.StrokeLineCap, shape.StrokeLineJoin, shape.PathCommands, xShape, bottomY, shape.Height);
                }

                if (shape.ClipPath != null) {
                    new ContentStreamBuilder(sb)
                        .RestoreState();
                }
            }

            if (opacityState != null) {
                new ContentStreamBuilder(sb)
                    .RestoreState();
            }
        }

        void DrawShapeAt(ShapeBlock block, PdfDrawingStyle style, double containerX, double containerWidth, double topY) {
            double xShape = GetAlignedObjectX(containerX, containerWidth, block.Shape.Width, style.Align);
            DrawShapeGeometryAt(block.Shape, xShape, topY - block.Shape.Height);
        }

        void DrawDrawingAt(DrawingBlock block, PdfDrawingStyle style, double containerX, double containerWidth, double topY) {
            double xDrawing = GetAlignedObjectX(containerX, containerWidth, block.Drawing.Width, style.Align);
            for (int i = 0; i < block.Drawing.Shapes.Count; i++) {
                var item = block.Drawing.Shapes[i];
                double xShape = xDrawing + item.X;
                double bottomY = topY - item.Y - item.Shape.Height;
                DrawShapeGeometryAt(item.Shape, xShape, bottomY);
            }
        }

        void RenderShapeBlock(ShapeBlock block, double containerX, double containerWidth) {
            PdfDrawingStyle style = ResolveDrawingStyle(block, currentOpts);
            PdfDoc.ValidateDrawingStyle(style, "Shape");
            double spacingBefore = ResolveTopLevelSpacingBefore(style.SpacingBefore);
            double needed = spacingBefore + block.Shape.Height + style.SpacingAfter;
            EnsureFixedFlowBlockFits("Shape", block.Shape.Width, needed, containerWidth);
            if (y - needed < currentOpts.MarginBottom) {
                NewPage();
                spacingBefore = 0D;
            }
            if (spacingBefore > 0) y -= spacingBefore;
            DrawShapeAt(block, style, containerX, containerWidth, y);
            AddShapeLinkAnnotation(block, style, containerX, containerWidth, y);
            y -= block.Shape.Height + style.SpacingAfter;
        }

        void RenderDrawingBlock(DrawingBlock block, double containerX, double containerWidth) {
            PdfDrawingStyle style = ResolveDrawingStyle(block, currentOpts);
            PdfDoc.ValidateDrawingStyle(style, "Drawing");
            double spacingBefore = ResolveTopLevelSpacingBefore(style.SpacingBefore);
            double needed = spacingBefore + block.Drawing.Height + style.SpacingAfter;
            EnsureFixedFlowBlockFits("Drawing", block.Drawing.Width, needed, containerWidth);
            if (y - needed < currentOpts.MarginBottom) {
                NewPage();
                spacingBefore = 0D;
            }
            if (spacingBefore > 0) y -= spacingBefore;
            DrawDrawingAt(block, style, containerX, containerWidth, y);
            AddDrawingLinkAnnotation(block, style, containerX, containerWidth, y);
            y -= block.Drawing.Height + style.SpacingAfter;
        }

        void KeepFixedBlockWithNext(double needed, double nextHeight) {
            double keepHeight = needed + nextHeight;
            double availableHeight = currentOpts.PageHeight - currentOpts.MarginTop - currentOpts.MarginBottom;
            if (nextHeight > 0.001 && keepHeight <= availableHeight + 0.001 && y < yStart - 0.001 && y - keepHeight < currentOpts.MarginBottom) {
                NewPage();
            }
        }

        void RenderHorizontalRuleBlock(HorizontalRuleBlock block, double containerX, double containerWidth) {
            PdfHorizontalRuleStyle ruleStyle = ResolveHorizontalRuleStyle(block, currentOpts);
            ValidateHorizontalRule(ruleStyle);
            double spacingBefore = ResolveTopLevelSpacingBefore(ruleStyle.SpacingBefore);
            double needed = spacingBefore + ruleStyle.Thickness + ruleStyle.SpacingAfter;
            EnsureFixedFlowBlockFits("Horizontal rule", containerWidth, needed, containerWidth);
            if (y - needed < currentOpts.MarginBottom) {
                NewPage();
                spacingBefore = 0D;
            }
            if (spacingBefore > 0) y -= spacingBefore;
            double yLine = y - ruleStyle.Thickness * 0.5;
            DrawHLine(sb, ruleStyle.Color, ruleStyle.Thickness, containerX, containerX + containerWidth, yLine);
            pageDirty = true;
            y -= ruleStyle.Thickness + ruleStyle.SpacingAfter;
        }

        void RenderTextFieldBlock(TextFieldBlock block, double containerX, double containerWidth) {
            double spacingBefore = ResolveTopLevelSpacingBefore(block.SpacingBefore);
            double needed = spacingBefore + block.Height + block.SpacingAfter;
            EnsureFixedFlowBlockFits("Text field", block.Width, needed, containerWidth);
            if (y - needed < currentOpts.MarginBottom) {
                NewPage();
                spacingBefore = 0D;
            }

            if (spacingBefore > 0) {
                y -= spacingBefore;
            }

            double x = GetAlignedObjectX(containerX, containerWidth, block.Width, block.Align);
            currentPage!.FormFields.Add(new FormFieldAnnotation {
                X1 = x,
                Y1 = y - block.Height,
                X2 = x + block.Width,
                Y2 = y,
                Kind = FormFieldAnnotationKind.Text,
                Name = block.Name,
                Value = block.Value,
                FontSize = block.FontSize,
                Style = block.Style
            });
            pageDirty = true;
            y -= block.Height + block.SpacingAfter;
        }

        void RenderCheckBoxBlock(CheckBoxBlock block, double containerX, double containerWidth) {
            double spacingBefore = ResolveTopLevelSpacingBefore(block.SpacingBefore);
            double needed = spacingBefore + block.Size + block.SpacingAfter;
            EnsureFixedFlowBlockFits("Check box", block.Size, needed, containerWidth);
            if (y - needed < currentOpts.MarginBottom) {
                NewPage();
                spacingBefore = 0D;
            }

            if (spacingBefore > 0) {
                y -= spacingBefore;
            }

            double x = GetAlignedObjectX(containerX, containerWidth, block.Size, block.Align);
            currentPage!.FormFields.Add(new FormFieldAnnotation {
                X1 = x,
                Y1 = y - block.Size,
                X2 = x + block.Size,
                Y2 = y,
                Kind = FormFieldAnnotationKind.CheckBox,
                Name = block.Name,
                Value = block.IsChecked ? block.CheckedValueName : "Off",
                IsChecked = block.IsChecked,
                CheckedValueName = block.CheckedValueName,
                Style = block.Style
            });
            pageDirty = true;
            y -= block.Size + block.SpacingAfter;
        }

        void RenderChoiceFieldBlock(ChoiceFieldBlock block, double containerX, double containerWidth) {
            double spacingBefore = ResolveTopLevelSpacingBefore(block.SpacingBefore);
            double needed = spacingBefore + block.Height + block.SpacingAfter;
            EnsureFixedFlowBlockFits("Choice field", block.Width, needed, containerWidth);
            if (y - needed < currentOpts.MarginBottom) {
                NewPage();
                spacingBefore = 0D;
            }

            if (spacingBefore > 0) {
                y -= spacingBefore;
            }

            double x = GetAlignedObjectX(containerX, containerWidth, block.Width, block.Align);
            currentPage!.FormFields.Add(new FormFieldAnnotation {
                X1 = x,
                Y1 = y - block.Height,
                X2 = x + block.Width,
                Y2 = y,
                Kind = FormFieldAnnotationKind.Choice,
                Name = block.Name,
                Value = block.Value,
                Values = block.Values,
                FontSize = block.FontSize,
                Options = block.Options,
                IsComboBox = block.IsComboBox,
                AllowsMultipleSelection = block.AllowsMultipleSelection,
                Style = block.Style
            });
            pageDirty = true;
            y -= block.Height + block.SpacingAfter;
        }

        void RenderRadioButtonGroupBlock(RadioButtonGroupBlock block, double containerX, double containerWidth) {
            double spacingBefore = ResolveTopLevelSpacingBefore(block.SpacingBefore);
            double height = block.Height;
            double needed = spacingBefore + height + block.SpacingAfter;
            double groupWidth = GetRadioButtonGroupWidth(block);
            EnsureFixedFlowBlockFits("Radio button group", groupWidth, needed, containerWidth);
            if (y - needed < currentOpts.MarginBottom) {
                NewPage();
                spacingBefore = 0D;
            }

            if (spacingBefore > 0) {
                y -= spacingBefore;
            }

            double x = GetAlignedObjectX(containerX, containerWidth, groupWidth, block.Align);
            currentPage!.FormFields.Add(new FormFieldAnnotation {
                X1 = x,
                Y1 = y - height,
                X2 = x + block.Size,
                Y2 = y,
                Kind = FormFieldAnnotationKind.RadioButtonGroup,
                Name = block.Name,
                Value = block.Value,
                Options = block.Options,
                ButtonSize = block.Size,
                ButtonGap = block.Gap,
                Style = block.Style
            });
            RenderRadioButtonLabels(block, x, y);
            pageDirty = true;
            y -= height + block.SpacingAfter;
        }

        static string GetFormFieldBlockName(IPdfBlock block) {
            if (block is TextFieldBlock) {
                return "Text field";
            }

            if (block is CheckBoxBlock) {
                return "Check box";
            }

            if (block is RadioButtonGroupBlock) {
                return "Radio button group";
            }

            return "Choice field";
        }

        double GetFormFieldWidth(IPdfBlock block) {
            if (block is TextFieldBlock textField) {
                return textField.Width;
            }

            if (block is CheckBoxBlock checkBox) {
                return checkBox.Size;
            }

            if (block is RadioButtonGroupBlock radioButtonGroup) {
                return GetRadioButtonGroupWidth(radioButtonGroup);
            }

            return ((ChoiceFieldBlock)block).Width;
        }

        static double GetFormFieldHeight(IPdfBlock block) {
            if (block is TextFieldBlock textField) {
                return textField.Height;
            }

            if (block is CheckBoxBlock checkBox) {
                return checkBox.Size;
            }

            if (block is RadioButtonGroupBlock radioButtonGroup) {
                return radioButtonGroup.Height;
            }

            return ((ChoiceFieldBlock)block).Height;
        }

        static double GetFormFieldSpacingBefore(IPdfBlock block) {
            if (block is TextFieldBlock textField) {
                return textField.SpacingBefore;
            }

            if (block is CheckBoxBlock checkBox) {
                return checkBox.SpacingBefore;
            }

            if (block is RadioButtonGroupBlock radioButtonGroup) {
                return radioButtonGroup.SpacingBefore;
            }

            return ((ChoiceFieldBlock)block).SpacingBefore;
        }

        static double GetFormFieldSpacingAfter(IPdfBlock block) {
            if (block is TextFieldBlock textField) {
                return textField.SpacingAfter;
            }

            if (block is CheckBoxBlock checkBox) {
                return checkBox.SpacingAfter;
            }

            if (block is RadioButtonGroupBlock radioButtonGroup) {
                return radioButtonGroup.SpacingAfter;
            }

            return ((ChoiceFieldBlock)block).SpacingAfter;
        }

        static PdfAlign GetFormFieldAlign(IPdfBlock block) {
            if (block is TextFieldBlock textField) {
                return textField.Align;
            }

            if (block is CheckBoxBlock checkBox) {
                return checkBox.Align;
            }

            if (block is RadioButtonGroupBlock radioButtonGroup) {
                return radioButtonGroup.Align;
            }

            return ((ChoiceFieldBlock)block).Align;
        }

        void AddFormFieldAnnotation(IPdfBlock block, double x, double topY) {
            if (block is TextFieldBlock textField) {
                currentPage!.FormFields.Add(new FormFieldAnnotation {
                    X1 = x,
                    Y1 = topY - textField.Height,
                    X2 = x + textField.Width,
                    Y2 = topY,
                    Kind = FormFieldAnnotationKind.Text,
                    Name = textField.Name,
                    Value = textField.Value,
                    FontSize = textField.FontSize,
                    Style = textField.Style
                });
                return;
            }

            if (block is CheckBoxBlock checkBox) {
                currentPage!.FormFields.Add(new FormFieldAnnotation {
                    X1 = x,
                    Y1 = topY - checkBox.Size,
                    X2 = x + checkBox.Size,
                    Y2 = topY,
                    Kind = FormFieldAnnotationKind.CheckBox,
                    Name = checkBox.Name,
                    Value = checkBox.IsChecked ? checkBox.CheckedValueName : "Off",
                    IsChecked = checkBox.IsChecked,
                    CheckedValueName = checkBox.CheckedValueName,
                    Style = checkBox.Style
                });
                return;
            }

            if (block is RadioButtonGroupBlock radioButtonGroup) {
                currentPage!.FormFields.Add(new FormFieldAnnotation {
                    X1 = x,
                    Y1 = topY - radioButtonGroup.Height,
                    X2 = x + radioButtonGroup.Size,
                    Y2 = topY,
                    Kind = FormFieldAnnotationKind.RadioButtonGroup,
                    Name = radioButtonGroup.Name,
                    Value = radioButtonGroup.Value,
                    Options = radioButtonGroup.Options,
                    ButtonSize = radioButtonGroup.Size,
                    ButtonGap = radioButtonGroup.Gap,
                    Style = radioButtonGroup.Style
                });
                RenderRadioButtonLabels(radioButtonGroup, x, topY);
                return;
            }

            ChoiceFieldBlock choice = (ChoiceFieldBlock)block;
            currentPage!.FormFields.Add(new FormFieldAnnotation {
                X1 = x,
                Y1 = topY - choice.Height,
                X2 = x + choice.Width,
                Y2 = topY,
                Kind = FormFieldAnnotationKind.Choice,
                Name = choice.Name,
                Value = choice.Value,
                Values = choice.Values,
                FontSize = choice.FontSize,
                Options = choice.Options,
                IsComboBox = choice.IsComboBox,
                AllowsMultipleSelection = choice.AllowsMultipleSelection,
                Style = choice.Style
            });
        }

        void EnsureFixedFlowBlockFits(string blockName, double blockWidth, double blockHeight, double availableWidth) {
            if (blockWidth > availableWidth + 0.001) {
                throw new ArgumentException(blockName + " width exceeds the available page content width.");
            }

            double availableHeight = currentOpts.PageHeight - currentOpts.MarginTop - currentOpts.MarginBottom;
            if (blockHeight > availableHeight + 0.001) {
                throw new ArgumentException(blockName + " height exceeds the available page content height.");
            }
        }

        double GetRadioButtonGroupLabelFontSize(RadioButtonGroupBlock block) =>
            System.Math.Min(System.Math.Max(8D, currentOpts.DefaultFontSize), System.Math.Max(8D, block.Size));

        double GetRadioButtonGroupLabelGap(RadioButtonGroupBlock block) =>
            System.Math.Max(4D, block.Size * 0.4D);

        double GetRadioButtonGroupWidth(RadioButtonGroupBlock block) {
            PdfStandardFont font = ChooseNormal(currentOpts.DefaultFont);
            double fontSize = GetRadioButtonGroupLabelFontSize(block);
            double labelWidth = block.Options.Max(option => EstimateSimpleTextWidth(option, font, fontSize));
            return block.Size + GetRadioButtonGroupLabelGap(block) + labelWidth;
        }

        void RenderRadioButtonLabels(RadioButtonGroupBlock block, double x, double topY) {
            PdfStandardFont font = ChooseNormal(currentOpts.DefaultFont);
            string fontResource = GetStandardFontResourceName(font, font);
            double fontSize = GetRadioButtonGroupLabelFontSize(block);
            double labelX = x + block.Size + GetRadioButtonGroupLabelGap(block);
            double ascender = GetAscender(font, fontSize);
            double descender = GetDescender(font, fontSize);
            double labelBaselineOffset = (block.Size - ascender - descender) / 2D + descender;

            for (int i = 0; i < block.Options.Count; i++) {
                double optionTop = topY - i * (block.Size + block.Gap);
                double baseline = optionTop - block.Size + labelBaselineOffset;
                AppendPageText(sb, block.Options[i], fontResource, fontSize, block.Style.TextColor, labelX, baseline);
            }
        }

        void ValidateHorizontalRule(PdfHorizontalRuleStyle rule) {
            if (rule.Thickness <= 0 || double.IsNaN(rule.Thickness) || double.IsInfinity(rule.Thickness)) {
                throw new ArgumentException("Horizontal rule thickness must be a positive finite value.");
            }

            if (rule.SpacingBefore < 0 || double.IsNaN(rule.SpacingBefore) || double.IsInfinity(rule.SpacingBefore)) {
                throw new ArgumentException("Horizontal rule spacing before must be a non-negative finite value.");
            }

            if (rule.SpacingAfter < 0 || double.IsNaN(rule.SpacingAfter) || double.IsInfinity(rule.SpacingAfter)) {
                throw new ArgumentException("Horizontal rule spacing after must be a non-negative finite value.");
            }
        }

        void ValidatePanelStyle(PanelStyle style, double panelWidth) {
            Guard.LeftCenterRightAlign(style.Align, nameof(style.Align), "Panel box");

            if (style.BorderWidth < 0 || double.IsNaN(style.BorderWidth) || double.IsInfinity(style.BorderWidth)) {
                throw new ArgumentException("Panel border width must be a non-negative finite value.");
            }

            if (style.PaddingX < 0 || double.IsNaN(style.PaddingX) || double.IsInfinity(style.PaddingX)) {
                throw new ArgumentException("Panel horizontal padding must be a non-negative finite value.");
            }

            if (style.PaddingY < 0 || double.IsNaN(style.PaddingY) || double.IsInfinity(style.PaddingY)) {
                throw new ArgumentException("Panel vertical padding must be a non-negative finite value.");
            }

            if (style.MaxWidth.HasValue && (style.MaxWidth.Value <= 0 || double.IsNaN(style.MaxWidth.Value) || double.IsInfinity(style.MaxWidth.Value))) {
                throw new ArgumentException("Panel maximum width must be a positive finite value.");
            }

            if (style.SpacingBefore < 0 || double.IsNaN(style.SpacingBefore) || double.IsInfinity(style.SpacingBefore)) {
                throw new ArgumentException("Panel spacing before must be a non-negative finite value.");
            }

            if (style.SpacingAfter < 0 || double.IsNaN(style.SpacingAfter) || double.IsInfinity(style.SpacingAfter)) {
                throw new ArgumentException("Panel spacing after must be a non-negative finite value.");
            }

            if (panelWidth - 2 * style.PaddingX <= 0) {
                throw new ArgumentException("Panel horizontal padding must leave a positive text width.");
            }
        }

        void MarkRichFonts(System.Collections.Generic.IEnumerable<TextRun> runs) {
            foreach (TextRun run in runs) {
                PdfStandardFont runBaseFont = ChooseNormal(run.Font ?? currentOpts.DefaultFont);
                PdfStandardFont runFont = run.Bold && run.Italic
                    ? ChooseBoldItalic(runBaseFont)
                    : run.Bold
                        ? ChooseBold(runBaseFont)
                        : run.Italic
                            ? ChooseItalic(runBaseFont)
                            : runBaseFont;
                currentPage!.UsedFonts.Add(runFont);
            }

            if (runs.Any(r => r.Bold)) { currentPage!.UsedBold = true; usedBold = true; }
            if (runs.Any(r => r.Italic)) { currentPage!.UsedItalic = true; usedItalic = true; }
            if (runs.Any(r => r.Bold && r.Italic)) { currentPage!.UsedBoldItalic = true; usedBoldItalic = true; }
        }

        void RenderListItem(System.Collections.Generic.IReadOnlyList<TextRun> runs, System.Collections.Generic.List<System.Collections.Generic.List<RichSeg>> lines, System.Collections.Generic.List<double> lineHeights, string marker, double markerX, double markerWidth, PdfAlign markerAlign, double textX, double textWidth, PdfAlign textAlign, PdfColor? color, double size, double leading, double spacingBefore, double spacingAfter, string? bookmarkName) {
            int lineIndex = 0;
            bool firstSegment = true;
            var listFont = ChooseNormal(currentOpts.DefaultFont);
            spacingBefore = ResolveTopLevelSpacingBefore(spacingBefore);
            if (spacingBefore > 0) {
                if (y - spacingBefore < currentOpts.MarginBottom) {
                    NewPage();
                    spacingBefore = 0D;
                }

                if (spacingBefore > 0) y -= spacingBefore;
            }

            while (lineIndex < lines.Count) {
                double available = y - currentOpts.MarginBottom;
                double firstLineHeight = GetRichLineHeight(lineHeights, lineIndex, leading);
                if (available < firstLineHeight) {
                    NewPage();
                    available = y - currentOpts.MarginBottom;
                    if (available < firstLineHeight) {
                        break;
                    }
                }

                int take = 0;
                double heightSum = 0;
                for (int k = lineIndex; k < lines.Count; k++) {
                    double lineHeight = GetRichLineHeight(lineHeights, k, leading);
                    if (heightSum + lineHeight > available) {
                        break;
                    }

                    heightSum += lineHeight;
                    take++;
                }

                if (take == 0) {
                    NewPage();
                    continue;
                }

                var segmentLines = new System.Collections.Generic.List<System.Collections.Generic.List<RichSeg>>(take);
                var segmentHeights = new System.Collections.Generic.List<double>(take);
                for (int k = 0; k < take; k++) {
                    segmentLines.Add(lines[lineIndex + k]);
                    segmentHeights.Add(GetRichLineHeight(lineHeights, lineIndex + k, leading));
                }

                double baselineY = FirstTextBaselineFromTop(listFont, size, y);
                if (firstSegment) {
                    if (!string.IsNullOrEmpty(bookmarkName)) {
                        AddNamedDestinationName(bookmarkName!, y);
                    }

                    var markerLines = new System.Collections.Generic.List<string>(1) { marker };
                    WriteLinesInternal("F1", size, leading, markerX, markerWidth, baselineY, markerLines, markerAlign, color, applyBaselineTweak: true);
                }

                pageDirty = true;
                WriteRichParagraph(sb, new RichParagraphBlock(runs, textAlign, color), segmentLines, segmentHeights, currentOpts, baselineY, size, leading, currentPage!.Annotations, textX, textWidth);
                MarkRichFonts(runs);
                y -= heightSum;
                lineIndex += take;
                firstSegment = false;
                if (lineIndex < lines.Count) {
                    NewPage();
                } else {
                    y -= spacingAfter;
                }
            }
        }

        double MeasureListKeepTogetherHeight(System.Collections.Generic.IReadOnlyList<TableCellTextLayout> itemLayouts, double leading, double spacingBefore, double itemSpacing, double spacingAfter) {
            double total = 0D;
            for (int itemIndex = 0; itemIndex < itemLayouts.Count; itemIndex++) {
                total += itemIndex == 0 ? spacingBefore : 0D;
                total += MeasureRichLinesHeight(itemLayouts[itemIndex].LineHeights, itemLayouts[itemIndex].LineCount, leading);
                total += itemIndex == itemLayouts.Count - 1 ? spacingAfter : itemSpacing;
            }

            return total;
        }

        PdfParagraphStyle? EffectiveParagraphStyle(RichParagraphBlock paragraph) => paragraph.Style ?? currentOpts.DefaultParagraphStyleSnapshot;

        double MeasureNextParagraphFirstLineHeight(RichParagraphBlock paragraph, double frameX, double frameWidth, double fontSize) {
            PdfParagraphStyle? paragraphStyle = EffectiveParagraphStyle(paragraph);
            double leading = GetParagraphLeading(paragraphStyle, fontSize);
            double spacingBefore = GetParagraphSpacingBefore(paragraphStyle);
            var textFrame = GetParagraphTextFrame(paragraphStyle, frameX, frameWidth);
            var wrap = WrapRichRuns(paragraph.Runs, textFrame.Width, fontSize, ChooseNormal(currentOpts.DefaultFont), leading, textFrame.FirstLineWidth, GetParagraphTabStopWidth(paragraphStyle));
            return wrap.LineHeights.Count == 0 ? spacingBefore : spacingBefore + wrap.LineHeights[0];
        }

        double MeasureNextBlockFirstVisualHeight(IPdfBlock block, double frameX, double frameWidth, double fontSize) {
            if (block is RichParagraphBlock paragraph) {
                return MeasureNextParagraphFirstLineHeight(paragraph, frameX, frameWidth, fontSize);
            }

            if (block is HeadingBlock heading) {
                PdfHeadingStyle? headingStyle = ResolveHeadingStyle(heading, currentOpts);
                double headingSize = GetHeadingFontSize(heading, headingStyle);
                double headingLeading = GetHeadingLeading(headingStyle, headingSize);
                return (headingStyle?.SpacingBefore ?? 0D) + headingLeading;
            }

            if (block is SpacerBlock spacer) {
                return spacer.Height;
            }

            if (block is BulletListBlock bullets) {
                PdfListStyle? listStyle = ResolveListStyle(bullets, currentOpts);
                double size = GetListFontSize(listStyle, fontSize);
                double leading = GetListLeading(listStyle, size);
                string? firstItem = bullets.Items.Count > 0 ? bullets.Items[0] : null;
                if (firstItem == null) {
                    return listStyle?.SpacingBefore ?? 0D;
                }

                return (listStyle?.SpacingBefore ?? 0D) + leading;
            }

            if (block is NumberedListBlock numbered) {
                PdfListStyle? listStyle = ResolveListStyle(numbered, currentOpts);
                double size = GetListFontSize(listStyle, fontSize);
                double leading = GetListLeading(listStyle, size);
                string? firstItem = numbered.Items.Count > 0 ? numbered.Items[0] : null;
                if (firstItem == null) {
                    return listStyle?.SpacingBefore ?? 0D;
                }

                return (listStyle?.SpacingBefore ?? 0D) + leading;
            }

            if (block is PanelParagraphBlock panel) {
                PanelStyle panelStyle = ResolvePanelStyle(panel, currentOpts);
                double innerWidth = panelStyle.MaxWidth.HasValue ? Math.Min(frameWidth, panelStyle.MaxWidth.Value) : frameWidth;
                ValidatePanelStyle(panelStyle, innerWidth);
                double size = fontSize;
                double leading = size * 1.4;
                double textWidth = innerWidth - 2 * panelStyle.PaddingX;
                var wrap = WrapRichRuns(panel.Runs, textWidth, size, ChooseNormal(currentOpts.DefaultFont), leading);
                double firstLineHeight = wrap.LineHeights.Count == 0 ? 0D : wrap.LineHeights[0];
                return panelStyle.SpacingBefore + panelStyle.PaddingY + firstLineHeight + panelStyle.PaddingY;
            }

            if (block is TableBlock table) {
                PdfTableStyle style = table.Style ?? currentOpts.DefaultTableStyleSnapshot ?? TableStyles.Light();
                int columns = GetTableColumnCount(table);
                if (columns == 0) {
                    return style.SpacingBefore;
                }

                double padLeft = GetTableCellPaddingLeft(style);
                double padRight = GetTableCellPaddingRight(style);
                double padTop = GetTableCellPaddingTop(style);
                double padBottom = GetTableCellPaddingBottom(style);
                double columnGap = GetTableCellSpacing(style);
                ValidateTableRoleRowCounts(style, table.Rows.Count);
                int headerRowCount = style.HeaderRowCount;
                int footerRowCount = style.FooterRowCount;
                int footerStartRowIndex = table.Rows.Count - footerRowCount;
                ValidateTableCellStyleCoordinates(style, table.Rows.Count, columns);
                ValidateTableColumnStyleBounds(style, columns);
                ValidateTableRowStyleBounds(style, table.Rows.Count);
                ValidateTableRowSpansWithinRoleBoundaries(table, columns, headerRowCount, footerStartRowIndex);
                double tableFontSize = GetTableBodyFontSize(style, fontSize);
                TableColumnLayout columnLayout = ResolveTableColumnLayout(table, currentOpts, style, columns, frameWidth, tableFontSize, headerRowCount, footerStartRowIndex);
                double tableWidth = columnLayout.Width;
                double rowSize = GetTableRowFontSize(style, 0, headerRowCount, footerStartRowIndex, fontSize);
                double rowLeading = GetTableLeading(style, rowSize);
                bool rowUsesBold = GetTableRowBold(style, 0, headerRowCount, footerStartRowIndex);
                int maxLines = 1;
                var firstRowCells = GetTableCellLayouts(table, 0, columns);
                for (int cellIndex = 0; cellIndex < firstRowCells.Count; cellIndex++) {
                    TableCellLayout cell = firstRowCells[cellIndex];
                    double cellWidth = GetTableCellWidth(columnLayout.Widths, cell.Column, cell.ColumnSpan, columnGap);
                    double innerWidth = Math.Max(1D, cellWidth - GetTableCellPaddingLeft(style, 0, cell.Column) - GetTableCellPaddingRight(style, 0, cell.Column));
                    var lines = WrapSimpleText(cell.Text, innerWidth, GetTableRowFont(currentOpts, rowUsesBold), rowSize);
                    maxLines = Math.Max(maxLines, lines.Count);
                }

                    double firstRowHeight = Math.Max(maxLines * rowLeading + GetTableRowMaxPaddingTop(table, style, 0, columns) + GetTableRowMaxPaddingBottom(table, style, 0, columns), GetTableRowMinHeight(style, 0));
                double captionHeight = 0D;
                if (!string.IsNullOrWhiteSpace(style.Caption)) {
                    double captionSize = style.CaptionFontSize ?? fontSize;
                    double captionLeading = captionSize * 1.25D;
                    var captionLines = WrapSimpleText(style.Caption!, tableWidth, ChooseNormal(currentOpts.DefaultFont), captionSize);
                    captionHeight = captionLines.Count * captionLeading + style.CaptionSpacingAfter;
                }

                return style.SpacingBefore + captionHeight + firstRowHeight;
            }

            if (block is HorizontalRuleBlock rule) {
                PdfHorizontalRuleStyle style = ResolveHorizontalRuleStyle(rule, currentOpts);
                return style.SpacingBefore + style.Thickness + style.SpacingAfter;
            }

            if (block is TextFieldBlock textField) {
                return textField.SpacingBefore + textField.Height + textField.SpacingAfter;
            }

            if (block is CheckBoxBlock checkBox) {
                return checkBox.SpacingBefore + checkBox.Size + checkBox.SpacingAfter;
            }

            if (block is ChoiceFieldBlock choiceField) {
                return choiceField.SpacingBefore + choiceField.Height + choiceField.SpacingAfter;
            }

            if (block is RadioButtonGroupBlock radioButtonGroup) {
                return radioButtonGroup.SpacingBefore + radioButtonGroup.Height + radioButtonGroup.SpacingAfter;
            }

            if (block is ImageBlock image) {
                PdfImageStyle style = ResolveImageStyle(image, currentOpts);
                return style.SpacingBefore + image.Height + style.SpacingAfter;
            }

            if (block is ShapeBlock shape) {
                PdfDrawingStyle style = ResolveDrawingStyle(shape, currentOpts);
                return style.SpacingBefore + shape.Shape.Height + style.SpacingAfter;
            }

            if (block is DrawingBlock drawing) {
                PdfDrawingStyle style = ResolveDrawingStyle(drawing, currentOpts);
                return style.SpacingBefore + drawing.Drawing.Height + style.SpacingAfter;
            }

            if (block is RowBlock row) {
                int columns = row.Columns.Count;
                if (columns == 0) {
                    return 0D;
                }

                PdfRowStyle? rowStyle = row.StyleSnapshot ?? currentOpts.DefaultRowStyleSnapshot;
                    double rowGap = row.GapOverride ?? rowStyle?.Gap ?? PdfRowStyle.DefaultGap;
                double totalGap = rowGap * Math.Max(0, columns - 1);
                if (totalGap >= frameWidth) {
                    return rowStyle?.SpacingBefore ?? 0D;
                }

                double columnAreaWidth = frameWidth - totalGap;
                double tallestFirstVisual = 0D;
                for (int columnIndex = 0; columnIndex < columns; columnIndex++) {
                    RowColumn column = row.Columns[columnIndex];
                    if (column.Blocks.Count == 0) {
                        continue;
                    }

                    double columnWidth = Math.Max(0D, columnAreaWidth * (column.WidthPercent / 100D));
                    tallestFirstVisual = Math.Max(tallestFirstVisual, MeasureNextBlockFirstVisualHeight(column.Blocks[0], frameX, columnWidth, fontSize));
                }

                return (rowStyle?.SpacingBefore ?? 0D) + tallestFirstVisual;
            }

            return 0D;
        }

        void ConsumeSpacer(double height) {
            double remaining = height;
            while (remaining > 0.001D) {
                double available = y - currentOpts.MarginBottom;
                if (available <= 0.5D) {
                    NewPage();
                    continue;
                }

                double consumed = Math.Min(remaining, available);
                y -= consumed;
                remaining -= consumed;
                if (remaining > 0.001D) {
                    NewPage();
                }
            }
        }

        void ProcessBlocks(System.Collections.Generic.IEnumerable<IPdfBlock> sequence) {
            var blockList = sequence as System.Collections.Generic.IList<IPdfBlock> ?? sequence.ToList();
            for (int blockIndex = 0; blockIndex < blockList.Count; blockIndex++) {
                var block = blockList[blockIndex];
                IPdfBlock? nextBlock = blockIndex + 1 < blockList.Count ? blockList[blockIndex + 1] : null;
                if (block is PageBlock pageBlock) {
                    FlushPage(pageDirty || HasCurrentPageNonContentObjects());
                    optionsStack.Push(pageBlock.Options);
                    pageGroupStack.Push(currentPageGroupId);
                    currentOpts = pageBlock.Options;
                    currentPageGroupId = nextPageGroupId++;
                    currentPage = null;
                    StartPage(currentOpts);
                    ProcessBlocks(pageBlock.Blocks);
                    FlushPage(force: true);
                    optionsStack.Pop();
                    currentPageGroupId = pageGroupStack.Pop();
                    currentOpts = optionsStack.Peek();
                    currentPage = null;
                    continue;
                }

                EnsurePage();

                if (block is PageBreakBlock) { NewPage(); continue; }
                if (block is BookmarkBlock bookmark) { AddNamedDestination(bookmark, y); continue; }
                if (block is SpacerBlock spacer) { ConsumeSpacer(spacer.Height); continue; }
                if (block is HeadingBlock hb) {
                    PdfHeadingStyle? headingStyle = ResolveHeadingStyle(hb, currentOpts);
                    double size = GetHeadingFontSize(hb, headingStyle);
                    double leading = GetHeadingLeading(headingStyle, size);
                    double spacingBefore = (y < yStart - 0.001 || headingStyle?.ApplySpacingBeforeAtTop == true) ? headingStyle?.SpacingBefore ?? 0D : 0D;
                    double spacingAfter = GetHeadingSpacingAfter(headingStyle, leading);
                    var headingFont = GetHeadingFont(currentOpts, headingStyle);
                    var lines = WrapSimpleText(hb.Text, width, headingFont, size);
                    double needed = spacingBefore + lines.Count * leading + spacingAfter;
                    bool keepWithNext = headingStyle?.KeepWithNext ?? true;
                    if (keepWithNext && nextBlock != null) {
                        double keepHeight = needed + MeasureNextBlockFirstVisualHeight(nextBlock, currentOpts.MarginLeft, width, currentOpts.DefaultFontSize);
                        double availableHeight = currentOpts.PageHeight - currentOpts.MarginTop - currentOpts.MarginBottom;
                        if (keepHeight > needed + 0.001 && keepHeight <= availableHeight + 0.001 && y < yStart - 0.001 && y - keepHeight < currentOpts.MarginBottom) {
                            NewPage();
                            spacingBefore = headingStyle?.ApplySpacingBeforeAtTop == true ? headingStyle.SpacingBefore : 0D;
                            needed = spacingBefore + lines.Count * leading + spacingAfter;
                        }
                    }

                    if (y - needed < currentOpts.MarginBottom) {
                        NewPage();
                        spacingBefore = headingStyle?.ApplySpacingBeforeAtTop == true ? headingStyle.SpacingBefore : 0D;
                        needed = spacingBefore + lines.Count * leading + spacingAfter;
                    }
                    if (spacingBefore > 0) {
                        y -= spacingBefore;
                    }

                    if (currentOpts.CreateOutlineFromHeadings) {
                        currentPage!.Bookmarks.Add(new PageBookmark { Level = hb.Level, Title = hb.Text, Y = y });
                    }
                    double firstBaseline = FirstTextBaselineFromTop(headingFont, size, y);
                    AddHeadingLinkAnnotations(hb, lines, headingFont, size, leading, currentOpts.MarginLeft, width, firstBaseline);
                    string headingFontResource = GetHeadingFontResource(headingStyle);
                    WriteLines(headingFontResource, size, leading, currentOpts.MarginLeft, firstBaseline, lines, hb.Align, hb.Color ?? headingStyle?.Color, applyBaselineTweak: false);
                    if (GetHeadingBold(headingStyle)) {
                        currentPage!.UsedBold = true;
                        usedBold = true;
                    }
                    y -= lines.Count * leading + spacingAfter;
                } else if (block is RichParagraphBlock rpb) {
                    double size = currentOpts.DefaultFontSize;
                    PdfParagraphStyle? paragraphStyle = EffectiveParagraphStyle(rpb);
                    double leading = GetParagraphLeading(paragraphStyle, size);
                    double spacingBefore = GetParagraphSpacingBefore(paragraphStyle);
                    double spacingAfter = GetParagraphSpacingAfter(paragraphStyle, leading);
                    var textFrame = GetParagraphTextFrame(paragraphStyle, currentOpts.MarginLeft, width);
                    var (lines, lineHeights) = WrapRichRuns(rpb.Runs, textFrame.Width, size, ChooseNormal(currentOpts.DefaultFont), leading, textFrame.FirstLineWidth, GetParagraphTabStopWidth(paragraphStyle));
                    if (paragraphStyle?.KeepWithNext == true && nextBlock != null && lines.Count > 0) {
                        double nextHeight = MeasureNextBlockFirstVisualHeight(nextBlock, currentOpts.MarginLeft, width, size);
                        double keepHeight = spacingBefore + lineHeights.Sum() + spacingAfter + nextHeight;
                        double availableHeight = currentOpts.PageHeight - currentOpts.MarginTop - currentOpts.MarginBottom;
                        if (nextHeight > 0.001 && keepHeight <= availableHeight + 0.001 && y < yStart - 0.001 && y - keepHeight < currentOpts.MarginBottom) {
                            NewPage();
                        }
                    }

                    if (paragraphStyle?.KeepTogether == true) {
                        double paragraphHeight = spacingBefore + lineHeights.Sum();
                        double availableHeight = currentOpts.PageHeight - currentOpts.MarginTop - currentOpts.MarginBottom;
                        if (paragraphHeight > availableHeight + 0.001) {
                            throw new ArgumentException("Paragraph height exceeds the available page content height.");
                        }

                        if (y < yStart - 0.001 && y - paragraphHeight < currentOpts.MarginBottom) {
                            NewPage();
                        }
                    }

                    int lineIndex = 0;
                    bool firstSegment = true;
                    while (lineIndex < lines.Count) {
                        double available = y - currentOpts.MarginBottom;
                        if (available <= 0.5) {
                            NewPage();
                            firstSegment = false;
                            continue;
                        }

                        double segmentSpacingBefore = firstSegment && y < yStart - 0.001 ? spacingBefore : 0;
                        double minimumLineHeight = lineHeights[lineIndex];
                        if (available < segmentSpacingBefore + minimumLineHeight) {
                            NewPage();
                            available = y - currentOpts.MarginBottom;
                            if (y >= yStart - 0.001) {
                                segmentSpacingBefore = 0;
                            }
                            if (available < segmentSpacingBefore + minimumLineHeight) {
                                segmentSpacingBefore = Math.Max(0, available - minimumLineHeight);
                            }
                        }

                        double roomForText = Math.Max(0, available - segmentSpacingBefore);
                        int take = 0;
                        double heightSum = 0;
                        for (int k = lineIndex; k < lines.Count; k++) {
                            double lineHeight = lineHeights[k];
                            if (heightSum + lineHeight > roomForText) {
                                break;
                            }

                            heightSum += lineHeight;
                            take++;
                        }

                        if (TryApplyWidowControl(paragraphStyle, lines.Count, lineIndex, ref take, ref heightSum, lineHeights, y < yStart - 0.001)) {
                            NewPage();
                            firstSegment = false;
                            continue;
                        }

                        if (take == 0) {
                            NewPage();
                            firstSegment = false;
                            continue;
                        }

                        if (segmentSpacingBefore > 0) {
                            y -= segmentSpacingBefore;
                        }

                        var sliceLines = new System.Collections.Generic.List<System.Collections.Generic.List<RichSeg>>();
                        var sliceHeights = new System.Collections.Generic.List<double>();
                        for (int k = 0; k < take; k++) {
                            sliceLines.Add(lines[lineIndex + k]);
                            sliceHeights.Add(lineHeights[lineIndex + k]);
                        }

                        bool sliceStartsAtFirstLine = lineIndex == 0;
                        pageDirty = true;
                        var paragraphFont = ChooseNormal(currentOpts.DefaultFont);
                        WriteRichParagraph(sb, rpb, sliceLines, sliceHeights, currentOpts, FirstTextBaselineFromTop(paragraphFont, size, y), size, leading, currentPage!.Annotations, textFrame.X, textFrame.Width, sliceStartsAtFirstLine ? textFrame.FirstLineX : null, sliceStartsAtFirstLine ? textFrame.FirstLineWidth : null);
                        y -= heightSum;
                        lineIndex += take;
                        firstSegment = false;
                        if (lineIndex < lines.Count) {
                            NewPage();
                        } else {
                            y -= spacingAfter;
                        }
                    }

                    MarkRichFonts(rpb.Runs);
                } else if (block is BulletListBlock bl) {
                    PdfListStyle? listStyle = ResolveListStyle(bl, currentOpts);
                    double size = GetListFontSize(listStyle, currentOpts.DefaultFontSize);
                    double leading = GetListLeading(listStyle, size);
                    var baseFont = ChooseNormal(currentOpts.DefaultFont);
                    const string bulletGlyph = "•";
                    double bulletWidth = bl.RichItems.Count == 0
                        ? EstimateSimpleTextWidth(bulletGlyph, baseFont, size)
                        : bl.RichItems.Max(item => EstimateSimpleTextWidth(item.Marker ?? bulletGlyph, baseFont, size));
                    double spaceAdvance = EstimateSimpleTextWidth(" ", baseFont, size);
                    double markerGap = GetListMarkerGap(listStyle, spaceAdvance);
                    double indent = bulletWidth + markerGap;
                    double listLeftIndent = listStyle?.LeftIndent ?? 0D;
                    double rawTextWidth = width - listLeftIndent - indent;
                    double availableWidth = Math.Max(rawTextWidth, EstimateSimpleTextWidth("WW", baseFont, size));
                    double alignmentWidth = Math.Max(0, rawTextWidth);
                    double itemSpacing = GetListItemSpacing(listStyle, leading);
                    var wrappedItems = new System.Collections.Generic.List<TableCellTextLayout>(bl.RichItems.Count);
                    for (int itemIndex = 0; itemIndex < bl.RichItems.Count; itemIndex++) {
                        wrappedItems.Add(CreateListItemTextLayout(bl.RichItems[itemIndex], availableWidth, baseFont, size, leading));
                    }

                    double listSpacingBefore = ResolveTopLevelSpacingBefore(listStyle?.SpacingBefore ?? 0D);
                    double listSpacingAfter = listStyle?.GetSpacingAfter(itemSpacing) ?? itemSpacing;
                    double listHeight = MeasureListKeepTogetherHeight(wrappedItems, leading, listSpacingBefore, itemSpacing, listSpacingAfter);
                    if (listStyle?.KeepTogether == true) {
                        double availableHeight = currentOpts.PageHeight - currentOpts.MarginTop - currentOpts.MarginBottom;
                        if (listHeight > availableHeight + 0.001) {
                            throw new ArgumentException("List height exceeds the available page content height.");
                        }

                        if (y < yStart - 0.001 && y - listHeight < currentOpts.MarginBottom) {
                            NewPage();
                            listSpacingBefore = 0D;
                            listHeight = MeasureListKeepTogetherHeight(wrappedItems, leading, listSpacingBefore, itemSpacing, listSpacingAfter);
                        }
                    }

                    if (listStyle?.KeepWithNext == true && nextBlock != null && wrappedItems.Count > 0) {
                        double nextHeight = MeasureNextBlockFirstVisualHeight(nextBlock, currentOpts.MarginLeft, width, size);
                        double keepHeight = listHeight + nextHeight;
                        double availableHeight = currentOpts.PageHeight - currentOpts.MarginTop - currentOpts.MarginBottom;
                        if (nextHeight > 0.001 && keepHeight <= availableHeight + 0.001 && y < yStart - 0.001 && y - keepHeight < currentOpts.MarginBottom) {
                            NewPage();
                            listSpacingBefore = 0D;
                            listHeight = MeasureListKeepTogetherHeight(wrappedItems, leading, listSpacingBefore, itemSpacing, listSpacingAfter);
                        }
                    }

                    for (int itemIndex = 0; itemIndex < bl.RichItems.Count; itemIndex++) {
                        var item = bl.RichItems[itemIndex];
                        string marker = item.Marker ?? bulletGlyph;
                        var layout = wrappedItems[itemIndex];
                        double firstLineWidth = layout.Lines.Count > 0 ? MeasureRichLineWidth(layout.Lines[0]) : 0;
                        double firstLineDx = 0;
                        if (bl.Align == PdfAlign.Center) firstLineDx = Math.Max(0, (alignmentWidth - firstLineWidth) / 2);
                        else if (bl.Align == PdfAlign.Right) firstLineDx = Math.Max(0, alignmentWidth - firstLineWidth);

                        double spacingBefore = itemIndex == 0 ? listSpacingBefore : 0D;
                        double spacingAfter = itemIndex == bl.RichItems.Count - 1 ? listSpacingAfter : itemSpacing;
                        RenderListItem(item.Runs, layout.Lines, layout.LineHeights, marker, currentOpts.MarginLeft + listLeftIndent + firstLineDx, bulletWidth, PdfAlign.Left, currentOpts.MarginLeft + listLeftIndent + indent, alignmentWidth, bl.Align, bl.Color ?? listStyle?.Color, size, leading, spacingBefore, spacingAfter, item.BookmarkName);
                    }
                } else if (block is NumberedListBlock nl) {
                    PdfListStyle? listStyle = ResolveListStyle(nl, currentOpts);
                    double size = GetListFontSize(listStyle, currentOpts.DefaultFontSize);
                    double leading = GetListLeading(listStyle, size);
                    var baseFont = ChooseNormal(currentOpts.DefaultFont);
                    int lastNumber = nl.StartNumber + Math.Max(0, nl.RichItems.Count - 1);
                    string widestMarker = lastNumber.ToString(CultureInfo.InvariantCulture) + ".";
                    double markerWidth = nl.RichItems.Count == 0
                        ? EstimateSimpleTextWidth(widestMarker, baseFont, size)
                        : nl.RichItems
                            .Select((item, itemIndex) => item.Marker ?? ((nl.StartNumber + itemIndex).ToString(CultureInfo.InvariantCulture) + "."))
                            .Max(marker => EstimateSimpleTextWidth(marker, baseFont, size));
                    double spaceAdvance = EstimateSimpleTextWidth(" ", baseFont, size);
                    double markerGap = GetListMarkerGap(listStyle, spaceAdvance);
                    double indent = markerWidth + markerGap;
                    double rawTextWidth = width - (listStyle?.LeftIndent ?? 0D) - indent;
                    double availableWidth = Math.Max(rawTextWidth, EstimateSimpleTextWidth("WW", baseFont, size));
                    double alignmentWidth = Math.Max(0, rawTextWidth);
                    double itemSpacing = GetListItemSpacing(listStyle, leading);
                    double listLeftIndent = listStyle?.LeftIndent ?? 0D;
                    var wrappedItems = new System.Collections.Generic.List<TableCellTextLayout>(nl.RichItems.Count);
                    for (int itemIndex = 0; itemIndex < nl.RichItems.Count; itemIndex++) {
                        wrappedItems.Add(CreateListItemTextLayout(nl.RichItems[itemIndex], availableWidth, baseFont, size, leading));
                    }

                    double listSpacingBefore = ResolveTopLevelSpacingBefore(listStyle?.SpacingBefore ?? 0D);
                    double listSpacingAfter = listStyle?.GetSpacingAfter(itemSpacing) ?? itemSpacing;
                    double listHeight = MeasureListKeepTogetherHeight(wrappedItems, leading, listSpacingBefore, itemSpacing, listSpacingAfter);
                    if (listStyle?.KeepTogether == true) {
                        double availableHeight = currentOpts.PageHeight - currentOpts.MarginTop - currentOpts.MarginBottom;
                        if (listHeight > availableHeight + 0.001) {
                            throw new ArgumentException("List height exceeds the available page content height.");
                        }

                        if (y < yStart - 0.001 && y - listHeight < currentOpts.MarginBottom) {
                            NewPage();
                            listSpacingBefore = 0D;
                            listHeight = MeasureListKeepTogetherHeight(wrappedItems, leading, listSpacingBefore, itemSpacing, listSpacingAfter);
                        }
                    }

                    if (listStyle?.KeepWithNext == true && nextBlock != null && wrappedItems.Count > 0) {
                        double nextHeight = MeasureNextBlockFirstVisualHeight(nextBlock, currentOpts.MarginLeft, width, size);
                        double keepHeight = listHeight + nextHeight;
                        double availableHeight = currentOpts.PageHeight - currentOpts.MarginTop - currentOpts.MarginBottom;
                        if (nextHeight > 0.001 && keepHeight <= availableHeight + 0.001 && y < yStart - 0.001 && y - keepHeight < currentOpts.MarginBottom) {
                            NewPage();
                            listSpacingBefore = 0D;
                            listHeight = MeasureListKeepTogetherHeight(wrappedItems, leading, listSpacingBefore, itemSpacing, listSpacingAfter);
                        }
                    }

                    for (int itemIndex = 0; itemIndex < nl.RichItems.Count; itemIndex++) {
                        var item = nl.RichItems[itemIndex];
                        string marker = item.Marker ?? ((nl.StartNumber + itemIndex).ToString(CultureInfo.InvariantCulture) + ".");
                        var layout = wrappedItems[itemIndex];
                        double firstLineWidth = layout.Lines.Count > 0 ? MeasureRichLineWidth(layout.Lines[0]) : 0;
                        double firstLineDx = 0;
                        if (nl.Align == PdfAlign.Center) firstLineDx = Math.Max(0, (alignmentWidth - firstLineWidth) / 2);
                        else if (nl.Align == PdfAlign.Right) firstLineDx = Math.Max(0, alignmentWidth - firstLineWidth);

                        double spacingBefore = itemIndex == 0 ? listSpacingBefore : 0D;
                        double spacingAfter = itemIndex == nl.RichItems.Count - 1 ? listSpacingAfter : itemSpacing;
                        RenderListItem(item.Runs, layout.Lines, layout.LineHeights, marker, currentOpts.MarginLeft + listLeftIndent + firstLineDx, markerWidth, PdfAlign.Right, currentOpts.MarginLeft + listLeftIndent + indent, alignmentWidth, nl.Align, nl.Color ?? listStyle?.Color, size, leading, spacingBefore, spacingAfter, item.BookmarkName);
                    }
                } else if (block is TableBlock tb) {
                    PdfTableStyle style = tb.Style ?? currentOpts.DefaultTableStyleSnapshot ?? TableStyles.Light();
                    int cols = GetTableColumnCount(tb);
                    if (cols == 0) continue;
                    double padLeft = GetTableCellPaddingLeft(style);
                    double padRight = GetTableCellPaddingRight(style);
                    double padTop = GetTableCellPaddingTop(style);
                    double padBottom = GetTableCellPaddingBottom(style);
                    double cellSpacing = GetTableCellSpacing(style);
                    double colGapPx = cellSpacing;
                    double rowGapPx = cellSpacing;
                    double size = GetTableBodyFontSize(style, currentOpts.DefaultFontSize);
                    if (!IsValidPdfAlign(tb.Align)) {
                        throw new ArgumentException("Table alignment must be Left, Center, or Right.");
                    }
                    if (style.Alignments != null) {
                        foreach (var alignment in style.Alignments) {
                            if (!IsValidPdfColumnAlign(alignment)) {
                                throw new ArgumentException("Table column alignments must be Left, Center, or Right.");
                            }
                        }
                    }
                    if (style.VerticalAlignments != null) {
                        foreach (var verticalAlignment in style.VerticalAlignments) {
                            if (!IsValidPdfCellVerticalAlign(verticalAlignment)) {
                                throw new ArgumentException("Table vertical alignments must be defined PDF cell vertical alignment values.");
                            }
                        }
                    }
                    if (!IsValidPdfAlign(style.CaptionAlign)) {
                        throw new ArgumentException("Table caption alignment must be Left, Center, or Right.");
                    }
                    if (style.BorderWidth < 0 || double.IsNaN(style.BorderWidth) || double.IsInfinity(style.BorderWidth)) {
                        throw new ArgumentException("Table border width must be a non-negative finite value.");
                    }
                    if (style.RowSeparatorWidth < 0 || double.IsNaN(style.RowSeparatorWidth) || double.IsInfinity(style.RowSeparatorWidth)) {
                        throw new ArgumentException("Table row separator width must be a non-negative finite value.");
                    }
                    if (style.HeaderSeparatorWidth < 0 || double.IsNaN(style.HeaderSeparatorWidth) || double.IsInfinity(style.HeaderSeparatorWidth)) {
                        throw new ArgumentException("Table header separator width must be a non-negative finite value.");
                    }
                    if (style.FooterSeparatorWidth < 0 || double.IsNaN(style.FooterSeparatorWidth) || double.IsInfinity(style.FooterSeparatorWidth)) {
                        throw new ArgumentException("Table footer separator width must be a non-negative finite value.");
                    }
                    if (style.CellPaddingX < 0 || double.IsNaN(style.CellPaddingX) || double.IsInfinity(style.CellPaddingX)) {
                        throw new ArgumentException("Table horizontal cell padding must be a non-negative finite value.");
                    }
                    if (style.CellPaddingY < 0 || double.IsNaN(style.CellPaddingY) || double.IsInfinity(style.CellPaddingY)) {
                        throw new ArgumentException("Table vertical cell padding must be a non-negative finite value.");
                    }
                    if (style.MinRowHeight < 0 || double.IsNaN(style.MinRowHeight) || double.IsInfinity(style.MinRowHeight)) {
                        throw new ArgumentException("Table minimum row height must be a non-negative finite value.");
                    }
                    if (style.RowMinHeights != null) {
                        foreach (double? rowMinHeight in style.RowMinHeights) {
                            if (rowMinHeight.HasValue && (rowMinHeight.Value < 0 || double.IsNaN(rowMinHeight.Value) || double.IsInfinity(rowMinHeight.Value))) {
                                throw new ArgumentException("Table row minimum heights must be non-negative finite values.");
                            }
                        }
                    }
                    if (style.SpacingBefore < 0 || double.IsNaN(style.SpacingBefore) || double.IsInfinity(style.SpacingBefore)) {
                        throw new ArgumentException("Table spacing before must be a non-negative finite value.");
                    }
                    if (style.Caption != null && string.IsNullOrWhiteSpace(style.Caption)) {
                        throw new ArgumentException("Table caption cannot be empty or whitespace.");
                    }
                    if (style.CaptionFontSize.HasValue && (style.CaptionFontSize.Value <= 0 || double.IsNaN(style.CaptionFontSize.Value) || double.IsInfinity(style.CaptionFontSize.Value))) {
                        throw new ArgumentException("Table caption font size must be a positive finite value.");
                    }
                    if (style.CaptionSpacingAfter < 0 || double.IsNaN(style.CaptionSpacingAfter) || double.IsInfinity(style.CaptionSpacingAfter)) {
                        throw new ArgumentException("Table caption spacing after must be a non-negative finite value.");
                    }
                    if (style.SpacingAfter < 0 || double.IsNaN(style.SpacingAfter) || double.IsInfinity(style.SpacingAfter)) {
                        throw new ArgumentException("Table spacing after must be a non-negative finite value.");
                    }
                    if (double.IsNaN(style.RowBaselineOffset) || double.IsInfinity(style.RowBaselineOffset)) {
                        throw new ArgumentException("Table row baseline offset must be a finite value.");
                    }
                    if (style.CellFills != null) {
                        foreach (var cellFill in style.CellFills) {
                            if (cellFill.Key.Row < 0 || cellFill.Key.Column < 0) {
                                throw new ArgumentException("Table cell fill coordinates cannot be negative.");
                            }
                        }
                    }
                    if (style.CellBorders != null) {
                        foreach (var cellBorder in style.CellBorders) {
                            if (cellBorder.Key.Row < 0 || cellBorder.Key.Column < 0) {
                                throw new ArgumentException("Table cell border coordinates cannot be negative.");
                            }
                            if (cellBorder.Value == null || cellBorder.Value.Width < 0 || double.IsNaN(cellBorder.Value.Width) || double.IsInfinity(cellBorder.Value.Width)) {
                                throw new ArgumentException("Table cell border widths must be non-negative finite values.");
                            }
                        }
                    }
                    if (style.HeaderRowCount < 0) {
                        throw new ArgumentException("Table header row count cannot be negative.");
                    }
                    if (style.FooterRowCount < 0) {
                        throw new ArgumentException("Table footer row count cannot be negative.");
                    }

                    ValidateTableRoleRowCounts(style, tb.Rows.Count);
                    int headerRowCount = style.HeaderRowCount;
                    int repeatHeaderRowCount = GetTableRepeatHeaderRowCount(style);
                    int footerRowCount = style.FooterRowCount;
                    int footerStartRowIndex = tb.Rows.Count - footerRowCount;
                    ValidateTableCellStyleCoordinates(style, tb.Rows.Count, cols);
                    ValidateTableColumnStyleBounds(style, cols);
                    ValidateTableRowStyleBounds(style, tb.Rows.Count);
                    ValidateTableRowSpansWithinRoleBoundaries(tb, cols, headerRowCount, footerStartRowIndex);
                    double[]? autoFitWeights = style.AutoFitColumns
                        ? MeasureAutoFitColumnWeights(tb, currentOpts, style, size, headerRowCount, footerStartRowIndex)
                        : null;
                    double[]? autoFitMinimumWidths = style.AutoFitColumns
                        ? MeasureAutoFitColumnMinimumWidths(tb, currentOpts, style, size, headerRowCount, footerStartRowIndex)
                        : null;
                    double contentWidth = currentOpts.PageWidth - currentOpts.MarginLeft - currentOpts.MarginRight;
                    double tableWidth = ResolveTableAvailableWidth(style, contentWidth);
                    double[] colPixel = new double[cols];
                    double[] colWeights = new double[cols];
                    bool[] fixedColumns = new bool[cols];
                    double?[] minWidths = new double?[cols];
                    double?[] maxWidths = new double?[cols];
                    double fixedWidthTotal = 0;
                    double totalWeight = 0;
                    for (int c = 0; c < cols; c++) {
                        double? minWidth = GetOptionalColumnWidth(style.ColumnMinWidthPoints, c, "Table minimum column widths must be positive finite values.");
                        if (!minWidth.HasValue && autoFitMinimumWidths != null && c < autoFitMinimumWidths.Length) {
                            minWidth = autoFitMinimumWidths[c];
                        }

                        double? maxWidth = GetOptionalColumnWidth(style.ColumnMaxWidthPoints, c, "Table maximum column widths must be positive finite values.");
                        if (minWidth.HasValue && maxWidth.HasValue && minWidth.Value > maxWidth.Value + 0.001) {
                            throw new ArgumentException("Table minimum column widths cannot be greater than maximum column widths.");
                        }

                        minWidths[c] = minWidth;
                        maxWidths[c] = maxWidth;

                        if (style.ColumnWidthPoints != null &&
                            c < style.ColumnWidthPoints.Count &&
                            style.ColumnWidthPoints[c].HasValue) {
                            double fixedWidth = style.ColumnWidthPoints[c]!.Value;
                            if (fixedWidth <= 0 || double.IsNaN(fixedWidth) || double.IsInfinity(fixedWidth)) {
                                throw new ArgumentException("Table fixed column widths must be positive finite values.");
                            }
                            if (minWidth.HasValue && fixedWidth < minWidth.Value - 0.001) {
                                throw new ArgumentException("Table fixed column widths cannot be smaller than configured minimum widths.");
                            }
                            if (maxWidth.HasValue && fixedWidth > maxWidth.Value + 0.001) {
                                throw new ArgumentException("Table fixed column widths cannot be greater than configured maximum widths.");
                            }

                            colPixel[c] = fixedWidth;
                            fixedColumns[c] = true;
                            fixedWidthTotal += fixedWidth;
                            continue;
                        }

                        double weight = 1;
                        if (style.ColumnWidthWeights != null && c < style.ColumnWidthWeights.Count) {
                            weight = style.ColumnWidthWeights[c];
                            if (weight <= 0 || double.IsNaN(weight) || double.IsInfinity(weight)) {
                                throw new ArgumentException("Table column width weights must be positive finite values.");
                            }
                        } else if (autoFitWeights != null && c < autoFitWeights.Length) {
                            weight = autoFitWeights[c];
                        }

                        colWeights[c] = weight;
                        totalWeight += weight;
                    }
                    double tableInnerWidth = tableWidth - (cols - 1) * colGapPx;
                    if (tableInnerWidth <= 0.001 || double.IsNaN(tableInnerWidth) || double.IsInfinity(tableInnerWidth)) {
                        throw new ArgumentException("Table cell spacing must leave a positive table width.");
                    }

                    fixedWidthTotal = FitFixedTableColumnsToAvailableWidth(colPixel, fixedColumns, minWidths, fixedWidthTotal, tableInnerWidth);

                    if (totalWeight <= 0) {
                        tableInnerWidth = fixedWidthTotal;
                        tableWidth = tableInnerWidth + (cols - 1) * colGapPx;
                    }

                    double remainingWidth = Math.Max(0, tableInnerWidth - fixedWidthTotal);
                    DistributeFlexibleColumns(colPixel, colWeights, fixedColumns, minWidths, maxWidths, remainingWidth);
                    double usedTableInnerWidth = colPixel.Sum();
                    if (usedTableInnerWidth < tableInnerWidth - 0.001) {
                        tableInnerWidth = usedTableInnerWidth;
                        tableWidth = tableInnerWidth + (cols - 1) * colGapPx;
                    }
                    ValidateTableCellTextWidths(tb, style, cols, colPixel, colGapPx);

                    var rowLines = new TableCellTextLayout[tb.Rows.Count][];
                    var rowLineCounts = new int[tb.Rows.Count];
                    var rowHeights = new double[tb.Rows.Count];
                    var rowLeadings = new double[tb.Rows.Count];
                    var rowSizes = new double[tb.Rows.Count];
                    var rowBold = new bool[tb.Rows.Count];
                    for (int ri = 0; ri < tb.Rows.Count; ri++) {
                        double rowSize = GetTableRowFontSize(style, ri, headerRowCount, footerStartRowIndex, currentOpts.DefaultFontSize);
                        double rowLeading = GetTableLeading(style, rowSize);
                        bool rowUsesBold = GetTableRowBold(style, ri, headerRowCount, footerStartRowIndex);
                        rowSizes[ri] = rowSize;
                        rowLeadings[ri] = rowLeading;
                        rowBold[ri] = rowUsesBold;
                        rowLines[ri] = new TableCellTextLayout[cols];
                        int maxLines = 1;
                        double maxRequiredHeight = rowLeading + GetTableRowMaxPaddingTop(tb, style, ri, cols) + GetTableRowMaxPaddingBottom(tb, style, ri, cols);
                        for (int ci = 0; ci < cols; ci++) {
                            rowLines[ri][ci] = new TableCellTextLayout(new System.Collections.Generic.List<System.Collections.Generic.List<RichSeg>> { new() }, new System.Collections.Generic.List<double> { rowLeading });
                        }

                        var cells = GetTableCellLayouts(tb, ri, cols);
                        for (int cellIndex = 0; cellIndex < cells.Count; cellIndex++) {
                            TableCellLayout cell = cells[cellIndex];
                            var cellFont = GetTableRowFont(currentOpts, rowUsesBold);
                            double cellWidth = GetTableCellWidth(colPixel, cell.Column, cell.ColumnSpan, colGapPx);
                            double innerWidth = Math.Max(1, cellWidth - GetTableCellPaddingLeft(style, ri, cell.Column) - GetTableCellPaddingRight(style, ri, cell.Column));
                            TableCellTextLayout lines = CreateTableCellTextLayout(cell, innerWidth, cellFont, rowSize, rowLeading);
                            rowLines[ri][cell.Column] = lines;
                            if (cell.RowSpan <= 1) {
                                maxLines = Math.Max(maxLines, lines.LineCount);
                                maxRequiredHeight = Math.Max(maxRequiredHeight, MeasureTableCellContentHeight(cell, lines, 0, lines.LineCount, rowLeading) + GetTableCellPaddingTop(style, ri, cell.Column) + GetTableCellPaddingBottom(style, ri, cell.Column));
                            }
                        }
                        rowLineCounts[ri] = maxLines;
                        rowHeights[ri] = Math.Max(maxRequiredHeight, GetTableRowMinHeight(style, ri));
                    }
                    ApplyTableRowSpanHeights(tb, style, cols, rowLines, rowHeights, rowLeadings, rowGapPx);
                    double xOrigin = ResolveTableX(tb.Align, style, currentOpts.MarginLeft, contentWidth, tableWidth);

                    double maxContentHeight = currentOpts.PageHeight - currentOpts.MarginTop - currentOpts.MarginBottom;
                    string? captionText = string.IsNullOrWhiteSpace(style.Caption) ? null : style.Caption;
                    System.Collections.Generic.List<string>? captionLines = null;
                    double captionSize = style.CaptionFontSize ?? size;
                    double captionLeading = captionSize * 1.25;
                    double captionHeight = 0;
                    if (captionText != null) {
                        var captionFontForWrap = ChooseNormal(currentOpts.DefaultFont);
                        captionLines = WrapSimpleText(captionText, tableWidth, captionFontForWrap, captionSize).ToList();
                        captionHeight = captionLines.Count * captionLeading;
                        double firstRowHeight = rowHeights.Length > 0 ? rowHeights[0] : 0;
                        if (captionHeight + style.CaptionSpacingAfter + firstRowHeight > maxContentHeight + 0.001) {
                            throw new ArgumentException("Table caption and first row exceed the available page content height.");
                        }
                    }

                    double tableContentHeight = (captionLines == null ? 0 : captionHeight + style.CaptionSpacingAfter) + GetTableRowsHeight(rowHeights, 0, rowHeights.Length, rowGapPx);
                    double tableSpacingBefore = y < yStart - 0.001 ? style.SpacingBefore : 0D;
                    if (style.KeepTogether) {
                        double keepHeight = tableSpacingBefore + tableContentHeight + style.SpacingAfter;
                        if (keepHeight > maxContentHeight + 0.001) {
                            throw new ArgumentException("Table height exceeds the available page content height.");
                        }

                        if (y < yStart - 0.001 && y - keepHeight < currentOpts.MarginBottom) {
                            NewPage();
                            tableSpacingBefore = 0D;
                        }
                    }

                    if (style.KeepWithNext && nextBlock != null) {
                        double tableHeight = tableSpacingBefore + tableContentHeight + style.SpacingAfter;
                        double nextHeight = MeasureNextBlockFirstVisualHeight(nextBlock, currentOpts.MarginLeft, width, currentOpts.DefaultFontSize);
                        double keepHeight = tableHeight + nextHeight;
                        if (nextHeight > 0.001 && tableHeight <= maxContentHeight + 0.001 && keepHeight <= maxContentHeight + 0.001 && y < yStart - 0.001 && y - keepHeight < currentOpts.MarginBottom) {
                            NewPage();
                            tableSpacingBefore = 0D;
                        }
                    }

                    if (tableSpacingBefore > 0) {
                        if (y < yStart - 0.001 && y - tableSpacingBefore < currentOpts.MarginBottom) {
                            NewPage();
                            tableSpacingBefore = 0D;
                        }

                        y -= tableSpacingBefore;
                    }

                    if (captionLines != null) {
                        var captionFont = ChooseNormal(currentOpts.DefaultFont);
                        double firstRowHeight = rowHeights.Length > 0 ? rowHeights[0] : 0;
                        double captionAndFirstRowHeight = captionHeight + style.CaptionSpacingAfter + firstRowHeight;
                        if (y < yStart - 0.001 &&
                            y - Math.Min(captionAndFirstRowHeight, maxContentHeight) < currentOpts.MarginBottom) {
                            NewPage();
                        }

                        WriteLinesInternal("F1", captionSize, captionLeading, xOrigin, tableWidth, y - GetAscender(captionFont, captionSize), captionLines, style.CaptionAlign, style.CaptionColor);
                        y -= captionHeight + style.CaptionSpacingAfter;
                    }

                    if (TableUsesBold(style, tb.Rows.Count, headerRowCount, footerStartRowIndex)) {
                        currentPage!.UsedBold = true;
                        usedBold = true;
                    }

                    bool hasRepeatableHeader = repeatHeaderRowCount > 0 && tb.Rows.Count > headerRowCount;
                    double repeatHeaderHeight = 0;
                    for (int i = 0; i < repeatHeaderRowCount; i++) {
                        repeatHeaderHeight += rowHeights[i] + GetTableRowGapAfter(i, tb.Rows.Count, rowGapPx);
                    }

                    bool ShouldBreakBefore(double rowHeight) =>
                        y < yStart - 0.001 &&
                        y - rowHeight < currentOpts.MarginBottom &&
                        rowHeight <= maxContentHeight;

                    bool CanRepeatHeaderWithSegment(int rowIndex) =>
                        hasRepeatableHeader &&
                        rowIndex >= headerRowCount &&
                        repeatHeaderHeight + rowLeadings[rowIndex] + GetTableRowMaxPaddingTop(tb, style, rowIndex, cols) + GetTableRowMaxPaddingBottom(tb, style, rowIndex, cols) <= maxContentHeight + 0.001;

                    void DrawRepeatHeaders() {
                        for (int headerIndex = 0; headerIndex < repeatHeaderRowCount; headerIndex++) {
                            DrawTableRow(headerIndex, renderAsHeader: true, suppressCellObjects: true);
                        }
                    }

                    void NewTablePage(int rowIndex) {
                        NewPage();
                        if (CanRepeatHeaderWithSegment(rowIndex)) {
                            DrawRepeatHeaders();
                        }
                    }

                    void DrawTableRowSegment(int rowIndex, bool renderAsHeader, int startLine, int lineCount, bool suppressCellObjects = false) {
                        bool renderAsFooter = rowIndex >= footerStartRowIndex;
                        bool rowUsesBold = rowBold[rowIndex];
                        double rowSize = rowSizes[rowIndex];
                        double rowLeading = rowLeadings[rowIndex];
                        if (rowUsesBold) {
                            currentPage!.UsedBold = true;
                            usedBold = true;
                        }

                        var cells = GetTableCellLayouts(tb, rowIndex, cols);
                        bool wholeRowSegment = startLine == 0 && lineCount == rowLineCounts[rowIndex];
                        double rowPadTop = GetTableRowMaxPaddingTop(tb, style, rowIndex, cols);
                        double rowPadBottom = GetTableRowMaxPaddingBottom(tb, style, rowIndex, cols);
                        double rowHeight = wholeRowSegment ? rowHeights[rowIndex] : Math.Max(1, lineCount) * rowLeading + rowPadTop + rowPadBottom;
                        double rowBottom = y - rowHeight;
                        if (currentOpts.Debug?.ShowTableRowBoxes == true) { pageDirty = true; DrawRowRect(sb, new PdfColor(1, 0, 1), 0.6, xOrigin, rowBottom, tableWidth, rowHeight); }
                        int bodyRowIndex = rowIndex - headerRowCount;
                        bool stripeBodyRow = bodyRowIndex >= 0 && bodyRowIndex % 2 == 1;
                        bool[] rowFillSkips = GetRowSpanContinuationSkipColumns(tb, rowIndex, cols);
                        if (style?.HeaderFill is not null && renderAsHeader) { pageDirty = true; DrawTableRowFill(sb, style.HeaderFill.Value, xOrigin, colPixel, colGapPx, rowBottom, rowHeight, rowFillSkips); } else if (style?.FooterFill is not null && renderAsFooter) { pageDirty = true; DrawTableRowFill(sb, style.FooterFill.Value, xOrigin, colPixel, colGapPx, rowBottom, rowHeight, rowFillSkips); } else if (!renderAsHeader && !renderAsFooter && style?.RowStripeFill is not null && stripeBodyRow) { pageDirty = true; DrawTableRowFill(sb, style.RowStripeFill.Value, xOrigin, colPixel, colGapPx, rowBottom, rowHeight, rowFillSkips); }
                        if (!renderAsHeader && !renderAsFooter && style?.BodyColumnFills != null) {
                            bool[] bodyColumnFillSkips = GetMergedCellContinuationSkipColumns(tb, rowIndex, cols);
                            double fillX = xOrigin;
                            for (int fillColumn = 0; fillColumn < cols; fillColumn++) {
                                PdfColor? fill = fillColumn < style.BodyColumnFills.Count ? style.BodyColumnFills[fillColumn] : null;
                                if (fill.HasValue && (fillColumn >= bodyColumnFillSkips.Length || !bodyColumnFillSkips[fillColumn])) {
                                    pageDirty = true;
                                    DrawRowFill(sb, fill.Value, fillX, rowBottom, colPixel[fillColumn], rowHeight);
                                }
                                fillX += colPixel[fillColumn] + colGapPx;
                            }
                        }
                        if (style?.CellFills != null && style.CellFills.Count > 0) {
                            double fillX = xOrigin;
                            for (int fillColumn = 0; fillColumn < cols; fillColumn++) {
                                if (style.CellFills.TryGetValue((rowIndex, fillColumn), out PdfColor fill) &&
                                    TryGetTableCellLayoutAtColumn(cells, fillColumn, out TableCellLayout fillCell) &&
                                    (fillColumn >= rowFillSkips.Length || !rowFillSkips[fillColumn])) {
                                    pageDirty = true;
                                    int span = wholeRowSegment ? fillCell.ColumnSpan : 1;
                                    double fillHeight = rowHeight;
                                    double fillBottom = rowBottom;
                                    if (wholeRowSegment) {
                                        if (fillCell.RowSpan > 1) {
                                            fillHeight = GetTableCellHeight(rowHeights, rowIndex, fillCell.RowSpan, rowGapPx);
                                            fillBottom = y - fillHeight;
                                        }
                                    }

                                    DrawRowFill(sb, fill, fillX, fillBottom, GetTableCellWidth(colPixel, fillColumn, span, colGapPx), fillHeight);
                                }
                                fillX += colPixel[fillColumn] + colGapPx;
                            }
                        }
                        if (style != null && DrawTableCellDataBars(sb, style, cells, rowIndex, cols, xOrigin, y, rowBottom, rowHeight, colPixel, colGapPx, rowHeights, rowGapPx, wholeRowSegment, startLine, rowFillSkips)) {
                            pageDirty = true;
                        }
                        if (style != null && DrawTableCellIcons(sb, style, cells, rowIndex, cols, xOrigin, y, rowBottom, rowHeight, colPixel, colGapPx, rowHeights, rowGapPx, wholeRowSegment, startLine, rowFillSkips)) {
                            pageDirty = true;
                        }
                        if (currentOpts.Debug?.ShowTableBaselines == true) {
                            double x1 = xOrigin;
                            double x2 = xOrigin + tableWidth;
                            double baselineYDbg = y - padTop - GetAscender(GetTableRowFont(currentOpts, rowUsesBold), rowSize);
                            pageDirty = true;
                            DrawHLine(sb, new PdfColor(0, 0.6, 0), 0.4, x1, x2, baselineYDbg);
                        }
                        double xi = xOrigin;
                        double yRect = rowBottom;
                        double rowWidth = tableWidth;
                        double hRect = rowHeight;
                        var textColor = renderAsHeader ? style!.HeaderTextColor : renderAsFooter ? style!.FooterTextColor : style!.TextColor;
                        for (int cellIndex = 0; cellIndex < cells.Count; cellIndex++) {
                            TableCellLayout cell = cells[cellIndex];
                            int c = cell.Column;
                            xi = xOrigin;
                            for (int xColumn = 0; xColumn < c; xColumn++) {
                                xi += colPixel[xColumn] + colGapPx;
                            }

                            double cellWidth = GetTableCellWidth(colPixel, c, cell.ColumnSpan, colGapPx);
                            double cellPadLeft = GetTableCellPaddingLeft(style, rowIndex, c);
                            double cellPadRight = GetTableCellPaddingRight(style, rowIndex, c);
                            double cellPadTop = GetTableCellPaddingTop(style, rowIndex, c);
                            double cellPadBottom = GetTableCellPaddingBottom(style, rowIndex, c);
                            double innerW = cellWidth - cellPadLeft - cellPadRight;
                            double cellHeight = wholeRowSegment && cell.RowSpan > 1 ? GetTableCellHeight(rowHeights, rowIndex, cell.RowSpan, rowGapPx) : rowHeight;
                            double cellBottom = y - cellHeight;
                            PdfColumnAlign align = GetTableCellAlignment(style, rowIndex, c, cell.Text);
                            PdfCellVerticalAlign verticalAlign = GetTableCellVerticalAlignment(style, rowIndex, c);

                            var cellFont = GetTableRowFont(currentOpts, rowUsesBold);
                            TableCellTextLayout lines = rowLines[rowIndex][c];
                            int sourceStartLine = wholeRowSegment && cell.RowSpan > 1 ? 0 : startLine;
                            int requestedLineCount = wholeRowSegment && cell.RowSpan > 1 ? lines.LineCount : lineCount;
                            int visibleLineCount = Math.Max(0, Math.Min(requestedLineCount, lines.LineCount - sourceStartLine));
                            double verticalOffset = 0;
                            double visibleTextHeight = 0D;
                            if (visibleLineCount > 0) {
                                double availableTextHeight = Math.Max(0, cellHeight - cellPadTop - cellPadBottom);
                                visibleTextHeight = MeasureTableCellTextHeight(lines, sourceStartLine, visibleLineCount, rowLeading);
                                double visibleContentHeight = MeasureTableCellContentHeight(cell, lines, sourceStartLine, visibleLineCount, rowLeading);
                                double unusedTextHeight = Math.Max(0, availableTextHeight - visibleContentHeight);
                                if (verticalAlign == PdfCellVerticalAlign.Middle) verticalOffset = unusedTextHeight / 2;
                                else if (verticalAlign == PdfCellVerticalAlign.Bottom) verticalOffset = unusedTextHeight;
                            }

                            double firstBaseline = y - cellPadTop - verticalOffset - GetAscender(cellFont, rowSize) + style.RowBaselineOffset;

                            pageDirty = true;
                            if (cell.Runs.Any(run => run.Bold || rowUsesBold)) { currentPage!.UsedBold = true; usedBold = true; }
                            if (cell.Runs.Any(run => run.Italic)) { currentPage!.UsedItalic = true; usedItalic = true; }
                            if (cell.Runs.Any(run => (run.Bold || rowUsesBold) && run.Italic)) { currentPage!.UsedBoldItalic = true; usedBoldItalic = true; }
                            MarkRichFonts(cell.Runs);
                            string? linkUri = cell.LinkUri;
                            string? linkDestinationName = cell.LinkDestinationName;
                            string? linkContents = cell.LinkContents;
                            if (tb.Links.TryGetValue((rowIndex, c), out var uri)) {
                                linkUri = uri;
                                linkDestinationName = null;
                                linkContents = cell.Text;
                            }

                            if (sourceStartLine == 0) {
                                AddTableCellNamedDestinationName(cell.NamedDestinationName, y);
                            }

                            if (visibleLineCount > 0) {
                                var visibleLines = SliceTableCellLines(lines, sourceStartLine, visibleLineCount);
                                visibleLines = StripRichLineLinksWhenCellLinked(visibleLines, linkUri, linkDestinationName);
                                var visibleHeights = SliceTableCellLineHeights(lines, sourceStartLine, visibleLineCount, rowLeading);
                                var paragraph = new RichParagraphBlock(StripRunLinksWhenCellLinked(cell.Runs, linkUri, linkDestinationName), MapTableCellAlignment(align), textColor);
                                WriteClippedRichParagraph(sb, paragraph, visibleLines, visibleHeights, currentOpts, firstBaseline, rowSize, rowLeading, currentPage!.Annotations, xi - TableCellClipBleed, cellBottom - TableCellClipBleed, cellWidth + (TableCellClipBleed * 2D), cellHeight + (TableCellClipBleed * 2D), xi + cellPadLeft, innerW);
                            }
                            if (!suppressCellObjects && (cell.Images.Count > 0 || cell.CheckBoxes.Count > 0 || cell.FormFields.Count > 0) && sourceStartLine == 0) {
                                if (CanRenderTableCellCheckBoxInline(cell, lines, sourceStartLine, visibleLineCount)) {
                                    RenderTableCellInlineCheckBox(currentPage!, cell, align, lines.Lines[sourceStartLine], xi + cellPadLeft, innerW, firstBaseline);
                                } else {
                                    double formFieldTop = y - cellPadTop - verticalOffset - (string.IsNullOrEmpty(cell.Text) ? 0D : visibleTextHeight + TableCellCheckBoxGap);
                                    RenderTableCellObjects(currentPage!, cell, align, xi + cellPadLeft, innerW, formFieldTop);
                                }
                            }

                            if (HasCellLinkTarget(linkUri, linkDestinationName)) {
                                double x1 = xi + cellPadLeft;
                                double x2 = xi + cellWidth - cellPadRight;
                                double y1 = cellBottom;
                                double y2 = y;
                                currentPage!.Annotations.Add(new LinkAnnotation { X1 = x1, Y1 = y1, X2 = x2, Y2 = y2, Uri = linkUri, DestinationName = linkDestinationName, Contents = linkContents ?? cell.Text });
                            }
                        }
                        if (style?.BorderColor is not null && style.BorderWidth > 0) {
                            pageDirty = true;
                            bool[] topBorderSkips = GetRowSpanBoundarySkipColumns(tb, rowIndex - 1, cols);
                            bool[] bottomBorderSkips = GetRowSpanBoundarySkipColumns(tb, rowIndex, cols);
                            bool segmentBorderRows = HasSkippedColumns(topBorderSkips, cols) || HasSkippedColumns(bottomBorderSkips, cols);
                            if (segmentBorderRows) {
                                DrawTableHorizontalLine(sb, style.BorderColor.Value, style.BorderWidth, xOrigin, colPixel, colGapPx, yRect + hRect, topBorderSkips);
                                DrawTableHorizontalLine(sb, style.BorderColor.Value, style.BorderWidth, xOrigin, colPixel, colGapPx, yRect, bottomBorderSkips);
                                DrawVLine(sb, style.BorderColor.Value, style.BorderWidth, xOrigin, yRect + hRect, yRect);
                                DrawVLine(sb, style.BorderColor.Value, style.BorderWidth, xOrigin + tableWidth, yRect + hRect, yRect);
                            } else {
                                DrawRowRect(sb, style.BorderColor.Value, style.BorderWidth, xOrigin, yRect, rowWidth, hRect);
                            }

                            double xi2 = xOrigin;
                            double yTop = yRect + hRect;
                            double yBottom = yRect;
                            for (int c = 0; c < cols - 1; c++) {
                                xi2 += colPixel[c];
                                if (IsTableBoundaryInsideSpannedCell(tb, rowIndex, c, cols)) {
                                    xi2 += colGapPx;
                                    continue;
                                }

                                if (currentOpts.Debug?.ShowTableColumnGuides == true)
                                    DrawVLine(sb, new PdfColor(0, 0, 1), Math.Max(0.3, style.BorderWidth), xi2, yTop, yBottom);
                                else
                                    DrawVLine(sb, style.BorderColor.Value, style.BorderWidth, xi2, yTop, yBottom);
                                xi2 += colGapPx;
                            }
                        }
                        if (style != null && renderAsFooter && rowIndex == footerStartRowIndex) {
                            PdfColor? footerSeparatorColor = style.FooterSeparatorColor ?? style.RowSeparatorColor;
                            double footerSeparatorWidth = style.FooterSeparatorWidth > 0 ? style.FooterSeparatorWidth : style.RowSeparatorWidth;
                            if (footerSeparatorColor is not null && footerSeparatorWidth > 0) {
                                pageDirty = true;
                                DrawTableHorizontalLine(sb, footerSeparatorColor.Value, footerSeparatorWidth, xOrigin, colPixel, colGapPx, y, GetRowSpanBoundarySkipColumns(tb, rowIndex - 1, cols));
                            }
                        }
                        PdfColor? separatorColor = renderAsHeader && style?.HeaderSeparatorColor is not null ? style.HeaderSeparatorColor : style?.RowSeparatorColor;
                        double separatorWidth = renderAsHeader && style?.HeaderSeparatorWidth > 0 ? style.HeaderSeparatorWidth : style?.RowSeparatorWidth ?? 0;
                        if (separatorColor is not null && separatorWidth > 0) {
                            pageDirty = true;
                            DrawTableHorizontalLine(sb, separatorColor.Value, separatorWidth, xOrigin, colPixel, colGapPx, rowBottom, GetRowSpanBoundarySkipColumns(tb, rowIndex, cols));
                        }
                        if (style?.CellBorders != null && style.CellBorders.Count > 0) {
                            double borderX = xOrigin;
                            for (int borderColumn = 0; borderColumn < cols; borderColumn++) {
                                if (style.CellBorders.TryGetValue((rowIndex, borderColumn), out PdfCellBorder? cellBorder) &&
                                    TryGetTableCellLayoutAtColumn(cells, borderColumn, out TableCellLayout borderCell) &&
                                    (borderColumn >= rowFillSkips.Length || !rowFillSkips[borderColumn]) &&
                                    HasRenderableCellBorder(cellBorder)) {
                                    int span = wholeRowSegment ? borderCell.ColumnSpan : 1;
                                    double borderHeight = hRect;
                                    double borderBottom = yRect;
                                    if (wholeRowSegment) {
                                        if (borderCell.RowSpan > 1) {
                                            borderHeight = GetTableCellHeight(rowHeights, rowIndex, borderCell.RowSpan, rowGapPx);
                                            borderBottom = y - borderHeight;
                                        }
                                    }

                                    pageDirty = true;
                                    DrawCellBorder(sb, cellBorder, borderX, borderBottom, GetTableCellWidth(colPixel, borderColumn, span, colGapPx), borderHeight);
                                }
                                borderX += colPixel[borderColumn] + colGapPx;
                            }
                        }
                        y -= rowHeight;
                        if (wholeRowSegment) {
                            y -= GetTableRowGapAfter(rowIndex, tb.Rows.Count, rowGapPx);
                        }
                    }

                    void DrawTableRow(int rowIndex, bool renderAsHeader, bool suppressCellObjects = false) =>
                        DrawTableRowSegment(rowIndex, renderAsHeader, 0, rowLineCounts[rowIndex], suppressCellObjects);

                    void DrawSplitTableRow(int rowIndex, bool renderAsHeader) {
                        int startLine = 0;
                        int totalLines = rowLineCounts[rowIndex];
                        while (startLine < totalLines) {
                            double available = y - currentOpts.MarginBottom;
                            double rowPadTop = GetTableRowMaxPaddingTop(tb, style, rowIndex, cols);
                            double rowPadBottom = GetTableRowMaxPaddingBottom(tb, style, rowIndex, cols);
                            double minimumRowSegmentHeight = rowLeadings[rowIndex] + rowPadTop + rowPadBottom;
                            if (available < minimumRowSegmentHeight - 0.001) {
                                NewTablePage(rowIndex);
                                available = y - currentOpts.MarginBottom;
                            }

                            int maxLinesThisPage = Math.Max(1, (int)Math.Floor((available - rowPadTop - rowPadBottom) / rowLeadings[rowIndex]));
                            int take = Math.Min(totalLines - startLine, maxLinesThisPage);
                            DrawTableRowSegment(rowIndex, renderAsHeader && startLine == 0, startLine, take);
                            startLine += take;

                            if (startLine < totalLines) {
                                NewTablePage(rowIndex);
                            }
                        }
                    }

                    for (int rowIndex = 0; rowIndex < tb.Rows.Count; rowIndex++) {
                        if (rowHeights[rowIndex] > maxContentHeight + 0.001) {
                            if (!GetTableRowAllowBreakAcrossPages(style, rowIndex)) {
                                throw new ArgumentException("Table row height exceeds the available page content height and row splitting is disabled.");
                            }

                            DrawSplitTableRow(rowIndex, renderAsHeader: rowIndex < headerRowCount);
                            y -= GetTableRowGapAfter(rowIndex, tb.Rows.Count, rowGapPx);
                            continue;
                        }

                        if (ShouldBreakBefore(rowHeights[rowIndex])) {
                            NewPage();
                            if (hasRepeatableHeader && rowIndex >= headerRowCount && repeatHeaderHeight + rowHeights[rowIndex] <= maxContentHeight + 0.001) {
                                DrawRepeatHeaders();
                            }
                        }

                        DrawTableRow(rowIndex, renderAsHeader: rowIndex < headerRowCount);
                    }

                    y -= style.SpacingAfter;
                } else if (block is HorizontalRuleBlock hr) {
                    PdfHorizontalRuleStyle ruleStyle = ResolveHorizontalRuleStyle(hr, currentOpts);
                    ValidateHorizontalRule(ruleStyle);
                    if (ruleStyle.KeepWithNext && nextBlock != null) {
                        double needed = ruleStyle.SpacingBefore + ruleStyle.Thickness + ruleStyle.SpacingAfter;
                        double nextHeight = MeasureNextBlockFirstVisualHeight(nextBlock, currentOpts.MarginLeft, width, currentOpts.DefaultFontSize);
                        KeepFixedBlockWithNext(needed, nextHeight);
                    }

                    RenderHorizontalRuleBlock(hr, currentOpts.MarginLeft, width);
                } else if (block is TextFieldBlock tf) {
                    RenderTextFieldBlock(tf, currentOpts.MarginLeft, width);
                } else if (block is CheckBoxBlock cbx) {
                    RenderCheckBoxBlock(cbx, currentOpts.MarginLeft, width);
                } else if (block is ChoiceFieldBlock choice) {
                    RenderChoiceFieldBlock(choice, currentOpts.MarginLeft, width);
                } else if (block is RadioButtonGroupBlock radioButtonGroup) {
                    RenderRadioButtonGroupBlock(radioButtonGroup, currentOpts.MarginLeft, width);
                } else if (block is ShapeBlock sbk) {
                    PdfDrawingStyle shapeStyle = ResolveDrawingStyle(sbk, currentOpts);
                    PdfDoc.ValidateDrawingStyle(shapeStyle, "Shape");
                    if (shapeStyle.KeepWithNext && nextBlock != null) {
                        double needed = shapeStyle.SpacingBefore + sbk.Shape.Height + shapeStyle.SpacingAfter;
                        double nextHeight = MeasureNextBlockFirstVisualHeight(nextBlock, currentOpts.MarginLeft, width, currentOpts.DefaultFontSize);
                        KeepFixedBlockWithNext(needed, nextHeight);
                    }

                    RenderShapeBlock(sbk, currentOpts.MarginLeft, width);
                } else if (block is DrawingBlock dbk) {
                    PdfDrawingStyle drawingStyle = ResolveDrawingStyle(dbk, currentOpts);
                    PdfDoc.ValidateDrawingStyle(drawingStyle, "Drawing");
                    if (drawingStyle.KeepWithNext && nextBlock != null) {
                        double needed = drawingStyle.SpacingBefore + dbk.Drawing.Height + drawingStyle.SpacingAfter;
                        double nextHeight = MeasureNextBlockFirstVisualHeight(nextBlock, currentOpts.MarginLeft, width, currentOpts.DefaultFontSize);
                        KeepFixedBlockWithNext(needed, nextHeight);
                    }

                    RenderDrawingBlock(dbk, currentOpts.MarginLeft, width);
                } else if (block is RowBlock rb) {
                    double contentWidth = currentOpts.PageWidth - currentOpts.MarginLeft - currentOpts.MarginRight;
                    int ncols = rb.Columns.Count;
                    PdfRowStyle? rowStyle = rb.StyleSnapshot ?? currentOpts.DefaultRowStyleSnapshot;
                    double rowGap = rb.GapOverride ?? rowStyle?.Gap ?? PdfRowStyle.DefaultGap;
                    double rowSpacingBefore = ResolveTopLevelSpacingBefore(rowStyle?.SpacingBefore ?? 0D);
                    double rowSpacingAfter = rowStyle?.SpacingAfter ?? 0D;
                    double totalGap = rowGap * Math.Max(0, ncols - 1);
                    if (totalGap >= contentWidth) {
                        throw new ArgumentException("Row column gaps must be smaller than the available page content width.");
                    }

                    double columnAreaWidth = contentWidth - totalGap;
                    double[] colXs = new double[ncols];
                    double[] colWs = new double[ncols];
                    double xAcc = currentOpts.MarginLeft;
                    for (int i = 0; i < ncols; i++) { double wCol = Math.Max(0, columnAreaWidth * (rb.Columns[i].WidthPercent / 100.0)); colXs[i] = xAcc; colWs[i] = wCol; xAcc += wCol + rowGap; }

                    void DrawRowColumnSeparators(double topY, double bottomY) {
                        if (ncols <= 1 || rowStyle?.ColumnSeparatorColor == null || rowStyle.ColumnSeparatorWidth <= 0D || topY - bottomY <= 0.001D) {
                            return;
                        }

                        for (int boundary = 0; boundary < ncols - 1; boundary++) {
                            double separatorX = colXs[boundary] + colWs[boundary] + (rowGap / 2D);
                            DrawVLine(sb, rowStyle.ColumnSeparatorColor.Value, rowStyle.ColumnSeparatorWidth, separatorX, topY, bottomY);
                        }

                        pageDirty = true;
                    }

                    var colStates = new System.Collections.Generic.List<(int idx, int line, int subline)>(ncols);
                    var colItems = new System.Collections.Generic.List<System.Collections.Generic.List<ColItem>>(ncols);
                    for (int i = 0; i < ncols; i++) {
                        colStates.Add((0, 0, 0));
                        var items = new System.Collections.Generic.List<ColItem>();
                        foreach (var cb in rb.Columns[i].Blocks) {
                            if (cb is HeadingBlock hb2) {
                                PdfHeadingStyle? headingStyle = ResolveHeadingStyle(hb2, currentOpts);
                                double size = GetHeadingFontSize(hb2, headingStyle);
                                double leading = GetHeadingLeading(headingStyle, size);
                                var headingFont = GetHeadingFont(currentOpts, headingStyle);
                                var lines = WrapSimpleText(hb2.Text, colWs[i], headingFont, size);
                                items.Add(new ColHead {
                                    Block = hb2,
                                    Lines = lines,
                                    Leading = leading,
                                    Size = size,
                                    SpacingBefore = headingStyle?.SpacingBefore ?? 0D,
                                    SpacingAfter = GetHeadingSpacingAfter(headingStyle, leading),
                                    Bold = GetHeadingBold(headingStyle),
                                    ApplySpacingBeforeAtTop = headingStyle?.ApplySpacingBeforeAtTop ?? false,
                                    KeepWithNext = headingStyle?.KeepWithNext ?? true,
                                    Color = hb2.Color ?? headingStyle?.Color
                                });
                            } else if (cb is RichParagraphBlock rpb2) {
                                double size = currentOpts.DefaultFontSize;
                                PdfParagraphStyle? paragraphStyle = EffectiveParagraphStyle(rpb2);
                                double leading = GetParagraphLeading(paragraphStyle, size);
                                var textFrame = GetParagraphTextFrame(paragraphStyle, 0, colWs[i]);
                                var wrap = WrapRichRuns(rpb2.Runs, textFrame.Width, size, ChooseNormal(currentOpts.DefaultFont), leading, textFrame.FirstLineWidth, GetParagraphTabStopWidth(paragraphStyle));
                                items.Add(new ColPar { Block = rpb2, Lines = wrap.Lines, Heights = wrap.LineHeights, Leading = leading, Size = size, XOffset = textFrame.X, TextWidth = textFrame.Width, FirstLineXOffset = textFrame.FirstLineX, FirstLineTextWidth = textFrame.FirstLineWidth });
                            } else if (cb is BulletListBlock bl2) {
                                PdfListStyle? listStyle = ResolveListStyle(bl2, currentOpts);
                                double size = GetListFontSize(listStyle, currentOpts.DefaultFontSize);
                                double leading = GetListLeading(listStyle, size);
                                var baseFont = ChooseNormal(currentOpts.DefaultFont);
                                const string bulletGlyph = "•";
                                double bulletWidth = bl2.RichItems.Count == 0
                                    ? EstimateSimpleTextWidth(bulletGlyph, baseFont, size)
                                    : bl2.RichItems.Max(item => EstimateSimpleTextWidth(item.Marker ?? bulletGlyph, baseFont, size));
                                double spaceAdvance = EstimateSimpleTextWidth(" ", baseFont, size);
                                double markerGap = GetListMarkerGap(listStyle, spaceAdvance);
                                double indent = bulletWidth + markerGap;
                                double listLeftIndent = listStyle?.LeftIndent ?? 0D;
                                double rawTextWidth = colWs[i] - listLeftIndent - indent;
                                double availableWidth = Math.Max(rawTextWidth, EstimateSimpleTextWidth("WW", baseFont, size));
                                double alignmentWidth = Math.Max(0, rawTextWidth);
                                double itemSpacing = GetListItemSpacing(listStyle, leading);
                                var listItems = new System.Collections.Generic.List<ColListItem>(bl2.RichItems.Count);
                                for (int itemIndex = 0; itemIndex < bl2.RichItems.Count; itemIndex++) {
                                    var item = bl2.RichItems[itemIndex];
                                    string marker = item.Marker ?? bulletGlyph;
                                    var layout = CreateListItemTextLayout(item, availableWidth, baseFont, size, leading);
                                    double firstLineWidth = layout.Lines.Count > 0 ? MeasureRichLineWidth(layout.Lines[0]) : 0;
                                    double firstLineDx = 0;
                                    if (bl2.Align == PdfAlign.Center) firstLineDx = Math.Max(0, (alignmentWidth - firstLineWidth) / 2);
                                    else if (bl2.Align == PdfAlign.Right) firstLineDx = Math.Max(0, alignmentWidth - firstLineWidth);
                                    double spacingBefore = itemIndex == 0 ? listStyle?.SpacingBefore ?? 0D : 0D;
                                    double spacingAfter = itemIndex == bl2.RichItems.Count - 1 ? listStyle?.GetSpacingAfter(itemSpacing) ?? itemSpacing : itemSpacing;
                                    listItems.Add(new ColListItem { Runs = item.Runs, Lines = layout.Lines, Heights = layout.LineHeights, Marker = marker, MarkerXOffset = listLeftIndent + firstLineDx, MarkerWidth = bulletWidth, MarkerAlign = PdfAlign.Left, TextXOffset = listLeftIndent + indent, TextWidth = alignmentWidth, TextAlign = bl2.Align, Color = bl2.Color ?? listStyle?.Color, Leading = leading, Size = size, SpacingBefore = spacingBefore, SpacingAfter = spacingAfter, BookmarkName = item.BookmarkName });
                                }

                                if ((listStyle?.KeepTogether == true || listStyle?.KeepWithNext == true) && listItems.Count > 0) {
                                    double listGroupHeight = 0D;
                                    foreach (var listItem in listItems) {
                                        listGroupHeight += listItem.SpacingBefore + MeasureRichLinesHeight(listItem.Heights, listItem.Lines.Count, listItem.Leading) + listItem.SpacingAfter;
                                    }

                                    if (listStyle?.KeepTogether == true) {
                                        listItems[0].IsFirstInKeepGroup = true;
                                        foreach (var listItem in listItems) {
                                            listItem.KeepTogether = true;
                                            listItem.KeepGroupHeight = listGroupHeight;
                                        }
                                    }

                                    if (listStyle?.KeepWithNext == true) {
                                        listItems[0].IsFirstInKeepWithNextGroup = true;
                                        foreach (var listItem in listItems) {
                                            listItem.KeepWithNext = true;
                                            listItem.KeepWithNextGroupItemCount = listItems.Count;
                                            listItem.KeepWithNextGroupHeight = listGroupHeight;
                                        }
                                    }
                                }

                                items.AddRange(listItems);
                            } else if (cb is NumberedListBlock nl2) {
                                PdfListStyle? listStyle = ResolveListStyle(nl2, currentOpts);
                                double size = GetListFontSize(listStyle, currentOpts.DefaultFontSize);
                                double leading = GetListLeading(listStyle, size);
                                var baseFont = ChooseNormal(currentOpts.DefaultFont);
                                int lastNumber = nl2.StartNumber + Math.Max(0, nl2.RichItems.Count - 1);
                                string widestMarker = lastNumber.ToString(CultureInfo.InvariantCulture) + ".";
                                double markerWidth = nl2.RichItems.Count == 0
                                    ? EstimateSimpleTextWidth(widestMarker, baseFont, size)
                                    : nl2.RichItems
                                        .Select((item, itemIndex) => item.Marker ?? ((nl2.StartNumber + itemIndex).ToString(CultureInfo.InvariantCulture) + "."))
                                        .Max(marker => EstimateSimpleTextWidth(marker, baseFont, size));
                                double spaceAdvance = EstimateSimpleTextWidth(" ", baseFont, size);
                                double markerGap = GetListMarkerGap(listStyle, spaceAdvance);
                                double indent = markerWidth + markerGap;
                                double listLeftIndent = listStyle?.LeftIndent ?? 0D;
                                double rawTextWidth = colWs[i] - listLeftIndent - indent;
                                double availableWidth = Math.Max(rawTextWidth, EstimateSimpleTextWidth("WW", baseFont, size));
                                double alignmentWidth = Math.Max(0, rawTextWidth);
                                double itemSpacing = GetListItemSpacing(listStyle, leading);
                                var listItems = new System.Collections.Generic.List<ColListItem>(nl2.RichItems.Count);
                                for (int itemIndex = 0; itemIndex < nl2.RichItems.Count; itemIndex++) {
                                    var item = nl2.RichItems[itemIndex];
                                    string marker = item.Marker ?? ((nl2.StartNumber + itemIndex).ToString(CultureInfo.InvariantCulture) + ".");
                                    var layout = CreateListItemTextLayout(item, availableWidth, baseFont, size, leading);
                                    double firstLineWidth = layout.Lines.Count > 0 ? MeasureRichLineWidth(layout.Lines[0]) : 0;
                                    double firstLineDx = 0;
                                    if (nl2.Align == PdfAlign.Center) firstLineDx = Math.Max(0, (alignmentWidth - firstLineWidth) / 2);
                                    else if (nl2.Align == PdfAlign.Right) firstLineDx = Math.Max(0, alignmentWidth - firstLineWidth);
                                    double spacingBefore = itemIndex == 0 ? listStyle?.SpacingBefore ?? 0D : 0D;
                                    double spacingAfter = itemIndex == nl2.RichItems.Count - 1 ? listStyle?.GetSpacingAfter(itemSpacing) ?? itemSpacing : itemSpacing;
                                    listItems.Add(new ColListItem { Runs = item.Runs, Lines = layout.Lines, Heights = layout.LineHeights, Marker = marker, MarkerXOffset = listLeftIndent + firstLineDx, MarkerWidth = markerWidth, MarkerAlign = PdfAlign.Right, TextXOffset = listLeftIndent + indent, TextWidth = alignmentWidth, TextAlign = nl2.Align, Color = nl2.Color ?? listStyle?.Color, Leading = leading, Size = size, SpacingBefore = spacingBefore, SpacingAfter = spacingAfter, BookmarkName = item.BookmarkName });
                                }

                                if ((listStyle?.KeepTogether == true || listStyle?.KeepWithNext == true) && listItems.Count > 0) {
                                    double listGroupHeight = 0D;
                                    foreach (var listItem in listItems) {
                                        listGroupHeight += listItem.SpacingBefore + MeasureRichLinesHeight(listItem.Heights, listItem.Lines.Count, listItem.Leading) + listItem.SpacingAfter;
                                    }

                                    if (listStyle?.KeepTogether == true) {
                                        listItems[0].IsFirstInKeepGroup = true;
                                        foreach (var listItem in listItems) {
                                            listItem.KeepTogether = true;
                                            listItem.KeepGroupHeight = listGroupHeight;
                                        }
                                    }

                                    if (listStyle?.KeepWithNext == true) {
                                        listItems[0].IsFirstInKeepWithNextGroup = true;
                                        foreach (var listItem in listItems) {
                                            listItem.KeepWithNext = true;
                                            listItem.KeepWithNextGroupItemCount = listItems.Count;
                                            listItem.KeepWithNextGroupHeight = listGroupHeight;
                                        }
                                    }
                                }

                                items.AddRange(listItems);
                            } else if (cb is PanelParagraphBlock ppb2) {
                                double size = currentOpts.DefaultFontSize;
                                double leading = size * 1.4;
                                var panelFont = ChooseNormal(currentOpts.DefaultFont);
                                double firstBaselineOffset = GetAscender(panelFont, size);
                                PanelStyle panelStyle = ResolvePanelStyle(ppb2, currentOpts);
                                double innerWidth = panelStyle.MaxWidth.HasValue ? Math.Min(colWs[i], panelStyle.MaxWidth.Value) : colWs[i];
                                ValidatePanelStyle(panelStyle, innerWidth);
                                double textWidthAvail = innerWidth - 2 * panelStyle.PaddingX;
                                var wrap = WrapRichRuns(ppb2.Runs, textWidthAvail, size, panelFont, leading);
                                double xOffset = 0;
                                if (panelStyle.Align == PdfAlign.Center) xOffset = Math.Max(0, (colWs[i] - innerWidth) / 2);
                                else if (panelStyle.Align == PdfAlign.Right) xOffset = Math.Max(0, colWs[i] - innerWidth);
                                items.Add(new ColPanel { Block = ppb2, Style = panelStyle, Lines = wrap.Lines, Heights = wrap.LineHeights, Leading = leading, Size = size, FirstBaselineOffset = firstBaselineOffset, XOffset = xOffset, PanelWidth = innerWidth, TextWidth = textWidthAvail });
                            } else if (cb is TableBlock tb2) {
                                PdfTableStyle style = tb2.Style ?? currentOpts.DefaultTableStyleSnapshot ?? TableStyles.Light();
                                int cols = GetTableColumnCount(tb2);
                                if (cols == 0) {
                                    continue;
                                }

                                double padLeft = GetTableCellPaddingLeft(style);
                                double padRight = GetTableCellPaddingRight(style);
                                double padTop = GetTableCellPaddingTop(style);
                                double padBottom = GetTableCellPaddingBottom(style);
                                double cellSpacing = GetTableCellSpacing(style);
                                double columnGap = cellSpacing;
                                double tableRowGap = cellSpacing;
                                double size = currentOpts.DefaultFontSize;
                                ValidateTableRoleRowCounts(style, tb2.Rows.Count);
                                int headerRowCount = style.HeaderRowCount;
                                int repeatHeaderRowCount = GetTableRepeatHeaderRowCount(style);
                                int footerRowCount = style.FooterRowCount;
                                int footerStartRowIndex = tb2.Rows.Count - footerRowCount;
                                ValidateTableCellStyleCoordinates(style, tb2.Rows.Count, cols);
                                ValidateTableColumnStyleBounds(style, cols);
                                ValidateTableRowStyleBounds(style, tb2.Rows.Count);
                                ValidateTableRowSpansWithinRoleBoundaries(tb2, cols, headerRowCount, footerStartRowIndex);
                                double[]? autoFitWeights = style.AutoFitColumns
                                    ? MeasureAutoFitColumnWeights(tb2, currentOpts, style, size, headerRowCount, footerStartRowIndex)
                                    : null;
                                double[]? autoFitMinimumWidths = style.AutoFitColumns
                                    ? MeasureAutoFitColumnMinimumWidths(tb2, currentOpts, style, size, headerRowCount, footerStartRowIndex)
                                    : null;
                                double[] colPixel = new double[cols];
                                double[] colWeights = new double[cols];
                                bool[] fixedColumns = new bool[cols];
                                double?[] minWidths = new double?[cols];
                                double?[] maxWidths = new double?[cols];
                                double fixedWidthTotal = 0;
                                double totalWeight = 0;
                                for (int c = 0; c < cols; c++) {
                                    double? minWidth = GetOptionalColumnWidth(style.ColumnMinWidthPoints, c, "Table minimum column widths must be positive finite values.");
                                    if (!minWidth.HasValue && autoFitMinimumWidths != null && c < autoFitMinimumWidths.Length) {
                                        minWidth = autoFitMinimumWidths[c];
                                    }

                                    double? maxWidth = GetOptionalColumnWidth(style.ColumnMaxWidthPoints, c, "Table maximum column widths must be positive finite values.");
                                    if (minWidth.HasValue && maxWidth.HasValue && minWidth.Value > maxWidth.Value + 0.001) {
                                        throw new ArgumentException("Table minimum column widths cannot be greater than maximum column widths.");
                                    }

                                    minWidths[c] = minWidth;
                                    maxWidths[c] = maxWidth;

                                    if (style.ColumnWidthPoints != null && c < style.ColumnWidthPoints.Count && style.ColumnWidthPoints[c].HasValue) {
                                        double fixedWidth = style.ColumnWidthPoints[c]!.Value;
                                        if (minWidth.HasValue && fixedWidth < minWidth.Value - 0.001) {
                                            throw new ArgumentException("Table fixed column widths cannot be smaller than configured minimum widths.");
                                        }
                                        if (maxWidth.HasValue && fixedWidth > maxWidth.Value + 0.001) {
                                            throw new ArgumentException("Table fixed column widths cannot be greater than configured maximum widths.");
                                        }

                                        colPixel[c] = fixedWidth;
                                        fixedColumns[c] = true;
                                        fixedWidthTotal += fixedWidth;
                                        continue;
                                    }

                                    double weight = 1;
                                    if (style.ColumnWidthWeights != null && c < style.ColumnWidthWeights.Count) {
                                        weight = style.ColumnWidthWeights[c];
                                    } else if (autoFitWeights != null && c < autoFitWeights.Length) {
                                        weight = autoFitWeights[c];
                                    }

                                    colWeights[c] = weight;
                                    totalWeight += weight;
                                }

                                double tableAvailableWidth = ResolveTableAvailableWidth(style, colWs[i]);
                                double tableInnerAvailableWidth = tableAvailableWidth - (cols - 1) * columnGap;
                                if (tableInnerAvailableWidth <= 0.001 || double.IsNaN(tableInnerAvailableWidth) || double.IsInfinity(tableInnerAvailableWidth)) {
                                    throw new ArgumentException("Table cell spacing must leave a positive table width.");
                                }

                                fixedWidthTotal = FitFixedTableColumnsToAvailableWidth(colPixel, fixedColumns, minWidths, fixedWidthTotal, tableInnerAvailableWidth);

                                double remainingWidth = Math.Max(0, tableInnerAvailableWidth - fixedWidthTotal);
                                if (totalWeight <= 0) {
                                    remainingWidth = 0;
                                }

                                DistributeFlexibleColumns(colPixel, colWeights, fixedColumns, minWidths, maxWidths, remainingWidth);
                                double tableWidth = Math.Min(tableAvailableWidth, colPixel.Sum() + (cols - 1) * columnGap);
                                ValidateTableCellTextWidths(tb2, style, cols, colPixel, columnGap);

                                var rowLines = new TableCellTextLayout[tb2.Rows.Count][];
                                var rowLineCounts = new int[tb2.Rows.Count];
                                var rowHeights = new double[tb2.Rows.Count];
                                var rowLeadings = new double[tb2.Rows.Count];
                                var rowSizes = new double[tb2.Rows.Count];
                                var rowBold = new bool[tb2.Rows.Count];
                                for (int ri = 0; ri < tb2.Rows.Count; ri++) {
                                    double rowSize = GetTableRowFontSize(style, ri, headerRowCount, footerStartRowIndex, currentOpts.DefaultFontSize);
                                    double rowLeading = GetTableLeading(style, rowSize);
                                    bool rowUsesBold = GetTableRowBold(style, ri, headerRowCount, footerStartRowIndex);
                                    rowSizes[ri] = rowSize;
                                    rowLeadings[ri] = rowLeading;
                                    rowBold[ri] = rowUsesBold;
                                    rowLines[ri] = new TableCellTextLayout[cols];
                                    int maxLines = 1;
                                    double maxRequiredHeight = rowLeading + GetTableRowMaxPaddingTop(tb2, style, ri, cols) + GetTableRowMaxPaddingBottom(tb2, style, ri, cols);
                                    for (int ci = 0; ci < cols; ci++) {
                                        rowLines[ri][ci] = new TableCellTextLayout(new System.Collections.Generic.List<System.Collections.Generic.List<RichSeg>> { new() }, new System.Collections.Generic.List<double> { rowLeading });
                                    }

                                    var cells = GetTableCellLayouts(tb2, ri, cols);
                                    for (int cellIndex = 0; cellIndex < cells.Count; cellIndex++) {
                                        TableCellLayout cell = cells[cellIndex];
                                        var cellFont = GetTableRowFont(currentOpts, rowUsesBold);
                                        double cellWidth = GetTableCellWidth(colPixel, cell.Column, cell.ColumnSpan, columnGap);
                                        double innerWidth = Math.Max(1, cellWidth - GetTableCellPaddingLeft(style, ri, cell.Column) - GetTableCellPaddingRight(style, ri, cell.Column));
                                        TableCellTextLayout lines = CreateTableCellTextLayout(cell, innerWidth, cellFont, rowSize, rowLeading);
                                        rowLines[ri][cell.Column] = lines;
                                        if (cell.RowSpan <= 1) {
                                            maxLines = Math.Max(maxLines, lines.LineCount);
                                        maxRequiredHeight = Math.Max(maxRequiredHeight, MeasureTableCellContentHeight(cell, lines, 0, lines.LineCount, rowLeading) + GetTableCellPaddingTop(style, ri, cell.Column) + GetTableCellPaddingBottom(style, ri, cell.Column));
                                        }
                                    }

                                    rowLineCounts[ri] = maxLines;
                                    rowHeights[ri] = Math.Max(maxRequiredHeight, GetTableRowMinHeight(style, ri));
                                }
                                ApplyTableRowSpanHeights(tb2, style, cols, rowLines, rowHeights, rowLeadings, tableRowGap);

                                System.Collections.Generic.List<string>? captionLines = null;
                                double captionLeading = 0;
                                double captionHeight = 0;
                                if (!string.IsNullOrWhiteSpace(style.Caption)) {
                                    double captionSize = style.CaptionFontSize ?? size;
                                    captionLeading = captionSize * 1.25;
                                    var captionFont = ChooseNormal(currentOpts.DefaultFont);
                                    captionLines = WrapSimpleText(style.Caption!, tableWidth, captionFont, captionSize);
                                    captionHeight = captionLines.Count * captionLeading + style.CaptionSpacingAfter;
                                    double firstRowHeight = rowHeights.Length > 0 ? rowHeights[0] : 0;
                                    double maxContentHeightForCaption = currentOpts.PageHeight - currentOpts.MarginTop - currentOpts.MarginBottom;
                                    if (captionHeight + firstRowHeight > maxContentHeightForCaption + 0.001) {
                                        throw new ArgumentException("Table caption and first row exceed the available page content height.");
                                    }
                                }

                                items.Add(new ColTable { Block = tb2, Style = style, Columns = cols, ColumnWidths = colPixel, RowLines = rowLines, RowLineCounts = rowLineCounts, RowHeights = rowHeights, RowLeadings = rowLeadings, RowSizes = rowSizes, RowBold = rowBold, Width = tableWidth, Size = size, HeaderRowCount = headerRowCount, RepeatHeaderRowCount = repeatHeaderRowCount, FooterStartRowIndex = footerStartRowIndex, CaptionLines = captionLines, CaptionLeading = captionLeading, CaptionHeight = captionHeight });
                            } else if (cb is HorizontalRuleBlock hr2) {
                                items.Add(new ColRule { Block = hr2 });
                            } else if (cb is ImageBlock ib2) {
                                items.Add(new ColImg { Block = ib2 });
                            } else if (cb is ShapeBlock sb2) {
                                items.Add(new ColShape { Block = sb2 });
                            } else if (cb is DrawingBlock db2) {
                                items.Add(new ColDrawing { Block = db2 });
                            } else if (cb is TextFieldBlock || cb is CheckBoxBlock || cb is ChoiceFieldBlock || cb is RadioButtonGroupBlock) {
                                items.Add(new ColForm { Block = cb });
                            } else if (cb is BookmarkBlock bookmark2) {
                                items.Add(new ColBookmark { Block = bookmark2 });
                            } else if (cb is SpacerBlock spacer2) {
                                items.Add(new ColSpacer { Block = spacer2 });
                            }
                        }
                        colItems.Add(items);
                    }

                    double MeasureRowKeepTogetherHeight(System.Collections.Generic.List<ColItem> items) {
                        double total = 0D;
                        foreach (var item in items) {
                            if (item is ColPar paragraph) {
                                PdfParagraphStyle? paragraphStyle = EffectiveParagraphStyle(paragraph.Block);
                                total += ResolveColumnSpacingBefore(GetParagraphSpacingBefore(paragraphStyle), total) + paragraph.Heights.Sum() + GetParagraphSpacingAfter(paragraphStyle, paragraph.Leading);
                            } else if (item is ColHead heading) {
                                total += ResolveColumnSpacingBefore(heading.SpacingBefore, total) + heading.Lines.Count * heading.Leading + heading.SpacingAfter;
                            } else if (item is ColListItem listItem) {
                                total += ResolveColumnSpacingBefore(listItem.SpacingBefore, total) + MeasureRichLinesHeight(listItem.Heights, listItem.Lines.Count, listItem.Leading) + listItem.SpacingAfter;
                            } else if (item is ColPanel panel) {
                                total += ResolveColumnSpacingBefore(panel.Style.SpacingBefore, total) + panel.Style.PaddingY + panel.Heights.Sum() + panel.Style.PaddingY + panel.Style.SpacingAfter;
                            } else if (item is ColTable table) {
                                total += ResolveColumnSpacingBefore(table.Style.SpacingBefore, total) + table.CaptionHeight + GetTableRowsHeight(table.RowHeights, 0, table.RowHeights.Length, GetTableCellSpacing(table.Style)) + table.Style.SpacingAfter;
                            } else if (item is ColRule rule) {
                                PdfHorizontalRuleStyle ruleStyle = ResolveHorizontalRuleStyle(rule.Block, currentOpts);
                                ValidateHorizontalRule(ruleStyle);
                                total += ResolveColumnSpacingBefore(ruleStyle.SpacingBefore, total) + ruleStyle.Thickness + ruleStyle.SpacingAfter;
                            } else if (item is ColImg image) {
                                PdfImageStyle imageStyle = ResolveImageStyle(image.Block, currentOpts);
                                total += ResolveColumnSpacingBefore(imageStyle.SpacingBefore, total) + image.Block.Height + imageStyle.SpacingAfter;
                            } else if (item is ColShape shape) {
                                PdfDrawingStyle shapeStyle = ResolveDrawingStyle(shape.Block, currentOpts);
                                total += ResolveColumnSpacingBefore(shapeStyle.SpacingBefore, total) + shape.Block.Shape.Height + shapeStyle.SpacingAfter;
                            } else if (item is ColDrawing drawing) {
                                PdfDrawingStyle drawingStyle = ResolveDrawingStyle(drawing.Block, currentOpts);
                                total += ResolveColumnSpacingBefore(drawingStyle.SpacingBefore, total) + drawing.Block.Drawing.Height + drawingStyle.SpacingAfter;
                            } else if (item is ColForm form) {
                                total += ResolveColumnSpacingBefore(GetFormFieldSpacingBefore(form.Block), total) + GetFormFieldHeight(form.Block) + GetFormFieldSpacingAfter(form.Block);
                            } else if (item is ColSpacer spacerItem) {
                                total += spacerItem.Block.Height;
                            }
                        }

                        return total;
                    }

                    double MeasureColItemFirstVisualHeight(ColItem item) {
                        if (item is ColPar paragraph) {
                            PdfParagraphStyle? paragraphStyle = EffectiveParagraphStyle(paragraph.Block);
                            return GetParagraphSpacingBefore(paragraphStyle) + (paragraph.Heights.Count == 0 ? 0D : paragraph.Heights[0]);
                        }

                        if (item is ColHead heading) {
                            return heading.SpacingBefore + (heading.Lines.Count == 0 ? 0D : heading.Leading);
                        }

                        if (item is ColListItem listItem) {
                            return listItem.SpacingBefore + (listItem.Lines.Count == 0 ? 0D : GetRichLineHeight(listItem.Heights, 0, listItem.Leading));
                        }

                        if (item is ColPanel panel) {
                            return panel.Style.SpacingBefore + panel.Style.PaddingY + (panel.Heights.Count == 0 ? 0D : panel.Heights[0]) + panel.Style.PaddingY;
                        }

                        if (item is ColTable table) {
                            double firstRowHeight = table.RowHeights.Length == 0 ? 0D : table.RowHeights[0];
                            return table.Style.SpacingBefore + table.CaptionHeight + firstRowHeight;
                        }

                        if (item is ColRule rule) {
                            PdfHorizontalRuleStyle ruleStyle = ResolveHorizontalRuleStyle(rule.Block, currentOpts);
                            return ruleStyle.SpacingBefore + ruleStyle.Thickness + ruleStyle.SpacingAfter;
                        }

                        if (item is ColImg image) {
                            PdfImageStyle imageStyle = ResolveImageStyle(image.Block, currentOpts);
                            return imageStyle.SpacingBefore + image.Block.Height + imageStyle.SpacingAfter;
                        }

                        if (item is ColShape shape) {
                            PdfDrawingStyle shapeStyle = ResolveDrawingStyle(shape.Block, currentOpts);
                            return shapeStyle.SpacingBefore + shape.Block.Shape.Height + shapeStyle.SpacingAfter;
                        }

                        if (item is ColDrawing drawing) {
                            PdfDrawingStyle drawingStyle = ResolveDrawingStyle(drawing.Block, currentOpts);
                            return drawingStyle.SpacingBefore + drawing.Block.Drawing.Height + drawingStyle.SpacingAfter;
                        }

                        if (item is ColForm form) {
                            return GetFormFieldSpacingBefore(form.Block) + GetFormFieldHeight(form.Block) + GetFormFieldSpacingAfter(form.Block);
                        }

                        if (item is ColSpacer spacerItem) {
                            return spacerItem.Block.Height;
                        }

                        return 0D;
                    }

                    double? rowContentHeightCache = null;
                    double GetRowContentHeight() {
                        if (rowContentHeightCache.HasValue) {
                            return rowContentHeightCache.Value;
                        }

                        double measuredHeight = 0D;
                        foreach (var items in colItems) {
                            measuredHeight = Math.Max(measuredHeight, MeasureRowKeepTogetherHeight(items));
                        }

                        rowContentHeightCache = measuredHeight;
                        return measuredHeight;
                    }

                    if (rowStyle?.KeepTogether == true) {
                        double rowContentHeight = GetRowContentHeight();
                        double rowKeepHeight = rowSpacingBefore + rowContentHeight + rowSpacingAfter;
                        double availableHeight = currentOpts.PageHeight - currentOpts.MarginTop - currentOpts.MarginBottom;
                        if (rowKeepHeight > availableHeight + 0.001) {
                            throw new ArgumentException("Row height exceeds the available page content height.");
                        }

                        if (y < yStart - 0.001 && y - rowKeepHeight < currentOpts.MarginBottom) {
                            NewPage();
                            rowSpacingBefore = 0D;
                        }
                    }

                    if (rowStyle?.KeepWithNext == true && nextBlock != null) {
                        double rowContentHeight = GetRowContentHeight();
                        double rowHeight = rowSpacingBefore + rowContentHeight + rowSpacingAfter;
                        double nextHeight = MeasureNextBlockFirstVisualHeight(nextBlock, currentOpts.MarginLeft, width, currentOpts.DefaultFontSize);
                        double keepHeight = rowHeight + nextHeight;
                        double availableHeight = currentOpts.PageHeight - currentOpts.MarginTop - currentOpts.MarginBottom;
                        if (nextHeight > 0.001 && rowHeight <= availableHeight + 0.001 && keepHeight <= availableHeight + 0.001 && y < yStart - 0.001 && y - keepHeight < currentOpts.MarginBottom) {
                            NewPage();
                            rowSpacingBefore = 0D;
                        }
                    }

                    if (rowSpacingBefore > 0) {
                        if (y - rowSpacingBefore < currentOpts.MarginBottom) {
                            NewPage();
                            rowSpacingBefore = 0D;
                        }

                        if (rowSpacingBefore > 0) y -= rowSpacingBefore;
                    }

                    bool AnyRemaining() {
                        for (int i = 0; i < ncols; i++) if (colStates[i].idx < colItems[i].Count) return true; return false;
                    }

                    int rowColumnFlowGuard = 0;
                    while (AnyRemaining()) {
                        rowColumnFlowGuard++;
                        if (rowColumnFlowGuard > 10000) {
                            throw new InvalidOperationException("Row column layout did not make forward progress.");
                        }

                        double avail = y - currentOpts.MarginBottom;
                        if (avail <= 0.5) { NewPage(); avail = y - currentOpts.MarginBottom; }

                        double maxConsumed = 0;
                        bool anyColumnAdvanced = false;
                        for (int ci = 0; ci < ncols; ci++) {
                            var items = colItems[ci];
                            var (idx, line, subline) = colStates[ci];
                            var startState = (idx, line, subline);
                            double xCol = colXs[ci];
                            double wCol = colWs[ci];
                            double yCol = y;
                            double consumed = 0;
                            double remain = avail;
                            while (idx < items.Count && remain > 0.1) {
                                var it = items[idx];
                                if (it is ColPar par) {
                                    var pblock = par.Block;
                                    var lines = par.Lines;
                                    var heights = par.Heights;
                                    double leading = par.Leading;
                                    double size = par.Size;
                                    PdfParagraphStyle? paragraphStyle = EffectiveParagraphStyle(pblock);
                                    double spacingBefore = line == 0 && consumed > 0.001 ? GetParagraphSpacingBefore(paragraphStyle) : 0;
                                    double spacingAfter = GetParagraphSpacingAfter(paragraphStyle, leading);
                                    if (paragraphStyle?.KeepWithNext == true && line == 0 && idx + 1 < items.Count) {
                                        double nextHeight = MeasureColItemFirstVisualHeight(items[idx + 1]);
                                        double keepHeight = spacingBefore + heights.Sum() + spacingAfter + nextHeight;
                                        double availableHeight = currentOpts.PageHeight - currentOpts.MarginTop - currentOpts.MarginBottom;
                                        if (nextHeight > 0.001 && keepHeight <= availableHeight + 0.001 && keepHeight > remain + 0.001) {
                                            if (consumed > 0) break;
                                            remain = 0;
                                            break;
                                        }
                                    }

                                    if (paragraphStyle?.KeepTogether == true && line == 0) {
                                        double paragraphHeight = spacingBefore + heights.Sum() + spacingAfter;
                                        double availableHeight = currentOpts.PageHeight - currentOpts.MarginTop - currentOpts.MarginBottom;
                                        if (paragraphHeight > availableHeight + 0.001) {
                                            throw new ArgumentException("Paragraph height exceeds the available page content height.");
                                        }

                                        if (paragraphHeight > remain + 0.001) {
                                            if (consumed > 0) break;
                                            remain = 0;
                                            break;
                                        }
                                    }

                                    double availableForLines = remain - spacingBefore;
                                    if (availableForLines < 0) {
                                        if (consumed > 0) break;
                                        remain = 0;
                                        break;
                                    }

                                    int start = line;
                                    int take = 0; double hsum = 0;
                                    for (int li2 = start; li2 < lines.Count; li2++) {
                                        double hAdd = heights[li2];
                                        if (hsum + hAdd + (li2 == lines.Count - 1 ? spacingAfter : 0) > availableForLines) break;
                                        hsum += hAdd; take++;
                                    }

                                    if (TryApplyWidowControl(paragraphStyle, lines.Count, start, ref take, ref hsum, heights, consumed > 0 || y < yStart - 0.001)) {
                                        break;
                                    }

                                    if (take == 0) break;
                                    if (spacingBefore > 0) {
                                        yCol -= spacingBefore;
                                        remain -= spacingBefore;
                                        consumed += spacingBefore;
                                    }

                                    var sliceLines = new System.Collections.Generic.List<System.Collections.Generic.List<RichSeg>>();
                                    var sliceHeights = new System.Collections.Generic.List<double>();
                                    for (int k = 0; k < take; k++) { sliceLines.Add(lines[start + k]); sliceHeights.Add(heights[start + k]); }
                                    pageDirty = true;
                                    var paragraphFont = ChooseNormal(currentOpts.DefaultFont);
                                    WriteRichParagraph(sb, pblock, sliceLines, sliceHeights, currentOpts, FirstTextBaselineFromTop(paragraphFont, size, yCol), size, leading, currentPage!.Annotations, xCol + par.XOffset, par.TextWidth, start == 0 ? xCol + par.FirstLineXOffset : null, start == 0 ? par.FirstLineTextWidth : null);
                                    MarkRichFonts(pblock.Runs);
                                    yCol -= hsum; remain -= hsum; consumed += hsum; line += take;
                                    if (line >= lines.Count) { double space = spacingAfter; if (space <= remain) { yCol -= space; remain -= space; consumed += space; } idx++; line = 0; }
                                } else if (it is ColHead ch) {
                                    var hb2 = ch.Block;
                                    var lines = ch.Lines;
                                    double leading = ch.Leading;
                                    double size = ch.Size;
                                    double spacingBefore = (consumed > 0.001 || ch.ApplySpacingBeforeAtTop) ? ch.SpacingBefore : 0D;
                                    double needed = spacingBefore + lines.Count * leading + ch.SpacingAfter;
                                    if (ch.KeepWithNext && idx + 1 < items.Count) {
                                        double nextHeight = MeasureColItemFirstVisualHeight(items[idx + 1]);
                                        double keepHeight = needed + nextHeight;
                                        double availableHeight = currentOpts.PageHeight - currentOpts.MarginTop - currentOpts.MarginBottom;
                                        if (nextHeight > 0.001 && keepHeight <= availableHeight + 0.001 && keepHeight > remain + 0.001) {
                                            if (consumed > 0) break;
                                            remain = 0;
                                            break;
                                        }
                                    }

                                    if (needed > remain && consumed > 0) break;
                                    if (needed > remain && consumed == 0) { remain = 0; break; }
                                    if (spacingBefore > 0) {
                                        yCol -= spacingBefore;
                                        remain -= spacingBefore;
                                        consumed += spacingBefore;
                                    }

                                    if (currentOpts.CreateOutlineFromHeadings) {
                                        currentPage!.Bookmarks.Add(new PageBookmark { Level = hb2.Level, Title = hb2.Text, Y = yCol });
                                    }
                                    var headingFont = ch.Bold ? ChooseBold(ChooseNormal(currentOpts.DefaultFont)) : ChooseNormal(currentOpts.DefaultFont);
                                    double firstBaseline = FirstTextBaselineFromTop(headingFont, size, yCol);
                                    AddHeadingLinkAnnotations(hb2, lines, headingFont, size, leading, xCol, wCol, firstBaseline);
                                    WriteLinesInternal(ch.Bold ? "F2" : "F1", size, leading, xCol, wCol, firstBaseline, lines, hb2.Align, ch.Color, applyBaselineTweak: false);
                                    if (ch.Bold) {
                                        currentPage!.UsedBold = true;
                                        usedBold = true;
                                    }
                                    double consumedHeight = lines.Count * leading + ch.SpacingAfter;
                                    yCol -= consumedHeight; remain -= consumedHeight; consumed += consumedHeight; idx++;
                                } else if (it is ColListItem listItem) {
                                    var lines = listItem.Lines;
                                    double leading = listItem.Leading;
                                    double spacingBefore = line == 0 ? ResolveColumnSpacingBefore(listItem.SpacingBefore, consumed) : 0D;
                                    if (line == 0 && listItem.KeepTogether && listItem.IsFirstInKeepGroup) {
                                        double keepGroupHeight = listItem.KeepGroupHeight - listItem.SpacingBefore + spacingBefore;
                                        double availableHeight = currentOpts.PageHeight - currentOpts.MarginTop - currentOpts.MarginBottom;
                                        if (keepGroupHeight > availableHeight + 0.001) {
                                            throw new ArgumentException("List height exceeds the available page content height.");
                                        }

                                        if (keepGroupHeight > remain + 0.001) {
                                            if (consumed > 0) break;
                                            remain = 0;
                                            break;
                                        }
                                    }

                                    if (line == 0 && listItem.KeepWithNext && listItem.IsFirstInKeepWithNextGroup) {
                                        int nextItemIndex = idx + listItem.KeepWithNextGroupItemCount;
                                        if (nextItemIndex < items.Count) {
                                            double nextHeight = MeasureColItemFirstVisualHeight(items[nextItemIndex]);
                                            double keepHeight = listItem.KeepWithNextGroupHeight - listItem.SpacingBefore + spacingBefore + nextHeight;
                                            double availableHeight = currentOpts.PageHeight - currentOpts.MarginTop - currentOpts.MarginBottom;
                                            if (nextHeight > 0.001 && keepHeight <= availableHeight + 0.001 && keepHeight > remain + 0.001) {
                                                if (consumed > 0) break;
                                                remain = 0;
                                                break;
                                            }
                                        }
                                    }

                                    if (line == 0 && spacingBefore > 0) {
                                        if (spacingBefore > remain && consumed > 0) break;
                                        if (spacingBefore > remain && consumed == 0) { remain = 0; break; }
                                        yCol -= spacingBefore;
                                        remain -= spacingBefore;
                                        consumed += spacingBefore;
                                    }

                                    double availableForLines = remain;
                                    int start = line;
                                    int take = 0;
                                    double hsum = 0;
                                    for (int li2 = start; li2 < lines.Count; li2++) {
                                        double lineHeight = GetRichLineHeight(listItem.Heights, li2, leading);
                                        if (hsum + lineHeight > availableForLines) break;
                                        hsum += lineHeight;
                                        take++;
                                    }
                                    if (take == 0) break;

                                    var sliceLines = new System.Collections.Generic.List<System.Collections.Generic.List<RichSeg>>(take);
                                    var sliceHeights = new System.Collections.Generic.List<double>(take);
                                    for (int k = 0; k < take; k++) {
                                        sliceLines.Add(lines[start + k]);
                                        sliceHeights.Add(GetRichLineHeight(listItem.Heights, start + k, leading));
                                    }

                                    pageDirty = true;
                                    var listFont = ChooseNormal(currentOpts.DefaultFont);
                                    double baselineY = FirstTextBaselineFromTop(listFont, listItem.Size, yCol);
                                    if (line == 0) {
                                        if (!string.IsNullOrEmpty(listItem.BookmarkName)) {
                                            AddNamedDestinationName(listItem.BookmarkName!, yCol);
                                        }

                                        var markerLines = new System.Collections.Generic.List<string>(1) { listItem.Marker };
                                        WriteLinesInternal("F1", listItem.Size, leading, xCol + listItem.MarkerXOffset, listItem.MarkerWidth, baselineY, markerLines, listItem.MarkerAlign, listItem.Color, applyBaselineTweak: true);
                                    }

                                    WriteRichParagraph(sb, new RichParagraphBlock(listItem.Runs, listItem.TextAlign, listItem.Color), sliceLines, sliceHeights, currentOpts, baselineY, listItem.Size, leading, currentPage!.Annotations, xCol + listItem.TextXOffset, listItem.TextWidth);
                                    MarkRichFonts(listItem.Runs);
                                    yCol -= hsum;
                                    remain -= hsum;
                                    consumed += hsum;
                                    line += take;
                                    if (line >= lines.Count) {
                                        double space = listItem.SpacingAfter;
                                        if (space <= remain) {
                                            yCol -= space;
                                            remain -= space;
                                            consumed += space;
                                        }

                                        idx++;
                                        line = 0;
                                    }
                                } else if (it is ColPanel panel) {
                                    var pblock = panel.Block;
                                    var panelStyle = panel.Style;
                                    var lines = panel.Lines;
                                    var heights = panel.Heights;
                                    double xPanel = xCol + panel.XOffset;
                                    double spacingBefore = line == 0 ? ResolveColumnSpacingBefore(panelStyle.SpacingBefore, consumed) : 0D;
                                    if (line == 0 && panelStyle.KeepWithNext && idx + 1 < items.Count) {
                                        double nextHeight = MeasureColItemFirstVisualHeight(items[idx + 1]);
                                        double panelHeight = spacingBefore + panelStyle.PaddingY + heights.Sum() + panelStyle.PaddingY + panelStyle.SpacingAfter;
                                        double keepHeight = panelHeight + nextHeight;
                                        double availableHeight = currentOpts.PageHeight - currentOpts.MarginTop - currentOpts.MarginBottom;
                                        if (nextHeight > 0.001 && keepHeight <= availableHeight + 0.001 && keepHeight > remain + 0.001) {
                                            if (consumed > 0) break;
                                            remain = 0;
                                            break;
                                        }
                                    }

                                    if (line == 0 && spacingBefore > 0) {
                                        if (spacingBefore > remain && consumed > 0) break;
                                        if (spacingBefore > remain && consumed == 0) { remain = 0; break; }
                                        yCol -= spacingBefore;
                                        remain -= spacingBefore;
                                        consumed += spacingBefore;
                                    }

                                    if (panelStyle.KeepTogether) {
                                        double textHeight = heights.Sum();
                                        double panelHeight = panelStyle.PaddingY + textHeight + panelStyle.PaddingY;
                                        double availableHeight = currentOpts.PageHeight - currentOpts.MarginTop - currentOpts.MarginBottom;
                                        if (panelHeight > availableHeight + 0.001) {
                                            throw new ArgumentException("Panel height exceeds the available page content height.");
                                        }

                                        if (panelHeight > remain && consumed > 0) break;
                                        if (panelHeight > remain && consumed == 0) { remain = 0; break; }

                                        double panelTop = yCol;
                                        double panelBottom = yCol - panelHeight;
                                        if (panelStyle.Background.HasValue) { pageDirty = true; DrawRowFill(sb, panelStyle.Background.Value, xPanel, panelBottom, panel.PanelWidth, panelTop - panelBottom); }
                                        if (DrawPanelBorder(sb, panelStyle, xPanel, panelBottom, panel.PanelWidth, panelTop - panelBottom)) { pageDirty = true; }
                                        pageDirty = true;
                                        WriteRichParagraph(sb, new RichParagraphBlock(pblock.Runs, pblock.Align, pblock.DefaultColor), lines, heights, currentOpts, panelTop - panelStyle.PaddingY - panel.FirstBaselineOffset, panel.Size, panel.Leading, currentPage!.Annotations, xPanel + panelStyle.PaddingX, panel.TextWidth);
                                        MarkRichFonts(pblock.Runs);

                                        yCol = panelBottom;
                                        remain -= panelHeight;
                                        consumed += panelHeight;
                                        if (panelStyle.SpacingAfter > 0 && panelStyle.SpacingAfter <= remain) {
                                            yCol -= panelStyle.SpacingAfter;
                                            remain -= panelStyle.SpacingAfter;
                                            consumed += panelStyle.SpacingAfter;
                                        }
                                        idx++;
                                        line = 0;
                                    } else {
                                        int start = line;
                                        double topPad = start == 0 ? panelStyle.PaddingY : 0;
                                        double minLine = heights[start];
                                        if (remain < topPad + minLine) {
                                            if (consumed > 0) break;
                                            remain = 0;
                                            break;
                                        }

                                        double roomForText = remain - topPad - panelStyle.PaddingY;
                                        if (roomForText < minLine) {
                                            roomForText = remain - topPad;
                                        }

                                        int take = 0;
                                        double hsum = 0;
                                        for (int k = start; k < lines.Count; k++) {
                                            double h = heights[k];
                                            if (hsum + h > roomForText) break;
                                            hsum += h;
                                            take++;
                                        }

                                        if (take == 0) break;

                                        bool lastSeg = start + take >= lines.Count;
                                        double panelTop = yCol;
                                        double usedBottomPad = lastSeg ? panelStyle.PaddingY : Math.Max(0, remain - (topPad + hsum));
                                        double panelBottom = yCol - (topPad + hsum + usedBottomPad);
                                        if (panelStyle.Background.HasValue) { pageDirty = true; DrawRowFill(sb, panelStyle.Background.Value, xPanel, panelBottom, panel.PanelWidth, panelTop - panelBottom); }
                                        if (DrawPanelBorder(sb, panelStyle, xPanel, panelBottom, panel.PanelWidth, panelTop - panelBottom)) { pageDirty = true; }

                                        var sliceLines = new System.Collections.Generic.List<System.Collections.Generic.List<RichSeg>>();
                                        var sliceHeights = new System.Collections.Generic.List<double>();
                                        for (int k = 0; k < take; k++) {
                                            sliceLines.Add(lines[start + k]);
                                            sliceHeights.Add(heights[start + k]);
                                        }

                                        pageDirty = true;
                                        WriteRichParagraph(sb, new RichParagraphBlock(pblock.Runs, pblock.Align, pblock.DefaultColor), sliceLines, sliceHeights, currentOpts, panelTop - topPad - panel.FirstBaselineOffset, panel.Size, panel.Leading, currentPage!.Annotations, xPanel + panelStyle.PaddingX, panel.TextWidth);
                                        MarkRichFonts(pblock.Runs);

                                        double segmentHeight = panelTop - panelBottom;
                                        yCol = panelBottom;
                                        remain -= segmentHeight;
                                        consumed += segmentHeight;
                                        line += take;
                                        if (line >= lines.Count) {
                                            if (panelStyle.SpacingAfter > 0 && panelStyle.SpacingAfter <= remain) {
                                                yCol -= panelStyle.SpacingAfter;
                                                remain -= panelStyle.SpacingAfter;
                                                consumed += panelStyle.SpacingAfter;
                                            }
                                            idx++;
                                            line = 0;
                                        } else {
                                            break;
                                        }
                                    }
                                } else if (it is ColTable table) {
                                    var tbColumn = table.Block;
                                    var tableStyle = table.Style;
                                    double padLeft = GetTableCellPaddingLeft(tableStyle);
                                    double padRight = GetTableCellPaddingRight(tableStyle);
                                    double padTop = GetTableCellPaddingTop(tableStyle);
                                    double padBottom = GetTableCellPaddingBottom(tableStyle);
                                    double columnGap = GetTableCellSpacing(tableStyle);
                                    double columnTableRowGap = columnGap;
                                    double xTable = ResolveTableX(tbColumn.Align, tableStyle, xCol, wCol, table.Width);

                                    double maxContentHeight = currentOpts.PageHeight - currentOpts.MarginTop - currentOpts.MarginBottom;
                                    double tableSpacingBefore = line == 0 && consumed > 0.001 ? tableStyle.SpacingBefore : 0D;
                                    if (line == 0 && tableStyle.KeepTogether) {
                                        double keepHeight = tableSpacingBefore + table.CaptionHeight + GetTableRowsHeight(table.RowHeights, 0, table.RowHeights.Length, columnTableRowGap) + tableStyle.SpacingAfter;
                                        if (keepHeight > maxContentHeight + 0.001) {
                                            throw new ArgumentException("Table height exceeds the available page content height.");
                                        }

                                        if (keepHeight > remain + 0.001) {
                                            if (consumed > 0) break;
                                            remain = 0;
                                            break;
                                        }
                                    }

                                    if (line == 0 && tableStyle.KeepWithNext && idx + 1 < items.Count) {
                                        double tableHeight = tableSpacingBefore + table.CaptionHeight + GetTableRowsHeight(table.RowHeights, 0, table.RowHeights.Length, columnTableRowGap) + tableStyle.SpacingAfter;
                                        double nextHeight = MeasureColItemFirstVisualHeight(items[idx + 1]);
                                        double keepHeight = tableHeight + nextHeight;
                                        if (nextHeight > 0.001 && tableHeight <= maxContentHeight + 0.001 && keepHeight <= maxContentHeight + 0.001 && keepHeight > remain + 0.001) {
                                            if (consumed > 0) break;
                                            remain = 0;
                                            break;
                                        }
                                    }

                                    if (line == 0 && tableSpacingBefore > 0) {
                                        if (tableSpacingBefore > remain && consumed > 0) break;
                                        if (tableSpacingBefore > remain && consumed == 0) { remain = 0; break; }
                                        yCol -= tableSpacingBefore;
                                        remain -= tableSpacingBefore;
                                        consumed += tableSpacingBefore;
                                    }

                                    if (line == 0 && table.CaptionLines != null) {
                                        double firstRowHeight = table.RowHeights.Length > 0 ? table.RowHeights[0] : 0;
                                        double neededWithFirstRow = table.CaptionHeight + firstRowHeight;
                                        if (neededWithFirstRow > maxContentHeight + 0.001) {
                                            throw new ArgumentException("Table caption and first row exceed the available page content height.");
                                        }
                                        if (neededWithFirstRow > remain && consumed > 0) break;
                                        if (neededWithFirstRow > remain && consumed == 0) { remain = 0; break; }

                                        double captionSize = tableStyle.CaptionFontSize ?? table.Size;
                                        var captionFont = ChooseNormal(currentOpts.DefaultFont);
                                        pageDirty = true;
                                        WriteLinesInternal("F1", captionSize, table.CaptionLeading, xTable, table.Width, yCol - GetAscender(captionFont, captionSize), table.CaptionLines, tableStyle.CaptionAlign, tableStyle.CaptionColor);
                                        yCol -= table.CaptionHeight;
                                        remain -= table.CaptionHeight;
                                        consumed += table.CaptionHeight;
                                    }

                                    double repeatHeaderHeight = 0;
                                    for (int headerIndex = 0; headerIndex < table.RepeatHeaderRowCount; headerIndex++) {
                                        repeatHeaderHeight += table.RowHeights[headerIndex] + GetTableRowGapAfter(headerIndex, tbColumn.Rows.Count, columnTableRowGap);
                                    }

                                    bool HasRepeatableHeader() =>
                                        table.RepeatHeaderRowCount > 0 &&
                                        tbColumn.Rows.Count > table.HeaderRowCount;

                                    bool AtContinuationPageTop() =>
                                        Math.Abs(yCol - yStart) <= 0.001;

                                    void DrawColumnTableRowSegment(int rowIndex, bool renderAsHeader, int startLine, int lineCount, bool suppressCellObjects = false) {
                                        bool renderAsFooter = rowIndex >= table.FooterStartRowIndex;
                                        bool rowUsesBold = table.RowBold[rowIndex];
                                        double rowSize = table.RowSizes[rowIndex];
                                        double rowLeading = table.RowLeadings[rowIndex];
                                        bool wholeRowSegment = startLine == 0 && lineCount == table.RowLineCounts[rowIndex];
                                        double rowPadTop = GetTableRowMaxPaddingTop(tbColumn, tableStyle, rowIndex, table.Columns);
                                        double rowPadBottom = GetTableRowMaxPaddingBottom(tbColumn, tableStyle, rowIndex, table.Columns);
                                        double rowHeight = wholeRowSegment ? table.RowHeights[rowIndex] : Math.Max(1, lineCount) * rowLeading + rowPadTop + rowPadBottom;
                                        if (rowUsesBold) {
                                            currentPage!.UsedBold = true;
                                            usedBold = true;
                                        }

                                        var cells = GetTableCellLayouts(tbColumn, rowIndex, table.Columns);
                                        double rowBottom = yCol - rowHeight;
                                        int bodyRowIndex = rowIndex - table.HeaderRowCount;
                                        bool stripeBodyRow = bodyRowIndex >= 0 && bodyRowIndex % 2 == 1;
                                        bool[] rowFillSkips = GetRowSpanContinuationSkipColumns(tbColumn, rowIndex, table.Columns);
                                        if (tableStyle.HeaderFill is not null && renderAsHeader) { pageDirty = true; DrawTableRowFill(sb, tableStyle.HeaderFill.Value, xTable, table.ColumnWidths, columnGap, rowBottom, rowHeight, rowFillSkips); }
                                        else if (tableStyle.FooterFill is not null && renderAsFooter) { pageDirty = true; DrawTableRowFill(sb, tableStyle.FooterFill.Value, xTable, table.ColumnWidths, columnGap, rowBottom, rowHeight, rowFillSkips); }
                                        else if (!renderAsHeader && !renderAsFooter && tableStyle.RowStripeFill is not null && stripeBodyRow) { pageDirty = true; DrawTableRowFill(sb, tableStyle.RowStripeFill.Value, xTable, table.ColumnWidths, columnGap, rowBottom, rowHeight, rowFillSkips); }

                                        if (!renderAsHeader && !renderAsFooter && tableStyle.BodyColumnFills != null) {
                                            bool[] bodyColumnFillSkips = GetMergedCellContinuationSkipColumns(tbColumn, rowIndex, table.Columns);
                                            double fillX = xTable;
                                            for (int fillColumn = 0; fillColumn < table.Columns; fillColumn++) {
                                                PdfColor? fill = fillColumn < tableStyle.BodyColumnFills.Count ? tableStyle.BodyColumnFills[fillColumn] : null;
                                                if (fill.HasValue && (fillColumn >= bodyColumnFillSkips.Length || !bodyColumnFillSkips[fillColumn])) {
                                                    pageDirty = true;
                                                    DrawRowFill(sb, fill.Value, fillX, rowBottom, table.ColumnWidths[fillColumn], rowHeight);
                                                }
                                                fillX += table.ColumnWidths[fillColumn] + columnGap;
                                            }
                                        }

                                        if (tableStyle.CellFills != null && tableStyle.CellFills.Count > 0) {
                                            double fillX = xTable;
                                            for (int fillColumn = 0; fillColumn < table.Columns; fillColumn++) {
                                                if (tableStyle.CellFills.TryGetValue((rowIndex, fillColumn), out PdfColor fill) &&
                                                    TryGetTableCellLayoutAtColumn(cells, fillColumn, out TableCellLayout fillCell) &&
                                                    (fillColumn >= rowFillSkips.Length || !rowFillSkips[fillColumn])) {
                                                    int span = wholeRowSegment ? fillCell.ColumnSpan : 1;
                                                    double fillHeight = rowHeight;
                                                    double fillBottom = rowBottom;
                                                    if (wholeRowSegment) {
                                                        if (fillCell.RowSpan > 1) {
                                                            fillHeight = GetTableCellHeight(table.RowHeights, rowIndex, fillCell.RowSpan, columnTableRowGap);
                                                            fillBottom = yCol - fillHeight;
                                                        }
                                                    }

                                                    pageDirty = true;
                                                    DrawRowFill(sb, fill, fillX, fillBottom, GetTableCellWidth(table.ColumnWidths, fillColumn, span, columnGap), fillHeight);
                                                }
                                                fillX += table.ColumnWidths[fillColumn] + columnGap;
                                            }
                                        }
                                        if (DrawTableCellDataBars(sb, tableStyle, cells, rowIndex, table.Columns, xTable, yCol, rowBottom, rowHeight, table.ColumnWidths, columnGap, table.RowHeights, columnTableRowGap, wholeRowSegment, startLine, rowFillSkips)) {
                                            pageDirty = true;
                                        }
                                        if (DrawTableCellIcons(sb, tableStyle, cells, rowIndex, table.Columns, xTable, yCol, rowBottom, rowHeight, table.ColumnWidths, columnGap, table.RowHeights, columnTableRowGap, wholeRowSegment, startLine, rowFillSkips)) {
                                            pageDirty = true;
                                        }

                                        var textColor = renderAsHeader ? tableStyle.HeaderTextColor : renderAsFooter ? tableStyle.FooterTextColor : tableStyle.TextColor;
                                        double xi = xTable;
                                        for (int cellIndex = 0; cellIndex < cells.Count; cellIndex++) {
                                            TableCellLayout cell = cells[cellIndex];
                                            int c = cell.Column;
                                            xi = xTable;
                                            for (int xColumn = 0; xColumn < c; xColumn++) {
                                                xi += table.ColumnWidths[xColumn] + columnGap;
                                            }

                                            double cellWidth = GetTableCellWidth(table.ColumnWidths, c, cell.ColumnSpan, columnGap);
                                            double cellPadLeft = GetTableCellPaddingLeft(tableStyle, rowIndex, c);
                                            double cellPadRight = GetTableCellPaddingRight(tableStyle, rowIndex, c);
                                            double cellPadTop = GetTableCellPaddingTop(tableStyle, rowIndex, c);
                                            double cellPadBottom = GetTableCellPaddingBottom(tableStyle, rowIndex, c);
                                            double innerW = cellWidth - cellPadLeft - cellPadRight;
                                            double cellHeight = wholeRowSegment && cell.RowSpan > 1 ? GetTableCellHeight(table.RowHeights, rowIndex, cell.RowSpan, columnTableRowGap) : rowHeight;
                                            double cellBottom = yCol - cellHeight;
                                            PdfColumnAlign align = GetTableCellAlignment(tableStyle, rowIndex, c, cell.Text);
                                            PdfCellVerticalAlign verticalAlign = GetTableCellVerticalAlignment(tableStyle, rowIndex, c);
                                            var cellFont = GetTableRowFont(currentOpts, rowUsesBold);
                                            TableCellTextLayout lines = table.RowLines[rowIndex][c];
                                            int sourceStartLine = wholeRowSegment && cell.RowSpan > 1 ? 0 : startLine;
                                            int requestedLineCount = wholeRowSegment && cell.RowSpan > 1 ? lines.LineCount : lineCount;
                                            int visibleLineCount = Math.Max(0, Math.Min(requestedLineCount, lines.LineCount - sourceStartLine));
                                            double verticalOffset = 0;
                                            double visibleTextHeight = 0D;
                                            if (visibleLineCount > 0) {
                                                double availableTextHeight = Math.Max(0, cellHeight - cellPadTop - cellPadBottom);
                                                visibleTextHeight = MeasureTableCellTextHeight(lines, sourceStartLine, visibleLineCount, rowLeading);
                                                double visibleContentHeight = MeasureTableCellContentHeight(cell, lines, sourceStartLine, visibleLineCount, rowLeading);
                                                double unusedTextHeight = Math.Max(0, availableTextHeight - visibleContentHeight);
                                                if (verticalAlign == PdfCellVerticalAlign.Middle) verticalOffset = unusedTextHeight / 2;
                                                else if (verticalAlign == PdfCellVerticalAlign.Bottom) verticalOffset = unusedTextHeight;
                                            }

                                            double firstBaseline = yCol - cellPadTop - verticalOffset - GetAscender(cellFont, rowSize) + tableStyle.RowBaselineOffset;

                                            pageDirty = true;
                                            if (cell.Runs.Any(run => run.Bold || rowUsesBold)) { currentPage!.UsedBold = true; usedBold = true; }
                                            if (cell.Runs.Any(run => run.Italic)) { currentPage!.UsedItalic = true; usedItalic = true; }
                                            if (cell.Runs.Any(run => (run.Bold || rowUsesBold) && run.Italic)) { currentPage!.UsedBoldItalic = true; usedBoldItalic = true; }
                                            MarkRichFonts(cell.Runs);
                                            string? linkUri = cell.LinkUri;
                                            string? linkDestinationName = cell.LinkDestinationName;
                                            string? linkContents = cell.LinkContents;
                                            if (tbColumn.Links.TryGetValue((rowIndex, c), out var uri)) {
                                                linkUri = uri;
                                                linkDestinationName = null;
                                                linkContents = cell.Text;
                                            }

                                            if (sourceStartLine == 0) {
                                                AddTableCellNamedDestinationName(cell.NamedDestinationName, yCol);
                                            }

                                            if (visibleLineCount > 0) {
                                                var visibleLines = SliceTableCellLines(lines, sourceStartLine, visibleLineCount);
                                                visibleLines = StripRichLineLinksWhenCellLinked(visibleLines, linkUri, linkDestinationName);
                                                var visibleHeights = SliceTableCellLineHeights(lines, sourceStartLine, visibleLineCount, rowLeading);
                                                var paragraph = new RichParagraphBlock(StripRunLinksWhenCellLinked(cell.Runs, linkUri, linkDestinationName), MapTableCellAlignment(align), textColor);
                                                WriteClippedRichParagraph(sb, paragraph, visibleLines, visibleHeights, currentOpts, firstBaseline, rowSize, rowLeading, currentPage!.Annotations, xi - TableCellClipBleed, cellBottom - TableCellClipBleed, cellWidth + (TableCellClipBleed * 2D), cellHeight + (TableCellClipBleed * 2D), xi + cellPadLeft, innerW);
                                            }
                                            if (!suppressCellObjects && (cell.Images.Count > 0 || cell.CheckBoxes.Count > 0 || cell.FormFields.Count > 0) && sourceStartLine == 0) {
                                                if (CanRenderTableCellCheckBoxInline(cell, lines, sourceStartLine, visibleLineCount)) {
                                                    RenderTableCellInlineCheckBox(currentPage!, cell, align, lines.Lines[sourceStartLine], xi + cellPadLeft, innerW, firstBaseline);
                                                } else {
                                                    double formFieldTop = yCol - cellPadTop - verticalOffset - (string.IsNullOrEmpty(cell.Text) ? 0D : visibleTextHeight + TableCellCheckBoxGap);
                                                    RenderTableCellObjects(currentPage!, cell, align, xi + cellPadLeft, innerW, formFieldTop);
                                                }
                                            }

                                            if (HasCellLinkTarget(linkUri, linkDestinationName)) {
                                                currentPage!.Annotations.Add(new LinkAnnotation { X1 = xi + cellPadLeft, Y1 = cellBottom, X2 = xi + cellWidth - cellPadRight, Y2 = yCol, Uri = linkUri, DestinationName = linkDestinationName, Contents = linkContents ?? cell.Text });
                                            }
                                        }

                                        if (tableStyle.BorderColor is not null && tableStyle.BorderWidth > 0) {
                                            pageDirty = true;
                                            bool[] topBorderSkips = GetRowSpanBoundarySkipColumns(tbColumn, rowIndex - 1, table.Columns);
                                            bool[] bottomBorderSkips = GetRowSpanBoundarySkipColumns(tbColumn, rowIndex, table.Columns);
                                            bool segmentBorderRows = HasSkippedColumns(topBorderSkips, table.Columns) || HasSkippedColumns(bottomBorderSkips, table.Columns);
                                            if (segmentBorderRows) {
                                                DrawTableHorizontalLine(sb, tableStyle.BorderColor.Value, tableStyle.BorderWidth, xTable, table.ColumnWidths, columnGap, rowBottom + rowHeight, topBorderSkips);
                                                DrawTableHorizontalLine(sb, tableStyle.BorderColor.Value, tableStyle.BorderWidth, xTable, table.ColumnWidths, columnGap, rowBottom, bottomBorderSkips);
                                                DrawVLine(sb, tableStyle.BorderColor.Value, tableStyle.BorderWidth, xTable, rowBottom + rowHeight, rowBottom);
                                                DrawVLine(sb, tableStyle.BorderColor.Value, tableStyle.BorderWidth, xTable + table.Width, rowBottom + rowHeight, rowBottom);
                                            } else {
                                                DrawRowRect(sb, tableStyle.BorderColor.Value, tableStyle.BorderWidth, xTable, rowBottom, table.Width, rowHeight);
                                            }

                                            double xi2 = xTable;
                                            for (int c = 0; c < table.Columns - 1; c++) {
                                                xi2 += table.ColumnWidths[c];
                                                if (IsTableBoundaryInsideSpannedCell(tbColumn, rowIndex, c, table.Columns)) {
                                                    xi2 += columnGap;
                                                    continue;
                                                }

                                                DrawVLine(sb, tableStyle.BorderColor.Value, tableStyle.BorderWidth, xi2, rowBottom + rowHeight, rowBottom);
                                                xi2 += columnGap;
                                            }
                                        }

                                        if (renderAsFooter && rowIndex == table.FooterStartRowIndex) {
                                            PdfColor? footerSeparatorColor = tableStyle.FooterSeparatorColor ?? tableStyle.RowSeparatorColor;
                                            double footerSeparatorWidth = tableStyle.FooterSeparatorWidth > 0 ? tableStyle.FooterSeparatorWidth : tableStyle.RowSeparatorWidth;
                                            if (footerSeparatorColor is not null && footerSeparatorWidth > 0) {
                                                pageDirty = true;
                                                DrawTableHorizontalLine(sb, footerSeparatorColor.Value, footerSeparatorWidth, xTable, table.ColumnWidths, columnGap, yCol, GetRowSpanBoundarySkipColumns(tbColumn, rowIndex - 1, table.Columns));
                                            }
                                        }

                                        PdfColor? separatorColor = renderAsHeader && tableStyle.HeaderSeparatorColor is not null ? tableStyle.HeaderSeparatorColor : tableStyle.RowSeparatorColor;
                                        double separatorWidth = renderAsHeader && tableStyle.HeaderSeparatorWidth > 0 ? tableStyle.HeaderSeparatorWidth : tableStyle.RowSeparatorWidth;
                                        if (separatorColor is not null && separatorWidth > 0) {
                                            pageDirty = true;
                                            DrawTableHorizontalLine(sb, separatorColor.Value, separatorWidth, xTable, table.ColumnWidths, columnGap, rowBottom, GetRowSpanBoundarySkipColumns(tbColumn, rowIndex, table.Columns));
                                        }

                                        if (tableStyle.CellBorders != null && tableStyle.CellBorders.Count > 0) {
                                            double borderX = xTable;
                                            for (int borderColumn = 0; borderColumn < table.Columns; borderColumn++) {
                                                if (tableStyle.CellBorders.TryGetValue((rowIndex, borderColumn), out PdfCellBorder? cellBorder) &&
                                                    TryGetTableCellLayoutAtColumn(cells, borderColumn, out TableCellLayout borderCell) &&
                                                    (borderColumn >= rowFillSkips.Length || !rowFillSkips[borderColumn]) &&
                                                    HasRenderableCellBorder(cellBorder)) {
                                                    int span = wholeRowSegment ? borderCell.ColumnSpan : 1;
                                                    double borderHeight = rowHeight;
                                                    double borderBottom = rowBottom;
                                                    if (wholeRowSegment) {
                                                        if (borderCell.RowSpan > 1) {
                                                            borderHeight = GetTableCellHeight(table.RowHeights, rowIndex, borderCell.RowSpan, columnTableRowGap);
                                                            borderBottom = yCol - borderHeight;
                                                        }
                                                    }

                                                    pageDirty = true;
                                                    DrawCellBorder(sb, cellBorder, borderX, borderBottom, GetTableCellWidth(table.ColumnWidths, borderColumn, span, columnGap), borderHeight);
                                                }
                                                borderX += table.ColumnWidths[borderColumn] + columnGap;
                                            }
                                        }

                                        double rowAdvance = rowHeight + (wholeRowSegment ? GetTableRowGapAfter(rowIndex, tbColumn.Rows.Count, columnTableRowGap) : 0D);
                                        yCol -= rowAdvance;
                                        remain -= rowAdvance;
                                        consumed += rowAdvance;
                                    }

                                    void DrawColumnTableRow(int rowIndex, bool renderAsHeader, bool suppressCellObjects = false) =>
                                        DrawColumnTableRowSegment(rowIndex, renderAsHeader, 0, table.RowLineCounts[rowIndex], suppressCellObjects);

                                    int rowIndex = line;
                                    int rowStartLine = subline;
                                    while (rowIndex < tbColumn.Rows.Count) {
                                        double rowHeight = table.RowHeights[rowIndex];
                                        if (rowHeight > maxContentHeight + 0.001) {
                                            if (!GetTableRowAllowBreakAcrossPages(tableStyle, rowIndex)) {
                                                throw new ArgumentException("Table row height exceeds the available page content height and row splitting is disabled.");
                                            }

                                            int totalLines = table.RowLineCounts[rowIndex];
                                            double rowPadTop = GetTableRowMaxPaddingTop(tbColumn, tableStyle, rowIndex, table.Columns);
                                            double rowPadBottom = GetTableRowMaxPaddingBottom(tbColumn, tableStyle, rowIndex, table.Columns);
                                            bool repeatHeaderBeforeSegment = rowIndex >= table.HeaderRowCount &&
                                                HasRepeatableHeader() &&
                                                AtContinuationPageTop() &&
                                                repeatHeaderHeight + table.RowLeadings[rowIndex] + rowPadTop + rowPadBottom <= remain + 0.001;
                                            double neededForFirstSegment = table.RowLeadings[rowIndex] + rowPadTop + rowPadBottom + (repeatHeaderBeforeSegment ? repeatHeaderHeight : 0);
                                            if (neededForFirstSegment > remain && consumed > 0) break;
                                            if (neededForFirstSegment > remain && consumed == 0) { remain = 0; break; }

                                            if (repeatHeaderBeforeSegment) {
                                                for (int headerIndex = 0; headerIndex < table.RepeatHeaderRowCount; headerIndex++) {
                                                    DrawColumnTableRow(headerIndex, renderAsHeader: true, suppressCellObjects: true);
                                                }
                                            }

                                            int maxLinesThisPage = Math.Max(1, (int)Math.Floor((remain - rowPadTop - rowPadBottom) / table.RowLeadings[rowIndex]));
                                            int take = Math.Min(totalLines - rowStartLine, maxLinesThisPage);
                                            DrawColumnTableRowSegment(rowIndex, renderAsHeader: rowIndex < table.HeaderRowCount && rowStartLine == 0, rowStartLine, take);
                                            rowStartLine += take;

                                            if (rowStartLine < totalLines) {
                                                line = rowIndex;
                                                subline = rowStartLine;
                                                break;
                                            }

                                            double gapAfterSplitRow = GetTableRowGapAfter(rowIndex, tbColumn.Rows.Count, columnTableRowGap);
                                            if (gapAfterSplitRow > 0) {
                                                yCol -= gapAfterSplitRow;
                                                remain -= gapAfterSplitRow;
                                                consumed += gapAfterSplitRow;
                                            }

                                            rowIndex++;
                                            line = rowIndex;
                                            subline = 0;
                                            rowStartLine = 0;
                                            continue;
                                        }
                                        bool repeatHeaderBeforeRow = rowIndex >= table.HeaderRowCount &&
                                            HasRepeatableHeader() &&
                                            AtContinuationPageTop() &&
                                            repeatHeaderHeight + rowHeight <= remain + 0.001;
                                        double neededForNextRow = rowHeight + GetTableRowGapAfter(rowIndex, tbColumn.Rows.Count, columnTableRowGap) + (repeatHeaderBeforeRow ? repeatHeaderHeight : 0);
                                        if (neededForNextRow > remain && consumed > 0) break;
                                        if (neededForNextRow > remain && consumed == 0) { remain = 0; break; }

                                        if (repeatHeaderBeforeRow) {
                                            for (int headerIndex = 0; headerIndex < table.RepeatHeaderRowCount; headerIndex++) {
                                                DrawColumnTableRow(headerIndex, renderAsHeader: true, suppressCellObjects: true);
                                            }
                                        }

                                        DrawColumnTableRow(rowIndex, renderAsHeader: rowIndex < table.HeaderRowCount);
                                        rowIndex++;
                                        line = rowIndex;
                                        subline = 0;
                                        rowStartLine = 0;
                                    }

                                    if (rowIndex >= tbColumn.Rows.Count) {
                                        if (tableStyle.SpacingAfter > 0 && tableStyle.SpacingAfter <= remain) {
                                            yCol -= tableStyle.SpacingAfter;
                                            remain -= tableStyle.SpacingAfter;
                                            consumed += tableStyle.SpacingAfter;
                                        }
                                        idx++;
                                        line = 0;
                                        subline = 0;
                                    } else {
                                        break;
                                    }
                                } else if (it is ColRule cr) {
                                    PdfHorizontalRuleStyle hr2 = ResolveHorizontalRuleStyle(cr.Block, currentOpts);
                                    ValidateHorizontalRule(hr2);
                                    double spacingBefore = ResolveColumnSpacingBefore(hr2.SpacingBefore, consumed);
                                    double needed = spacingBefore + hr2.Thickness + hr2.SpacingAfter;
                                    EnsureFixedFlowBlockFits("Horizontal rule", wCol, needed, wCol);
                                    if (line == 0 && hr2.KeepWithNext && idx + 1 < items.Count) {
                                        double nextHeight = MeasureColItemFirstVisualHeight(items[idx + 1]);
                                        double keepHeight = needed + nextHeight;
                                        double availableHeight = currentOpts.PageHeight - currentOpts.MarginTop - currentOpts.MarginBottom;
                                        if (nextHeight > 0.001 && keepHeight <= availableHeight + 0.001 && keepHeight > remain + 0.001) {
                                            if (consumed > 0) break;
                                            remain = 0;
                                            break;
                                        }
                                    }

                                    if (needed > remain && consumed > 0) break;
                                    if (needed > remain && consumed == 0) { remain = 0; break; }
                                    if (spacingBefore > 0) yCol -= spacingBefore;
                                    double x1 = xCol, x2 = xCol + wCol, yLine = yCol - hr2.Thickness * 0.5;
                                    pageDirty = true;
                                    DrawHLine(sb, hr2.Color, hr2.Thickness, x1, x2, yLine);
                                    yCol -= hr2.Thickness + hr2.SpacingAfter; remain -= needed; consumed += needed; idx++;
                                } else if (it is ColImg ciimg) {
                                    var ib2 = ciimg.Block;
                                    PdfImageStyle imageStyle = ResolveImageStyle(ib2, currentOpts);
                                    PdfDoc.ValidateImageStyleForBox(imageStyle, ib2.Width, ib2.Height, nameof(imageStyle.ClipPath));
                                    PdfDoc.ValidateImageFitDimensions(ib2.Info, imageStyle.Fit, nameof(imageStyle.Fit));
                                    double spacingBefore = ResolveColumnSpacingBefore(imageStyle.SpacingBefore, consumed);
                                    double needed = spacingBefore + ib2.Height + imageStyle.SpacingAfter;
                                    EnsureFixedFlowBlockFits("Image", ib2.Width, needed, wCol);
                                    if (imageStyle.KeepWithNext && idx + 1 < items.Count) {
                                        double nextHeight = MeasureColItemFirstVisualHeight(items[idx + 1]);
                                        double keepHeight = needed + nextHeight;
                                        double availableHeight = currentOpts.PageHeight - currentOpts.MarginTop - currentOpts.MarginBottom;
                                        if (nextHeight > 0.001 && keepHeight <= availableHeight + 0.001 && keepHeight > remain + 0.001) {
                                            if (consumed > 0) break;
                                            remain = 0;
                                            break;
                                        }
                                    }

                                    if (needed > remain && consumed > 0) break;
                                    if (needed > remain && consumed == 0) { remain = 0; break; }
                                    if (spacingBefore > 0) yCol -= spacingBefore;
                                    double xImg = xCol;
                                    if (imageStyle.Align == PdfAlign.Center) xImg = xCol + Math.Max(0, (wCol - ib2.Width) / 2);
                                    else if (imageStyle.Align == PdfAlign.Right) xImg = xCol + Math.Max(0, wCol - ib2.Width);
                                    PageImage pageImage = CreatePageImage(ib2, imageStyle, xImg, yCol - ib2.Height);
                                    currentPage!.Images.Add(pageImage);
                                    AddImageLinkAnnotation(ib2, imageStyle, pageImage, xImg, yCol - ib2.Height);
                                    pageDirty = true;
                                    yCol -= ib2.Height + imageStyle.SpacingAfter; remain -= needed; consumed += needed; idx++;
                                } else if (it is ColShape cs) {
                                    var shape = cs.Block;
                                    PdfDrawingStyle shapeStyle = ResolveDrawingStyle(shape, currentOpts);
                                    PdfDoc.ValidateDrawingStyle(shapeStyle, "Shape");
                                    double spacingBefore = ResolveColumnSpacingBefore(shapeStyle.SpacingBefore, consumed);
                                    double needed = spacingBefore + shape.Shape.Height + shapeStyle.SpacingAfter;
                                    EnsureFixedFlowBlockFits("Shape", shape.Shape.Width, needed, wCol);
                                    if (shapeStyle.KeepWithNext && idx + 1 < items.Count) {
                                        double nextHeight = MeasureColItemFirstVisualHeight(items[idx + 1]);
                                        double keepHeight = needed + nextHeight;
                                        double availableHeight = currentOpts.PageHeight - currentOpts.MarginTop - currentOpts.MarginBottom;
                                        if (nextHeight > 0.001 && keepHeight <= availableHeight + 0.001 && keepHeight > remain + 0.001) {
                                            if (consumed > 0) break;
                                            remain = 0;
                                            break;
                                        }
                                    }

                                    if (needed > remain && consumed > 0) break;
                                    if (needed > remain && consumed == 0) { remain = 0; break; }
                                    if (spacingBefore > 0) yCol -= spacingBefore;
                                    DrawShapeAt(shape, shapeStyle, xCol, wCol, yCol);
                                    AddShapeLinkAnnotation(shape, shapeStyle, xCol, wCol, yCol);
                                    yCol -= shape.Shape.Height + shapeStyle.SpacingAfter;
                                    remain -= needed;
                                    consumed += needed;
                                    idx++;
                                } else if (it is ColDrawing cd) {
                                    var drawing = cd.Block;
                                    PdfDrawingStyle drawingStyle = ResolveDrawingStyle(drawing, currentOpts);
                                    PdfDoc.ValidateDrawingStyle(drawingStyle, "Drawing");
                                    double spacingBefore = ResolveColumnSpacingBefore(drawingStyle.SpacingBefore, consumed);
                                    double needed = spacingBefore + drawing.Drawing.Height + drawingStyle.SpacingAfter;
                                    EnsureFixedFlowBlockFits("Drawing", drawing.Drawing.Width, needed, wCol);
                                    if (drawingStyle.KeepWithNext && idx + 1 < items.Count) {
                                        double nextHeight = MeasureColItemFirstVisualHeight(items[idx + 1]);
                                        double keepHeight = needed + nextHeight;
                                        double availableHeight = currentOpts.PageHeight - currentOpts.MarginTop - currentOpts.MarginBottom;
                                        if (nextHeight > 0.001 && keepHeight <= availableHeight + 0.001 && keepHeight > remain + 0.001) {
                                            if (consumed > 0) break;
                                            remain = 0;
                                            break;
                                        }
                                    }

                                    if (needed > remain && consumed > 0) break;
                                    if (needed > remain && consumed == 0) { remain = 0; break; }
                                    if (spacingBefore > 0) yCol -= spacingBefore;
                                    DrawDrawingAt(drawing, drawingStyle, xCol, wCol, yCol);
                                    AddDrawingLinkAnnotation(drawing, drawingStyle, xCol, wCol, yCol);
                                    yCol -= drawing.Drawing.Height + drawingStyle.SpacingAfter;
                                    remain -= needed;
                                    consumed += needed;
                                    idx++;
                                } else if (it is ColForm form) {
                                    double spacingBefore = ResolveColumnSpacingBefore(GetFormFieldSpacingBefore(form.Block), consumed);
                                    double fieldWidth = GetFormFieldWidth(form.Block);
                                    double fieldHeight = GetFormFieldHeight(form.Block);
                                    double spacingAfter = GetFormFieldSpacingAfter(form.Block);
                                    double needed = spacingBefore + fieldHeight + spacingAfter;
                                    EnsureFixedFlowBlockFits(GetFormFieldBlockName(form.Block), fieldWidth, needed, wCol);
                                    if (needed > remain && consumed > 0) break;
                                    if (needed > remain && consumed == 0) { remain = 0; break; }
                                    if (spacingBefore > 0) yCol -= spacingBefore;
                                    double xField = GetAlignedObjectX(xCol, wCol, fieldWidth, GetFormFieldAlign(form.Block));
                                    AddFormFieldAnnotation(form.Block, xField, yCol);
                                    pageDirty = true;
                                    yCol -= fieldHeight + spacingAfter;
                                    remain -= needed;
                                    consumed += needed;
                                    idx++;
                                } else if (it is ColBookmark bookmarkItem) {
                                    AddNamedDestination(bookmarkItem.Block, yCol);
                                    idx++;
                                } else if (it is ColSpacer spacerItem) {
                                    double needed = spacerItem.Block.Height;
                                    double availableHeight = currentOpts.PageHeight - currentOpts.MarginTop - currentOpts.MarginBottom;
                                    if (needed > availableHeight + 0.001) {
                                        throw new ArgumentException("Spacer height exceeds the available page content height.");
                                    }

                                    if (needed > remain && consumed > 0) break;
                                    if (needed > remain && consumed == 0) { remain = 0; break; }
                                    yCol -= needed;
                                    remain -= needed;
                                    consumed += needed;
                                    idx++;
                                }
                            }
                            colStates[ci] = (idx, line, subline);
                            if (colStates[ci] != startState) {
                                anyColumnAdvanced = true;
                            }

                            if (consumed > maxConsumed) maxConsumed = consumed;
                        }

                        if (maxConsumed <= 0.01) {
                            if (anyColumnAdvanced && !AnyRemaining()) {
                                break;
                            }

                            if (Math.Abs(y - yStart) <= 0.001) {
                                throw new InvalidOperationException("Row column layout could not make progress on an empty page.");
                            }

                            NewPage();
                            continue;
                        }
                        DrawRowColumnSeparators(y, y - maxConsumed);
                        y -= maxConsumed;
                    }

                    if (rowSpacingAfter > 0) {
                        y -= rowSpacingAfter;
                    }
                } else if (block is ImageBlock ib) {
                    double xImg = currentOpts.MarginLeft;
                    double contentWidth = currentOpts.PageWidth - currentOpts.MarginLeft - currentOpts.MarginRight;
                    PdfImageStyle imageStyle = ResolveImageStyle(ib, currentOpts);
                    PdfDoc.ValidateImageStyleForBox(imageStyle, ib.Width, ib.Height, nameof(imageStyle.ClipPath));
                    PdfDoc.ValidateImageFitDimensions(ib.Info, imageStyle.Fit, nameof(imageStyle.Fit));
                    double imageSpacingBefore = ResolveTopLevelSpacingBefore(imageStyle.SpacingBefore);
                    double needed = imageSpacingBefore + ib.Height + imageStyle.SpacingAfter;
                    if (imageStyle.Align == PdfAlign.Center) xImg = currentOpts.MarginLeft + Math.Max(0, (contentWidth - ib.Width) / 2);
                    else if (imageStyle.Align == PdfAlign.Right) xImg = currentOpts.MarginLeft + Math.Max(0, contentWidth - ib.Width);
                    EnsureFixedFlowBlockFits("Image", ib.Width, needed, contentWidth);
                    if (imageStyle.KeepWithNext && nextBlock != null) {
                        double nextHeight = MeasureNextBlockFirstVisualHeight(nextBlock, currentOpts.MarginLeft, width, currentOpts.DefaultFontSize);
                        double keepHeight = needed + nextHeight;
                        double availableHeight = currentOpts.PageHeight - currentOpts.MarginTop - currentOpts.MarginBottom;
                        if (nextHeight > 0.001 && keepHeight <= availableHeight + 0.001 && y < yStart - 0.001 && y - keepHeight < currentOpts.MarginBottom) {
                            NewPage();
                            imageSpacingBefore = 0D;
                            needed = ib.Height + imageStyle.SpacingAfter;
                        }
                    }

                    if (y - needed < currentOpts.MarginBottom) {
                        NewPage();
                        imageSpacingBefore = 0D;
                    }
                    if (imageSpacingBefore > 0) y -= imageSpacingBefore;
                    EnsurePage();
                    PageImage pageImage = CreatePageImage(ib, imageStyle, xImg, y - ib.Height);
                    currentPage!.Images.Add(pageImage);
                    AddImageLinkAnnotation(ib, imageStyle, pageImage, xImg, y - ib.Height);
                    pageDirty = true;
                    y -= ib.Height + imageStyle.SpacingAfter;
                } else if (block is PanelParagraphBlock ppb) {
                    double size = currentOpts.DefaultFontSize;
                    double leading = size * 1.4;
                    var panelFont = ChooseNormal(currentOpts.DefaultFont);
                    double firstBaselineOffset = GetAscender(panelFont, size);
                    double contentWidth = currentOpts.PageWidth - currentOpts.MarginLeft - currentOpts.MarginRight;
                    PanelStyle panelStyle = ResolvePanelStyle(ppb, currentOpts);
                    double innerWidth = panelStyle.MaxWidth.HasValue ? Math.Min(contentWidth, panelStyle.MaxWidth.Value) : contentWidth;
                    ValidatePanelStyle(panelStyle, innerWidth);
                    double textWidthAvail = innerWidth - 2 * panelStyle.PaddingX;
                    var (lines, lineHeights) = WrapRichRuns(ppb.Runs, textWidthAvail, size, panelFont, leading);
                    double panelWidth = innerWidth;
                    double xLeft = currentOpts.MarginLeft;
                    if (panelStyle.Align == PdfAlign.Center) xLeft = currentOpts.MarginLeft + Math.Max(0, (contentWidth - innerWidth) / 2);
                    else if (panelStyle.Align == PdfAlign.Right) xLeft = currentOpts.MarginLeft + Math.Max(0, contentWidth - innerWidth);
                    double panelSpacingBefore = ResolveTopLevelSpacingBefore(panelStyle.SpacingBefore);

                    if (panelStyle.KeepWithNext && nextBlock != null && lines.Count > 0) {
                        double panelHeight = panelSpacingBefore + panelStyle.PaddingY + lineHeights.Sum() + panelStyle.PaddingY + panelStyle.SpacingAfter;
                        double nextHeight = MeasureNextBlockFirstVisualHeight(nextBlock, currentOpts.MarginLeft, width, size);
                        double keepHeight = panelHeight + nextHeight;
                        double availableHeight = currentOpts.PageHeight - currentOpts.MarginTop - currentOpts.MarginBottom;
                        if (nextHeight > 0.001 && keepHeight <= availableHeight + 0.001 && y < yStart - 0.001 && y - keepHeight < currentOpts.MarginBottom) {
                            NewPage();
                            panelSpacingBefore = 0D;
                        }
                    }

                    if (panelSpacingBefore > 0) {
                        if (y - panelSpacingBefore < currentOpts.MarginBottom) {
                            NewPage();
                            panelSpacingBefore = 0D;
                        }

                        if (panelSpacingBefore > 0) y -= panelSpacingBefore;
                    }

                    if (panelStyle.KeepTogether) {
                        double textHeight = lineHeights.Sum();
                        double panelHeight = panelStyle.PaddingY + textHeight + panelStyle.PaddingY;
                        double availableHeight = currentOpts.PageHeight - currentOpts.MarginTop - currentOpts.MarginBottom;
                        if (panelHeight > availableHeight + 0.001) {
                            throw new ArgumentException("Panel height exceeds the available page content height.");
                        }

                        double panelTop = y;
                        double panelBottom = y - panelHeight;
                        if (panelBottom < currentOpts.MarginBottom) { NewPage(); panelTop = y; panelBottom = y - panelHeight; }
                        if (panelStyle.Background.HasValue) { pageDirty = true; DrawRowFill(sb, panelStyle.Background.Value, xLeft, panelBottom, panelWidth, panelTop - panelBottom); }
                        if (DrawPanelBorder(sb, panelStyle, xLeft, panelBottom, panelWidth, panelTop - panelBottom)) { pageDirty = true; }
                        pageDirty = true;
                        WriteRichParagraph(sb, new RichParagraphBlock(ppb.Runs, ppb.Align, ppb.DefaultColor), lines, lineHeights, currentOpts, panelTop - panelStyle.PaddingY - firstBaselineOffset, size, leading, currentPage!.Annotations, xLeft + panelStyle.PaddingX, textWidthAvail);
                        MarkRichFonts(ppb.Runs);
                        y = panelBottom;
                        if (panelStyle.SpacingAfter > 0) {
                            if (y < yStart - 0.001 && y - panelStyle.SpacingAfter < currentOpts.MarginBottom) {
                                NewPage();
                            } else {
                                y -= panelStyle.SpacingAfter;
                            }
                        }
                    } else {
                        int li = 0; bool firstSeg = true;
                        while (li < lines.Count) {
                            double avail = y - currentOpts.MarginBottom;
                            if (avail < 0.5) { NewPage(); firstSeg = false; continue; }
                            double topPad = firstSeg ? panelStyle.PaddingY : 0;
                            double minLine = lineHeights[li];
                            if (avail < topPad + minLine) { NewPage(); firstSeg = false; continue; }
                            double roomForText = avail - topPad - panelStyle.PaddingY;
                            int take = 0; double hsum = 0;
                            for (int k = li; k < lines.Count; k++) {
                                double h = lineHeights[k];
                                if (hsum + h > roomForText) break;
                                hsum += h; take++;
                            }
                            bool lastSeg = (li + take) >= lines.Count;
                            double panelTop = y;
                            double usedBottomPad = panelStyle.PaddingY;
                            if (!lastSeg && topPad + hsum + usedBottomPad > avail) usedBottomPad = Math.Max(0, avail - (topPad + hsum));
                            double panelBottom = y - (topPad + hsum + usedBottomPad);
                            if (panelStyle.Background.HasValue) { pageDirty = true; DrawRowFill(sb, panelStyle.Background.Value, xLeft, panelBottom, panelWidth, panelTop - panelBottom); }
                            if (DrawPanelBorder(sb, panelStyle, xLeft, panelBottom, panelWidth, panelTop - panelBottom)) { pageDirty = true; }
                            var sliceLines = new System.Collections.Generic.List<System.Collections.Generic.List<RichSeg>>();
                            var sliceHeights = new System.Collections.Generic.List<double>();
                            for (int k = 0; k < take; k++) { sliceLines.Add(lines[li + k]); sliceHeights.Add(lineHeights[li + k]); }
                            pageDirty = true;
                            WriteRichParagraph(sb, new RichParagraphBlock(ppb.Runs, ppb.Align, ppb.DefaultColor), sliceLines, sliceHeights, currentOpts, panelTop - topPad - firstBaselineOffset, size, leading, currentPage!.Annotations, xLeft + panelStyle.PaddingX, textWidthAvail);
                            MarkRichFonts(ppb.Runs);
                            y = panelBottom; li += take; firstSeg = false;
                            if (li < lines.Count) {
                                NewPage();
                            } else if (panelStyle.SpacingAfter > 0) {
                                if (y < yStart - 0.001 && y - panelStyle.SpacingAfter < currentOpts.MarginBottom) {
                                    NewPage();
                                } else {
                                    y -= panelStyle.SpacingAfter;
                                }
                            }
                        }
                    }
                }
            }
        }

        ProcessBlocks(blocks);
        FlushPage(pageDirty || HasCurrentPageNonContentObjects());

        var result = new LayoutResult { UsedBold = usedBold, UsedItalic = usedItalic, UsedBoldItalic = usedBoldItalic };
        foreach (var p in pages) result.Pages.Add(p);
        return result;
    }

}

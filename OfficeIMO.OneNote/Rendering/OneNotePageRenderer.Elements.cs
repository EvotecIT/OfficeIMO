using OfficeIMO.Drawing;

namespace OfficeIMO.OneNote;

public static partial class OneNotePageRenderer {
    private sealed class RenderContext {
        private const double DefaultParagraphHeight = 20D;
        private readonly OfficeDrawing _drawing;
        private readonly OneNotePageRenderingOptions _options;
        private readonly IList<OfficeImageExportDiagnostic> _diagnostics;
        private readonly OfficeTextMeasurer _measurer;
        private readonly bool _pageRightToLeft;

        internal RenderContext(
            OfficeDrawing drawing,
            OneNotePageRenderingOptions options,
            IList<OfficeImageExportDiagnostic> diagnostics,
            bool pageRightToLeft) {
            _drawing = drawing;
            _options = options;
            _diagnostics = diagnostics;
            _measurer = OfficeTextMeasurer.Create(options.DefaultFont);
            _pageRightToLeft = pageRightToLeft;
        }

        internal double RenderOutline(OneNoteOutline outline, double x, double y, double width, bool? inheritedRightToLeft = null) {
            bool rightToLeft = outline.Layout?.RightToLeft ?? inheritedRightToLeft ?? _pageRightToLeft;
            double cursor = y;
            double pendingSpaceAfter = 0D;
            foreach (OneNoteElement child in outline.Children) {
                bool participatesInFlow = child.Layout?.Y.HasValue != true;
                double childX = child.Layout?.X.HasValue == true ? x + child.Layout.X.Value * PointsPerHalfInch : x;
                double childY = child.Layout?.Y.HasValue == true
                    ? y + child.Layout.Y.Value * PointsPerHalfInch
                    : cursor + Math.Max(pendingSpaceAfter, ParagraphSpaceBefore(child));
                double available = child.Layout?.Width.HasValue == true
                    ? child.Layout.Width.Value * PointsPerHalfInch
                    : width - Math.Max(0D, childX - x);
                double used = RenderElement(child, childX, childY, available, 0D, inheritedRightToLeft: rightToLeft);
                if (participatesInFlow) {
                    cursor = Math.Max(cursor, childY + used);
                    pendingSpaceAfter = child is OneNoteParagraph ? ParagraphSpaceAfter(child) : 5D;
                }
            }
            return Math.Max(DefaultParagraphHeight, cursor + pendingSpaceAfter - y);
        }

        internal double RenderElement(
            OneNoteElement element,
            double x,
            double y,
            double availableWidth,
            double availableHeight,
            bool forcePageBounds = false,
            bool? inheritedRightToLeft = null) {
            if (y >= _drawing.Height || x >= _drawing.Width) return 0D;
            x = Math.Max(0D, x);
            y = Math.Max(0D, y);
            availableWidth = Math.Min(availableWidth, _drawing.Width - x);
            if (availableWidth <= 0D) return 0D;
            bool rightToLeft = element.Layout?.RightToLeft ?? inheritedRightToLeft ?? _pageRightToLeft;
            if (element is OneNoteOutline outline) return RenderOutline(outline, x, y, availableWidth, rightToLeft);
            if (element is OneNoteParagraph paragraph) return RenderParagraph(paragraph, x, y, availableWidth, rightToLeft);
            if (element is OneNoteTable table) return RenderTable(table, x, y, availableWidth, rightToLeft);
            if (element is OneNoteImage image) return RenderImage(image, x, y, availableWidth, availableHeight, forcePageBounds);
            if (element is OneNoteInk ink) return RenderInk(ink, x, y, availableWidth, availableHeight);
            if (element is OneNoteMath math) return RenderMath(math, x, y, availableWidth, availableHeight);
            if (element is OneNoteBinaryElement binary) return RenderAttachment(binary, x, y, availableWidth);
            AddDiagnostic("ONENOTE_RENDER_UNSUPPORTED_ELEMENT", "A OneNote element was not projected to the Drawing scene.", element.Kind.ToString());
            return 0D;
        }

        private double RenderParagraph(OneNoteParagraph paragraph, double x, double y, double width, bool rightToLeft) {
            if (paragraph.Runs.Count == 0) return DefaultParagraphHeight;
            string prefix = CreateParagraphPrefix(paragraph);
            bool hasMath = paragraph.Runs.Any(run => run.MathExpression != null);
            if (!hasMath) {
                IReadOnlyList<OfficeRichTextRun> runs = CreateParagraphRichTextRuns(paragraph, prefix);
                double fontSize = paragraph.Runs.Max(run => run.Style.FontSize ?? _options.DefaultFont.Size);
                double lineHeight = ResolveParagraphLineHeight(paragraph, fontSize);
                double height = MeasureRichTextHeight(runs, width, lineHeight, CreateParagraphIndent(paragraph));
                height = Math.Min(height, Math.Max(1D, _drawing.Height - y));
                if (height <= 0D) return 0D;
                _drawing.AddRichText(
                    runs,
                    x,
                    y,
                    width,
                    height,
                    MapAlignment(paragraph.Style.Alignment, rightToLeft),
                    lineHeight: lineHeight,
                    wrapText: true,
                    paragraphIndent: CreateParagraphIndent(paragraph));
                return height;
            }

            return RenderInlineMathParagraph(paragraph, prefix, x, y, width, rightToLeft);
        }

        private double RenderInlineMathParagraph(OneNoteParagraph paragraph, string prefix, double x, double y, double width, bool rightToLeft) {
            IReadOnlyList<InlineMathLine> lines = CreateInlineMathLines(paragraph, prefix, width);
            double exactLineHeight = ParagraphDistance(paragraph.Style.ExactLineSpacing);
            double cursorY = y;
            foreach (InlineMathLine visualLine in lines) {
                double available = Math.Max(1D, width - visualLine.Indent);
                double alignmentOffset = AlignmentOffset(paragraph.Style.Alignment, rightToLeft, available, visualLine.Width);
                double cursorX = x + visualLine.Indent + alignmentOffset;
                foreach (InlineMathItem item in visualLine.Items) {
                    if (item.Expression != null && item.MathOptions != null) {
                        double inlineRight = Math.Min(_drawing.Width, x + width);
                        if (cursorY + item.Height <= _drawing.Height && cursorX + item.Width <= inlineRight) {
                            OfficeMathRenderer.AddToDrawing(_drawing, item.Expression, cursorX, cursorY, item.MathOptions);
                        } else {
                            AddDiagnostic("ONENOTE_RENDER_MATH_CLIPPED", "A mathematical expression exceeded the page canvas and was rendered as readable text.", item.Expression.ToPlainText());
                            double fallbackWidth = Math.Min(inlineRight - cursorX, _drawing.Width - cursorX);
                            double fallbackHeight = Math.Min(item.Height, _drawing.Height - cursorY);
                            if (fallbackWidth > 0D && fallbackHeight > 0D) {
                                RenderPlainMath(item.Expression, cursorX, cursorY, fallbackWidth, fallbackHeight, item.MathOptions.Font, item.MathOptions.Color);
                            }
                        }
                    } else if (item.Text.Length > 0) {
                        double drawableWidth = Math.Min(item.Width, _drawing.Width - cursorX);
                        double drawableHeight = Math.Min(item.Height, _drawing.Height - cursorY);
                        if (drawableWidth > 0D && drawableHeight > 0D) {
                            _drawing.AddText(item.Text, cursorX, cursorY, drawableWidth, drawableHeight, CreateFont(item.Run),
                                ResolveColor(item.Run.Style.ColorArgb), OfficeTextAlignment.Left, item.Height, wrapText: false);
                        }
                    }
                    cursorX += item.Width;
                }
                cursorY += Math.Max(visualLine.Height + 3D, exactLineHeight);
                if (cursorY >= _drawing.Height) break;
            }
            return Math.Max(DefaultParagraphHeight, cursorY - y - 3D);
        }

        private IReadOnlyList<InlineMathLine> CreateInlineMathLines(OneNoteParagraph paragraph, string prefix, double width) {
            OfficeTextParagraphIndent indent = CreateParagraphIndent(paragraph) ?? OfficeTextParagraphIndent.Empty;
            var items = new List<InlineMathItem>();
            if (prefix.Length > 0) AddInlineTextItems(items, paragraph.Runs[0], prefix, Math.Max(1D, width - indent.MaximumOffset));
            foreach (OneNoteTextRun run in paragraph.Runs) {
                if (run.MathExpression != null && _options.IncludeMath) {
                    OfficeMathRenderOptions mathOptions = CreateMathOptions(run);
                    OfficeMathLayoutMetrics metrics = OfficeMathRenderer.Measure(run.MathExpression, mathOptions);
                    items.Add(InlineMathItem.Math(run, run.MathExpression, mathOptions, metrics));
                } else {
                    string text = run.MathExpression?.ToPlainText() ?? run.Text ?? string.Empty;
                    AddInlineTextItems(items, run, text, Math.Max(1D, width - indent.MaximumOffset));
                }
            }

            var lines = new List<InlineMathLine>();
            InlineMathLine line = new InlineMathLine(indent.FirstLineOffset);
            foreach (InlineMathItem item in items) {
                if (item.IsLineBreak) {
                    lines.Add(line);
                    line = new InlineMathLine(indent.ContinuationLineOffset);
                    continue;
                }
                double available = Math.Max(1D, width - line.Indent);
                if (line.Items.Count > 0 && line.Width + item.Width > available) {
                    lines.Add(line);
                    line = new InlineMathLine(indent.ContinuationLineOffset);
                }
                if (line.Items.Count == 0 && item.IsWhitespace) continue;
                line.Add(item);
            }
            if (line.Items.Count > 0 || lines.Count == 0) lines.Add(line);
            return lines;
        }

        internal static double ParagraphSpaceBefore(OneNoteElement element) =>
            element is OneNoteParagraph paragraph ? ParagraphDistance(paragraph.Style.SpaceBefore) : 0D;

        internal static double ParagraphSpaceAfter(OneNoteElement element) =>
            element is OneNoteParagraph paragraph ? ParagraphDistance(paragraph.Style.SpaceAfter) : 0D;

        internal static double ParagraphDistance(double? value) {
            if (!value.HasValue || double.IsNaN(value.Value) || double.IsInfinity(value.Value) || value.Value <= 0D) return 0D;
            return value.Value * PointsPerHalfInch;
        }

        private static double ResolveParagraphLineHeight(OneNoteParagraph paragraph, double fontSize) =>
            Math.Max(fontSize * 1.25D, ParagraphDistance(paragraph.Style.ExactLineSpacing));

        internal double MeasureElementHeight(OneNoteElement element, double width) {
            if (element.Layout?.Height.HasValue == true) return Math.Max(1D, element.Layout.Height.Value * PointsPerHalfInch);
            if (element is OneNoteParagraph paragraph) return MeasureParagraphHeight(paragraph, width);
            if (element is OneNoteOutline outline) return MeasureElementsBounds(outline.Children, width).Bottom;
            if (element is OneNoteTable table) return MeasureTableHeight(table, width);
            if (element is OneNoteImage image) {
                if (image.HeightHalfInches.HasValue) return Math.Max(1D, image.HeightHalfInches.Value * PointsPerHalfInch);
                return Math.Max(80D, Math.Min(240D, width) * 0.6D);
            }
            if (element is OneNoteInk ink) {
                OfficeInkBounds bounds = ink.Ink.GetBounds();
                if (bounds.IsEmpty) return DefaultParagraphHeight;
                double sourceWidth = Math.Max(0.000001D, (bounds.X + bounds.Width) * PointsPerHalfInch);
                double fit = Math.Min(1D, Math.Max(1D, width) / sourceWidth);
                return Math.Max(DefaultParagraphHeight, (bounds.Y + bounds.Height) * PointsPerHalfInch * fit);
            }
            if (element is OneNoteMath math) {
                OfficeMathLayoutMetrics metrics = OfficeMathRenderer.Measure(math.GetExpression(), _options.Math);
                return Math.Max(DefaultParagraphHeight, metrics.Height);
            }
            if (element is OneNoteBinaryElement) return 34D;
            return 32D;
        }

        internal double MeasureElementsHeight(IEnumerable<OneNoteElement> elements, double width) {
            return MeasureElementsBounds(elements, width).Bottom;
        }

        internal (double Right, double Bottom) MeasureElementsBounds(IEnumerable<OneNoteElement> elements, double width) {
            double right = 0D;
            double bottom = 0D;
            double cursor = 0D;
            double pendingSpace = 0D;
            foreach (OneNoteElement element in elements) {
                bool participatesInFlow = element.Layout?.Y.HasValue != true;
                double elementX = element.Layout?.X.HasValue == true ? element.Layout.X.Value * PointsPerHalfInch : 0D;
                double elementY = element.Layout?.Y.HasValue == true
                    ? element.Layout.Y.Value * PointsPerHalfInch
                    : cursor + Math.Max(pendingSpace, ParagraphSpaceBefore(element));
                double remainingWidth = Math.Max(1D, width - Math.Max(0D, elementX));
                double elementWidth = element.Layout?.Width.HasValue == true
                    ? Math.Max(1D, element.Layout.Width.Value * PointsPerHalfInch)
                    : remainingWidth;
                double elementHeight = MeasureElementHeight(element, elementWidth);
                double extentWidth = elementWidth;
                if (element is OneNoteOutline outline && element.Layout?.Width.HasValue != true) {
                    extentWidth = Math.Max(extentWidth, MeasureElementsBounds(outline.Children, elementWidth).Right);
                }
                right = Math.Max(right, elementX + extentWidth);
                bottom = Math.Max(bottom, elementY + elementHeight);
                if (participatesInFlow) {
                    cursor = Math.Max(cursor, elementY + elementHeight);
                    pendingSpace = element is OneNoteParagraph ? ParagraphSpaceAfter(element) : 6D;
                }
            }
            return (Math.Max(1D, right), Math.Max(DefaultParagraphHeight, Math.Max(bottom, cursor + pendingSpace)));
        }

        private double MeasureParagraphHeight(OneNoteParagraph paragraph, double width) {
            if (paragraph.Runs.Count == 0) return DefaultParagraphHeight;
            string prefix = CreateParagraphPrefix(paragraph);
            if (paragraph.Runs.Any(run => run.MathExpression != null)) {
                IReadOnlyList<InlineMathLine> lines = CreateInlineMathLines(paragraph, prefix, width);
                double exactLineHeight = ParagraphDistance(paragraph.Style.ExactLineSpacing);
                double height = lines.Sum(line => Math.Max(line.Height + 3D, exactLineHeight)) - 3D;
                return Math.Max(DefaultParagraphHeight, height);
            }
            IReadOnlyList<OfficeRichTextRun> runs = CreateParagraphRichTextRuns(paragraph, prefix);
            double fontSize = paragraph.Runs.Max(run => run.Style.FontSize ?? _options.DefaultFont.Size);
            double lineHeight = ResolveParagraphLineHeight(paragraph, fontSize);
            return MeasureRichTextHeight(runs, width, lineHeight, CreateParagraphIndent(paragraph));
        }

        private double MeasureTableHeight(OneNoteTable table, double width) {
            int columns = Math.Max(table.ColumnWidths.Count, table.Rows.Count == 0 ? 0 : table.Rows.Max(row => row.Cells.Count));
            if (columns == 0) return DefaultParagraphHeight;
            double[] columnWidths = ResolveTableColumns(table, columns, width);
            double height = 0D;
            foreach (OneNoteTableRow row in table.Rows) {
                double rowHeight = 32D;
                for (int column = 0; column < row.Cells.Count && column < columnWidths.Length; column++) {
                    rowHeight = Math.Max(rowHeight, MeasureElementsHeight(row.Cells[column].Content, Math.Max(1D, columnWidths[column] - 8D)) + 8D);
                }
                height += rowHeight;
            }
            return Math.Max(DefaultParagraphHeight, height);
        }

        private IReadOnlyList<OfficeRichTextRun> CreateParagraphRichTextRuns(OneNoteParagraph paragraph, string prefix) {
            var runs = new List<OfficeRichTextRun>(paragraph.Runs.Count);
            for (int index = 0; index < paragraph.Runs.Count; index++) {
                OneNoteTextRun run = paragraph.Runs[index];
                runs.Add(CreateRichTextRun(run, (index == 0 ? prefix : string.Empty) + (run.Text ?? string.Empty)));
            }
            return runs;
        }

        private double MeasureRichTextHeight(
            IReadOnlyList<OfficeRichTextRun> runs,
            double width,
            double lineHeight,
            OfficeTextParagraphIndent? indent) {
            double maxFontSize = runs.Count == 0 ? _options.DefaultFont.Size : runs.Max(run => run.FontSize);
            OfficeRichTextBlockLayout layout = OfficeTextLayoutEngine.LayoutRichTextBlock(
                runs,
                Math.Max(1D, width),
                double.MaxValue,
                lineHeight / Math.Max(1D, maxFontSize),
                MeasureRichTextPoints,
                wrap: true,
                shrinkToFit: false,
                paragraphIndent: indent);
            return Math.Max(DefaultParagraphHeight, layout.Height + 4D);
        }

        private double MeasureRichTextPoints(string? text, double fontSize, string? fontFamily) =>
            MeasureTextPoints(text ?? string.Empty, new OfficeFontInfo(fontFamily ?? _options.DefaultFont.FamilyName, fontSize));

        private void AddInlineTextItems(ICollection<InlineMathItem> output, OneNoteTextRun run, string text, double maximumWidth) {
            string normalized = text.Replace("\r\n", "\n").Replace('\r', '\n');
            var token = new System.Text.StringBuilder();
            bool? whitespace = null;
            for (int index = 0; index < normalized.Length; index++) {
                char value = normalized[index];
                if (value == '\n') {
                    FlushInlineToken(output, run, token, maximumWidth);
                    whitespace = null;
                    output.Add(InlineMathItem.LineBreak(run));
                    continue;
                }
                bool currentWhitespace = char.IsWhiteSpace(value);
                if (whitespace.HasValue && whitespace.Value != currentWhitespace) FlushInlineToken(output, run, token, maximumWidth);
                whitespace = currentWhitespace;
                token.Append(value);
            }
            FlushInlineToken(output, run, token, maximumWidth);
        }

        private void FlushInlineToken(ICollection<InlineMathItem> output, OneNoteTextRun run, System.Text.StringBuilder token, double maximumWidth) {
            if (token.Length == 0) return;
            string value = token.ToString();
            token.Clear();
            OfficeFontInfo font = CreateFont(run);
            double height = Math.Max(DefaultParagraphHeight, font.Size * 1.3D);
            int start = 0;
            while (start < value.Length) {
                int length = value.Length - start;
                string part = value.Substring(start, length);
                double measured = MeasureTextPoints(part, font);
                while (length > 1 && measured > maximumWidth) {
                    length = Math.Max(1, length / 2);
                    part = value.Substring(start, length);
                    measured = MeasureTextPoints(part, font);
                }
                while (start + length < value.Length) {
                    string candidate = value.Substring(start, length + 1);
                    double candidateWidth = MeasureTextPoints(candidate, font);
                    if (candidateWidth > maximumWidth) break;
                    length++;
                    part = candidate;
                    measured = candidateWidth;
                }
                output.Add(InlineMathItem.TextRun(run, part, Math.Max(0D, measured), height));
                start += length;
            }
        }

        private static double AlignmentOffset(OneNoteParagraphAlignment? alignment, bool rightToLeft, double available, double used) {
            double remaining = Math.Max(0D, available - used);
            if (alignment == OneNoteParagraphAlignment.Right || (!alignment.HasValue && rightToLeft)) return remaining;
            if (alignment == OneNoteParagraphAlignment.Center) return remaining / 2D;
            return 0D;
        }

        private sealed class InlineMathLine {
            internal InlineMathLine(double indent) { Indent = indent; }
            internal IList<InlineMathItem> Items { get; } = new List<InlineMathItem>();
            internal double Indent { get; }
            internal double Width { get; private set; }
            internal double Height { get; private set; } = DefaultParagraphHeight;
            internal void Add(InlineMathItem item) {
                Items.Add(item);
                Width += item.Width;
                Height = Math.Max(Height, item.Height);
            }
        }

        private sealed class InlineMathItem {
            private InlineMathItem(OneNoteTextRun run, string text, double width, double height) {
                Run = run;
                Text = text;
                Width = width;
                Height = height;
                IsWhitespace = text.Length > 0 && text.All(char.IsWhiteSpace);
            }
            internal OneNoteTextRun Run { get; }
            internal string Text { get; }
            internal double Width { get; }
            internal double Height { get; }
            internal bool IsWhitespace { get; }
            internal bool IsLineBreak { get; private set; }
            internal OfficeMathExpression? Expression { get; private set; }
            internal OfficeMathRenderOptions? MathOptions { get; private set; }
            internal static InlineMathItem TextRun(OneNoteTextRun run, string text, double width, double height) => new InlineMathItem(run, text, width, height);
            internal static InlineMathItem LineBreak(OneNoteTextRun run) => new InlineMathItem(run, string.Empty, 0D, DefaultParagraphHeight) { IsLineBreak = true };
            internal static InlineMathItem Math(OneNoteTextRun run, OfficeMathExpression expression, OfficeMathRenderOptions options, OfficeMathLayoutMetrics metrics) =>
                new InlineMathItem(run, string.Empty, metrics.Width + 2D, System.Math.Max(DefaultParagraphHeight, metrics.Height)) { Expression = expression, MathOptions = options };
        }

        private double RenderTable(OneNoteTable table, double x, double y, double width, bool rightToLeft) {
            int columnCount = Math.Max(table.ColumnWidths.Count, table.Rows.Count == 0 ? 0 : table.Rows.Max(row => row.Cells.Count));
            if (columnCount == 0) return DefaultParagraphHeight;
            double[] columns = ResolveTableColumns(table, columnCount, width);
            double cursorY = y;
            foreach (OneNoteTableRow row in table.Rows) {
                double rowHeight = 32D;
                for (int column = 0; column < row.Cells.Count && column < columns.Length; column++) {
                    rowHeight = Math.Max(rowHeight, MeasureElementsHeight(row.Cells[column].Content, Math.Max(1D, columns[column] - 8D)) + 8D);
                }
                rowHeight = Math.Min(rowHeight, Math.Max(1D, _drawing.Height - cursorY));
                double cursorX = x;
                for (int column = 0; column < columnCount; column++) {
                    double cellWidth = Math.Min(columns[column], Math.Max(1D, _drawing.Width - cursorX));
                    OneNoteTableCell? cell = column < row.Cells.Count ? row.Cells[column] : null;
                    OfficeShape frame = OfficeShape.Rectangle(cellWidth, rowHeight);
                    frame.FillColor = cell?.ShadingColorArgb.HasValue == true ? ResolveColor(cell.ShadingColorArgb) : OfficeColor.White;
                    frame.StrokeColor = table.BordersVisible ? OfficeColor.FromRgb(166, 166, 166) : null;
                    frame.StrokeWidth = table.BordersVisible ? 0.75D : 0D;
                    _drawing.AddShape(frame, cursorX, cursorY);
                    if (cell != null) {
                        double contentY = cursorY + 4D;
                        double pendingCellSpace = 0D;
                        foreach (OneNoteElement content in cell.Content) {
                            double childY = contentY + Math.Max(pendingCellSpace, ParagraphSpaceBefore(content));
                            double used = RenderElement(
                                content,
                                cursorX + 4D,
                                childY,
                                Math.Max(1D, cellWidth - 8D),
                                Math.Max(1D, rowHeight - 8D),
                                inheritedRightToLeft: rightToLeft);
                            contentY = childY + used;
                            pendingCellSpace = content is OneNoteParagraph ? ParagraphSpaceAfter(content) : 2D;
                            if (contentY >= cursorY + rowHeight) break;
                        }
                    }
                    cursorX += columns[column];
                }
                cursorY += rowHeight;
                if (cursorY >= _drawing.Height) break;
            }
            return Math.Max(DefaultParagraphHeight, cursorY - y);
        }

        private double RenderImage(OneNoteImage image, double x, double y, double width, double height, bool forcePageBounds) {
            if (!_options.IncludeImages) return 0D;
            if (image.Payload == null) {
                AddDiagnostic("ONENOTE_RENDER_IMAGE_PAYLOAD_MISSING", "A OneNote image could not be rendered because its payload is unresolved.", image.FileName);
                return RenderImagePlaceholder(image, x, y, width);
            }
            byte[] bytes;
            try {
                bytes = image.Payload.ToArray(_options.MaxImageBytes);
            } catch (Exception exception) when (exception is IOException || exception is InvalidOperationException) {
                AddDiagnostic("ONENOTE_RENDER_IMAGE_PAYLOAD_FAILED", exception.Message, image.FileName);
                return RenderImagePlaceholder(image, x, y, width);
            }
            OfficeImageInfo? info = OfficeImageReader.TryIdentifyByContent(bytes, image.FileName, out OfficeImageInfo identified) ? identified : null;
            double renderWidth = forcePageBounds ? _drawing.Width : ResolveImageWidth(image, info, width);
            double renderHeight = forcePageBounds ? _drawing.Height : ResolveImageHeight(image, info, renderWidth);
            if (height > 0D) renderHeight = Math.Min(renderHeight, height);
            renderWidth = Math.Max(1D, Math.Min(renderWidth, _drawing.Width - x));
            renderHeight = Math.Max(1D, Math.Min(renderHeight, _drawing.Height - y));
            if (renderWidth <= 0D || renderHeight <= 0D) return 0D;
            try {
                var projection = new OfficeImageProjection(new OfficeImagePlacement(x, y, renderWidth, renderHeight));
                _drawing.AddImage(bytes, image.MediaType ?? info?.MimeType, projection, image.AltText);
            } catch (Exception exception) when (exception is ArgumentException || exception is NotSupportedException || exception is InvalidOperationException) {
                AddDiagnostic("ONENOTE_RENDER_IMAGE_UNSUPPORTED", exception.Message, image.FileName);
                return RenderImagePlaceholder(image, x, y, renderWidth);
            }
            return renderHeight;
        }

        private double RenderInk(OneNoteInk ink, double x, double y, double width, double height) {
            if (!_options.IncludeInk || ink.Strokes.Count == 0) return 0D;
            OfficeInkBounds sourceBounds = ink.Ink.GetBounds();
            double fit = 1D;
            if (!sourceBounds.IsEmpty) {
                double sourceRight = Math.Max(0.000001D, (sourceBounds.X + sourceBounds.Width) * PointsPerHalfInch);
                fit = Math.Min(1D, width / sourceRight);
                if (height > 0D) {
                    double sourceBottom = Math.Max(0.000001D, (sourceBounds.Y + sourceBounds.Height) * PointsPerHalfInch);
                    fit = Math.Min(fit, height / sourceBottom);
                }
            }
            var scaled = new OfficeInkDocument();
            foreach (OfficeInkStroke stroke in ink.Strokes) {
                OfficeInkStroke clone = stroke.Clone();
                clone.Transform = (clone.Transform ?? OfficeTransform.Identity).Then(OfficeTransform.Scale(PointsPerHalfInch * fit, PointsPerHalfInch * fit));
                scaled.Add(clone);
            }
            OfficeInkBounds bounds = scaled.GetBounds();
            OfficeInkRenderer.AddToDrawing(_drawing, scaled, x, y, _options.Ink);
            double rendered = bounds.IsEmpty ? DefaultParagraphHeight : Math.Max(DefaultParagraphHeight, bounds.Y + bounds.Height);
            if (height > 0D) rendered = Math.Min(rendered, height);
            return rendered;
        }

        private double RenderMath(OneNoteMath math, double x, double y, double availableWidth, double availableHeight) {
            OfficeMathExpression expression = math.GetExpression();
            if (!_options.IncludeMath) return RenderPlainMath(expression, x, y, availableWidth, availableHeight);
            OfficeMathLayoutMetrics metrics = OfficeMathRenderer.Measure(expression, _options.Math);
            if (x + metrics.Width > _drawing.Width || y + metrics.Height > _drawing.Height ||
                metrics.Width > availableWidth ||
                (availableHeight > 0D && metrics.Height > availableHeight)) {
                AddDiagnostic("ONENOTE_RENDER_MATH_CLIPPED", "A mathematical expression exceeded the page canvas and was rendered as readable text.", math.Text);
                return RenderPlainMath(expression, x, y, availableWidth, availableHeight);
            }
            OfficeMathRenderer.AddToDrawing(_drawing, expression, x, y, _options.Math);
            return Math.Max(DefaultParagraphHeight, metrics.Height);
        }

        private double RenderPlainMath(
            OfficeMathExpression expression,
            double x,
            double y,
            double availableWidth,
            double availableHeight,
            OfficeFontInfo? font = null,
            OfficeColor? color = null) {
            if (x >= _drawing.Width || y >= _drawing.Height || availableWidth <= 0D) return 0D;
            double height = Math.Min(24D, Math.Max(1D, _drawing.Height - y));
            if (availableHeight > 0D) height = Math.Min(height, availableHeight);
            double width = Math.Max(1D, Math.Min(availableWidth, _drawing.Width - x));
            _drawing.AddText(expression.ToPlainText(), x, y, width, height, font ?? _options.Math.Font, color ?? _options.Math.Color, wrapText: false);
            return height;
        }

        private double RenderAttachment(OneNoteBinaryElement binary, double x, double y, double width) {
            if (!_options.IncludeAttachmentPlaceholders) return 0D;
            double height = Math.Min(34D, Math.Max(1D, _drawing.Height - y));
            double frameWidth = Math.Min(Math.Max(120D, MeasureTextPoints(binary.FileName ?? binary.Kind.ToString(), _options.DefaultFont) + 34D), width);
            OfficeShape frame = OfficeShape.RoundedRectangle(frameWidth, height, 5D);
            frame.FillColor = OfficeColor.FromRgb(245, 247, 250);
            frame.StrokeColor = OfficeColor.FromRgb(150, 158, 170);
            frame.StrokeWidth = 0.8D;
            _drawing.AddShape(frame, x, y);
            string label = binary is OneNoteMedia ? "▶ " : "📎 ";
            _drawing.AddText(label + (binary.FileName ?? binary.Kind.ToString()), x + 8D, y + 6D, Math.Max(1D, frameWidth - 16D), Math.Max(1D, height - 12D),
                _options.DefaultFont, OfficeColor.FromRgb(45, 55, 72), wrapText: false);
            return height;
        }

        private double RenderImagePlaceholder(OneNoteImage image, double x, double y, double width) {
            if (!_options.IncludeAttachmentPlaceholders) return 0D;
            double height = Math.Min(80D, Math.Max(1D, _drawing.Height - y));
            double renderWidth = Math.Min(Math.Max(80D, width), Math.Max(1D, _drawing.Width - x));
            OfficeShape frame = OfficeShape.Rectangle(renderWidth, height);
            frame.FillColor = OfficeColor.FromRgb(247, 247, 247);
            frame.StrokeColor = OfficeColor.FromRgb(176, 176, 176);
            frame.StrokeDashStyle = OfficeStrokeDashStyle.Dash;
            _drawing.AddShape(frame, x, y);
            _drawing.AddText(image.AltText ?? image.FileName ?? "Image", x + 6D, y + 6D, Math.Max(1D, renderWidth - 12D), Math.Max(1D, height - 12D),
                _options.DefaultFont, OfficeColor.DimGray, OfficeTextAlignment.Center, verticalAlignment: OfficeTextVerticalAlignment.Center, wrapText: true);
            return height;
        }

        private static double[] ResolveTableColumns(OneNoteTable table, int count, double width) {
            var result = new double[count];
            if (table.ColumnWidths.Count == count && table.ColumnWidths.All(value => value > 0D)) {
                double sourceSum = table.ColumnWidths.Sum();
                double scaledSum = sourceSum * PointsPerHalfInch;
                double fit = scaledSum > width ? width / scaledSum : 1D;
                for (int index = 0; index < count; index++) result[index] = table.ColumnWidths[index] * PointsPerHalfInch * fit;
                return result;
            }
            for (int index = 0; index < count; index++) result[index] = width / count;
            return result;
        }

        private double ResolveImageWidth(OneNoteImage image, OfficeImageInfo? info, double fallback) {
            if (image.Layout?.Width.HasValue == true) return image.Layout.Width.Value * PointsPerHalfInch;
            if (image.WidthHalfInches.HasValue) return image.WidthHalfInches.Value * PointsPerHalfInch;
            if (info != null && info.Width > 0) return info.Width / info.DpiX * 72D;
            return Math.Min(240D, fallback);
        }

        private static double ResolveImageHeight(OneNoteImage image, OfficeImageInfo? info, double width) {
            if (image.Layout?.Height.HasValue == true) return image.Layout.Height.Value * PointsPerHalfInch;
            if (image.HeightHalfInches.HasValue) return image.HeightHalfInches.Value * PointsPerHalfInch;
            if (info?.AspectRatio.HasValue == true) return width / info.AspectRatio.Value;
            return Math.Max(80D, width * 0.6D);
        }

        private OfficeMathRenderOptions CreateMathOptions(OneNoteTextRun run) {
            OfficeMathRenderOptions options = _options.Math.Clone();
            options.Padding = 0D;
            options.Font = CreateFont(run);
            options.Color = ResolveColor(run.Style.ColorArgb);
            return options;
        }

        private OfficeRichTextRun CreateRichTextRun(OneNoteTextRun run, string text) => new OfficeRichTextRun(
            text,
            run.Style.FontSize ?? _options.DefaultFont.Size,
            ResolveColor(run.Style.ColorArgb),
            run.Style.Bold == true,
            run.Style.Italic == true,
            run.Style.Underline == true || !string.IsNullOrWhiteSpace(run.Hyperlink),
            run.Style.FontFamily ?? _options.DefaultFont.FamilyName,
            run.Style.Strikethrough == true,
            run.Style.HighlightColorArgb.HasValue ? ResolveColor(run.Style.HighlightColorArgb) : (OfficeColor?)null);

        private OfficeFontInfo CreateFont(OneNoteTextRun run) {
            OfficeFontStyle style = OfficeFontStyle.Regular;
            if (run.Style.Bold == true) style |= OfficeFontStyle.Bold;
            if (run.Style.Italic == true) style |= OfficeFontStyle.Italic;
            if (run.Style.Underline == true || !string.IsNullOrWhiteSpace(run.Hyperlink)) style |= OfficeFontStyle.Underline;
            if (run.Style.Strikethrough == true) style |= OfficeFontStyle.Strikethrough;
            return new OfficeFontInfo(run.Style.FontFamily ?? _options.DefaultFont.FamilyName, run.Style.FontSize ?? _options.DefaultFont.Size, style);
        }

        private double MeasureTextPoints(string text, OfficeFontInfo font) =>
            _measurer.MeasureWidth(text, _measurer.CreateStyle(font)) * 72D / OfficeTextMeasurer.DefaultDpi;

        private static OfficeTextAlignment MapAlignment(OneNoteParagraphAlignment? alignment, bool rightToLeft) {
            switch (alignment) {
                case OneNoteParagraphAlignment.Center: return OfficeTextAlignment.Center;
                case OneNoteParagraphAlignment.Right: return OfficeTextAlignment.Right;
                case OneNoteParagraphAlignment.Justify: return OfficeTextAlignment.Justify;
                default: return rightToLeft ? OfficeTextAlignment.Right : OfficeTextAlignment.Left;
            }
        }

        private static OfficeTextParagraphIndent? CreateParagraphIndent(OneNoteParagraph paragraph) {
            if (paragraph.List == null || paragraph.List.Level <= 0) return null;
            return new OfficeTextParagraphIndent(paragraph.List.Level * 18D, paragraph.List.Level * 18D);
        }

        private static string CreateParagraphPrefix(OneNoteParagraph paragraph) {
            var builder = new System.Text.StringBuilder();
            if (paragraph.List != null) {
                builder.Append(paragraph.List.Ordered ? Math.Max(1, paragraph.List.DisplayIndex ?? 1).ToString() + ". " : "• ");
            }
            foreach (OneNoteTag tag in paragraph.Tags) {
                if (tag.IsCheckable || tag.IsTask) builder.Append(tag.IsCompleted ? "☑ " : "☐ ");
                else if (!string.IsNullOrWhiteSpace(tag.Label)) builder.Append('[').Append(tag.Label).Append("] ");
            }
            return builder.ToString();
        }

        private static OfficeColor ResolveColor(uint? argb) {
            if (!argb.HasValue) return OfficeColor.Black;
            uint value = argb.Value;
            byte alpha = (byte)(value >> 24);
            if (alpha == 0) alpha = 255;
            return OfficeColor.FromRgba((byte)(value >> 16), (byte)(value >> 8), (byte)value, alpha);
        }

        private void AddDiagnostic(string code, string message, string? source) =>
            _diagnostics.Add(new OfficeImageExportDiagnostic(OfficeImageExportDiagnosticSeverity.Warning, code, message, source));
    }
}

using System.Collections.Generic;
using System.Globalization;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using OfficeIMO.Drawing;
using W = DocumentFormat.OpenXml.Wordprocessing;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Word.Pdf {
    public static partial class WordPdfConverterExtensions {
        private static bool TryRenderNativeCoverPageCanvas(INativePdfFlow pdf, WordDocument document, W.SdtBlock? sdtBlock, PdfSaveOptions? options) {
            W.SdtContentBlock? content = sdtBlock?.SdtContentBlock;
            WordSection? section = document.Sections.FirstOrDefault();
            if (content == null || section == null || !HasNativeVmlCoverDrawing(content)) {
                return false;
            }

            PdfCore.PageSize pageSize = GetNativePageSize(section, options);
            var rootFrame = new NativeVmlFrame(0D, 0D, pageSize.Width, pageSize.Height, pageSize.Width, pageSize.Height, 0D, 0D);
            bool rendered = false;
            pdf.Canvas(canvas => rendered = RenderNativeVmlCoverChildren(canvas, document, content.ChildElements, rootFrame, pageSize.Width, pageSize.Height) || rendered);
            return rendered;
        }

        private static bool HasNativeVmlCoverDrawing(OpenXmlElement element) {
            foreach (OpenXmlElement descendant in element.Descendants()) {
                string localName = descendant.LocalName;
                if ((localName == "group" || localName == "shape" || localName == "rect" || localName == "line") &&
                    descendant.NamespaceUri == "urn:schemas-microsoft-com:vml" &&
                    IsNativeVmlPositionedCoverElement(descendant)) {
                    return true;
                }
            }

            return false;
        }

        private static bool IsNativeVmlPositionedCoverElement(OpenXmlElement element) {
            Dictionary<string, string> style = ParseNativeVmlStyle(GetNativeOpenXmlAttribute(element, "style"));
            return style.TryGetValue("position", out string? position) && position.Equals("absolute", StringComparison.OrdinalIgnoreCase) ||
                   style.ContainsKey("mso-left-percent") ||
                   style.ContainsKey("mso-top-percent") ||
                   style.ContainsKey("margin-left") ||
                   style.ContainsKey("margin-top") ||
                   style.ContainsKey("left") ||
                   style.ContainsKey("top");
        }

        private static bool RenderNativeVmlCoverChildren(PdfCore.PdfPageCanvas canvas, WordDocument document, IEnumerable<OpenXmlElement> children, NativeVmlFrame frame, double pageWidth, double pageHeight) {
            bool rendered = false;
            foreach (OpenXmlElement child in children) {
                if (child.NamespaceUri == "urn:schemas-microsoft-com:vml") {
                    if (child.LocalName == "group") {
                        rendered = RenderNativeVmlGroup(canvas, document, child, frame, pageWidth, pageHeight) || rendered;
                    } else if (child.LocalName == "shape" || child.LocalName == "rect" || child.LocalName == "line") {
                        rendered = RenderNativeVmlShape(canvas, document, child, frame, pageWidth, pageHeight) || rendered;
                    }

                    continue;
                }

                rendered = RenderNativeVmlCoverChildren(canvas, document, child.ChildElements, frame, pageWidth, pageHeight) || rendered;
            }

            return rendered;
        }

        private static bool RenderNativeVmlGroup(PdfCore.PdfPageCanvas canvas, WordDocument document, OpenXmlElement group, NativeVmlFrame parentFrame, double pageWidth, double pageHeight) {
            if (!TryGetNativeVmlBox(group, parentFrame, pageWidth, pageHeight, out NativeVmlBox box)) {
                return RenderNativeVmlCoverChildren(canvas, document, group.ChildElements, parentFrame, pageWidth, pageHeight);
            }

            (double coordWidth, double coordHeight) = GetNativeVmlCoordSize(group, box.Width, box.Height);
            (double coordOriginX, double coordOriginY) = GetNativeVmlCoordOrigin(group);
            var frame = new NativeVmlFrame(box.X, box.Y, box.Width, box.Height, coordWidth, coordHeight, coordOriginX, coordOriginY);
            return RenderNativeVmlCoverChildren(canvas, document, group.ChildElements, frame, pageWidth, pageHeight);
        }

        private static bool RenderNativeVmlShape(PdfCore.PdfPageCanvas canvas, WordDocument document, OpenXmlElement element, NativeVmlFrame frame, double pageWidth, double pageHeight) {
            bool isLine = element.LocalName.Equals("line", StringComparison.OrdinalIgnoreCase);
            if (IsNativeVmlHidden(element) ||
                !TryGetNativeVmlBox(element, frame, pageWidth, pageHeight, out NativeVmlBox box) ||
                (!isLine && (box.Width <= 0D || box.Height <= 0D)) ||
                (isLine && box.Width <= 0D && box.Height <= 0D)) {
                return false;
            }

            bool rendered = false;
            if (TryCreateNativeVmlShape(element, box.Width, box.Height, out OfficeShape? shape) && shape != null) {
                if (IsNativeVmlBoxVisibleOnPage(box, pageWidth, pageHeight)) {
                    RenderNativeVmlVisible(canvas, box, pageWidth, pageHeight, target => target.Shape(shape, box.X, box.Y, new PdfCore.PdfDrawingStyle { Decorative = true }));
                    rendered = true;
                }
            }

            IReadOnlyList<PdfCore.TextRun> textRuns = GetNativeVmlTextRuns(document, element);
            if (textRuns.Count > 0) {
                PdfCore.PdfCanvasTextBoxStyle style = CreateNativeVmlTextBoxStyle(element, textRuns);
                if (IsNativeVmlBoxVisibleOnPage(box, pageWidth, pageHeight)) {
                    RenderNativeVmlVisible(canvas, box, pageWidth, pageHeight, target => target.TextBox(textRuns, box.X, box.Y, box.Width, box.Height, style));
                    rendered = true;
                }
            }

            return rendered;
        }

        private static bool TryCreateNativeVmlShape(OpenXmlElement element, double width, double height, out OfficeShape? shape) {
            shape = null;
            string localName = element.LocalName;
            if (localName == "line") {
                (double x1, double y1) = ParseNativeShapePoint(GetNativeOpenXmlAttribute(element, "from") ?? "0pt,0pt");
                (double x2, double y2) = ParseNativeShapePoint(GetNativeOpenXmlAttribute(element, "to") ?? (width.ToString(CultureInfo.InvariantCulture) + "pt,0pt"));
                double minX = Math.Min(x1, x2);
                double minY = Math.Min(y1, y2);
                shape = OfficeShape.Line(x1 - minX, y1 - minY, x2 - minX, y2 - minY);
            } else if (localName == "rect") {
                shape = OfficeShape.Rectangle(width, height);
            } else if (IsNativeVmlTextBoxShape(element) && string.IsNullOrWhiteSpace(GetNativeOpenXmlAttribute(element, "fillcolor")) && GetNativeVmlChild(element, "fill") == null) {
                return false;
            } else if (TryCreateNativeVmlPathShape(element, width, height, out shape)) {
            } else if (IsNativeVmlPentagonShape(element)) {
                shape = OfficeShape.Polygon(
                    new OfficePoint(0D, 0D),
                    new OfficePoint(width * 0.88D, 0D),
                    new OfficePoint(width, height / 2D),
                    new OfficePoint(width * 0.88D, height),
                    new OfficePoint(0D, height));
            } else {
                shape = OfficeShape.Rectangle(width, height);
            }

            if (shape == null) {
                return false;
            }

            ApplyNativeVmlShapeStyle(shape, element);
            bool hasFill = shape.Kind != OfficeShapeKind.Line && shape.FillColor.HasValue;
            bool hasStroke = shape.StrokeColor.HasValue && shape.StrokeWidth > 0D;
            if (!hasFill && !hasStroke) {
                shape = null;
                return false;
            }

            return true;
        }

        private static bool TryCreateNativeVmlPathShape(OpenXmlElement element, double width, double height, out OfficeShape? shape) {
            shape = null;
            string? path = GetNativeOpenXmlAttribute(element, "path");
            if (string.IsNullOrWhiteSpace(path)) {
                return false;
            }

            (double coordWidth, double coordHeight) = GetNativeVmlCoordSize(element, width, height);
            if (TryCreateNativeVmlCommandPath(path!, coordWidth, coordHeight, width, height, out shape)) {
                return true;
            }

            MatchCollection matches = Regex.Matches(path!, @"-?\d+(?:\.\d+)?");
            if (matches.Count < 6) {
                return false;
            }

            var points = new List<OfficePoint>();
            for (int i = 0; i + 1 < matches.Count; i += 2) {
                if (!double.TryParse(matches[i].Value, NumberStyles.Float, CultureInfo.InvariantCulture, out double x) ||
                    !double.TryParse(matches[i + 1].Value, NumberStyles.Float, CultureInfo.InvariantCulture, out double y)) {
                    return false;
                }

                points.Add(new OfficePoint(x / coordWidth * width, y / coordHeight * height));
            }

            if (points.Count < 3) {
                return false;
            }

            try {
                shape = OfficeShape.Polygon(points);
                return true;
            } catch (ArgumentException) {
                shape = null;
                return false;
            }
        }

        private static bool TryCreateNativeVmlCommandPath(string path, double coordWidth, double coordHeight, double width, double height, out OfficeShape? shape) {
            shape = null;
            var commands = new List<OfficePathCommand>();
            double currentX = 0D;
            double currentY = 0D;
            int index = 0;
            bool moved = false;

            while (index < path.Length) {
                char command = char.ToLowerInvariant(path[index]);
                if (!IsNativeVmlPathCommand(command)) {
                    index++;
                    continue;
                }

                index++;
                int start = index;
                while (index < path.Length && !IsNativeVmlPathCommand(char.ToLowerInvariant(path[index]))) {
                    index++;
                }

                string segment = path.Substring(start, index - start);
                if (command == 'e') {
                    break;
                }

                if (command == 'x') {
                    commands.Add(OfficePathCommand.Close());
                    continue;
                }

                List<(double X, double Y)> pairs = ParseNativeVmlPathPairs(segment);
                if (command == 'm' && pairs.Count == 0) {
                    pairs.Add((0D, 0D));
                }

                foreach ((double x, double y) in pairs) {
                    if (command == 'm') {
                        currentX = x;
                        currentY = y;
                        commands.Add(OfficePathCommand.MoveTo(
                            ScaleNativeVmlPathX(currentX, coordWidth, width),
                            ScaleNativeVmlPathY(currentY, coordHeight, height)));
                        moved = true;
                    } else if (command == 'l') {
                        if (!moved) {
                            commands.Add(OfficePathCommand.MoveTo(0D, 0D));
                            moved = true;
                        }

                        currentX = x;
                        currentY = y;
                        commands.Add(OfficePathCommand.LineTo(
                            ScaleNativeVmlPathX(currentX, coordWidth, width),
                            ScaleNativeVmlPathY(currentY, coordHeight, height)));
                    } else if (command == 'r') {
                        if (!moved) {
                            commands.Add(OfficePathCommand.MoveTo(0D, 0D));
                            moved = true;
                        }

                        currentX += x;
                        currentY += y;
                        commands.Add(OfficePathCommand.LineTo(
                            ScaleNativeVmlPathX(currentX, coordWidth, width),
                            ScaleNativeVmlPathY(currentY, coordHeight, height)));
                    }
                }
            }

            int drawingCommands = commands.Count(command => command.Kind != OfficePathCommandKind.Close);
            if (drawingCommands < 3) {
                return false;
            }

            try {
                shape = OfficeShape.Path(commands);
                return true;
            } catch (ArgumentException) {
                shape = null;
                return false;
            }
        }

        private static bool IsNativeVmlPathCommand(char value) =>
            value is 'm' or 'l' or 'r' or 'x' or 'e';

        private static List<(double X, double Y)> ParseNativeVmlPathPairs(string segment) {
            var pairs = new List<(double X, double Y)>();
            MatchCollection matches = Regex.Matches(segment, @"-?\d+(?:\.\d+)?");
            for (int i = 0; i < matches.Count; i += 2) {
                double x = ParseNativeVmlPathPart(matches[i].Value);
                double y = i + 1 < matches.Count ? ParseNativeVmlPathPart(matches[i + 1].Value) : 0D;
                pairs.Add((x, y));
            }

            return pairs;
        }

        private static double ParseNativeVmlPathPart(string value) =>
            ParseNativeVmlDouble(value) ?? 0D;

        private static double ScaleNativeVmlPathX(double value, double coordWidth, double width) =>
            coordWidth > 0D ? value / coordWidth * width : value;

        private static double ScaleNativeVmlPathY(double value, double coordHeight, double height) =>
            coordHeight > 0D ? value / coordHeight * height : value;

        private static void ApplyNativeVmlShapeStyle(OfficeShape shape, OpenXmlElement element) {
            if (shape.Kind != OfficeShapeKind.Line) {
                OpenXmlElement? fillElement = GetNativeVmlChild(element, "fill");
                string? childFillColor = fillElement is not null ? GetNativeOpenXmlAttribute(fillElement, "color") : null;
                PdfCore.PdfColor? fill = ParseNativeColor(NormalizeNativeVmlColor(GetNativeOpenXmlAttribute(element, "fillcolor") ?? childFillColor));
                if (fill.HasValue) {
                    shape.FillColor = fill.Value.ToOfficeColor();
                }

                double? fillOpacity = fillElement is not null ? ParseNativeVmlOpacity(GetNativeOpenXmlAttribute(fillElement, "opacity")) : null;
                if (fillOpacity.HasValue) {
                    shape.FillOpacity = fillOpacity.Value;
                }
            }

            string? strokeColor = NormalizeNativeVmlColor(GetNativeOpenXmlAttribute(element, "strokecolor"));
            bool stroked = !string.Equals(GetNativeOpenXmlAttribute(element, "stroked"), "f", StringComparison.OrdinalIgnoreCase) &&
                           !string.Equals(GetNativeOpenXmlAttribute(element, "stroked"), "false", StringComparison.OrdinalIgnoreCase);
            if (shape.Kind != OfficeShapeKind.Line && string.IsNullOrWhiteSpace(strokeColor)) {
                shape.StrokeColor = null;
                shape.StrokeWidth = 0D;
                return;
            }

            if (!stroked) {
                shape.StrokeColor = null;
                shape.StrokeWidth = 0D;
                return;
            }

            shape.StrokeColor = (ParseNativeColor(strokeColor) ?? PdfCore.PdfColor.Black).ToOfficeColor();
            shape.StrokeWidth = ParseNativeVmlStrokeWeight(GetNativeOpenXmlAttribute(element, "strokeweight")) ?? 1D;
        }

        private static OpenXmlElement? GetNativeVmlChild(OpenXmlElement? element, string localName) {
            return element?.ChildElements.FirstOrDefault(child =>
                child.NamespaceUri == "urn:schemas-microsoft-com:vml" &&
                child.LocalName.Equals(localName, StringComparison.OrdinalIgnoreCase));
        }

        private static IReadOnlyList<PdfCore.TextRun> GetNativeVmlTextRuns(WordDocument document, OpenXmlElement element) {
            var runs = new List<PdfCore.TextRun>();
            bool hasParagraph = false;
            foreach (W.Paragraph paragraph in element.Descendants<W.Paragraph>()) {
                if (hasParagraph) {
                    runs.Add(PdfCore.TextRun.LineBreak());
                }

                hasParagraph = true;
                foreach (W.Run run in paragraph.Descendants<W.Run>()) {
                    W.RunProperties? properties = run.RunProperties;
                    foreach (W.Text text in run.Descendants<W.Text>()) {
                        string value = ResolveNativeVmlRunText(document, run, text);
                        if (string.IsNullOrEmpty(value)) {
                            continue;
                        }

                        runs.Add(new PdfCore.TextRun(
                            value,
                            bold: HasNativeOnOff(properties?.Bold),
                            underline: properties?.Underline != null,
                            color: ParseNativeColor(properties?.Color?.Val?.Value),
                            italic: HasNativeOnOff(properties?.Italic),
                            strike: HasNativeOnOff(properties?.Strike),
                            fontSize: GetNativeVmlRunFontSize(properties)));
                    }
                }
            }

            while (runs.Count > 0 && runs[0].Text == "\n") {
                runs.RemoveAt(0);
            }

            while (runs.Count > 0 && runs[runs.Count - 1].Text == "\n") {
                runs.RemoveAt(runs.Count - 1);
            }

            return runs;
        }

        private static string ResolveNativeVmlRunText(WordDocument document, W.Run run, W.Text textElement) {
            string text = textElement.Text;
            string value = ResolveNativeBuiltInPropertyPlaceholders(document, text);
            string? propertyValue = GetNativeBuiltInPropertyValue(document, GetNativeVmlRunSdtProperties(textElement, run));
            if (!string.IsNullOrWhiteSpace(propertyValue) &&
                (string.Equals(value, text, StringComparison.Ordinal) || IsNativeVmlPlaceholderText(value))) {
                value = PreserveNativeVmlPlaceholderSpacing(text, propertyValue!);
            }

            W.RunProperties? properties = run.RunProperties;
            if (HasNativeOnOff(properties?.Caps) || HasNativeOnOff(properties?.SmallCaps)) {
                value = value.ToUpperInvariant();
            }

            return value;
        }

        private static W.SdtProperties? GetNativeVmlRunSdtProperties(W.Text textElement, W.Run run) {
            W.SdtRun? runSdt = textElement.Ancestors<W.SdtRun>().FirstOrDefault() ?? run.Ancestors<W.SdtRun>().FirstOrDefault();
            if (runSdt?.SdtProperties != null) {
                return runSdt.SdtProperties;
            }

            W.SdtBlock? blockSdt = textElement.Ancestors<W.SdtBlock>().FirstOrDefault() ?? run.Ancestors<W.SdtBlock>().FirstOrDefault();
            return blockSdt?.SdtProperties;
        }

        private static bool IsNativeVmlPlaceholderText(string text) {
            string trimmed = text.Trim();
            return trimmed.Length > 2 && trimmed[0] == '[' && trimmed[trimmed.Length - 1] == ']';
        }

        private static string PreserveNativeVmlPlaceholderSpacing(string sourceText, string value) {
            int leading = 0;
            while (leading < sourceText.Length && char.IsWhiteSpace(sourceText[leading])) {
                leading++;
            }

            int trailing = 0;
            while (trailing < sourceText.Length - leading && char.IsWhiteSpace(sourceText[sourceText.Length - trailing - 1])) {
                trailing++;
            }

            return sourceText.Substring(0, leading) + value + sourceText.Substring(sourceText.Length - trailing, trailing);
        }

        private static PdfCore.PdfCanvasTextBoxStyle CreateNativeVmlTextBoxStyle(OpenXmlElement element, IReadOnlyList<PdfCore.TextRun> runs) {
            var style = new PdfCore.PdfCanvasTextBoxStyle {
                Background = null,
                BorderColor = null,
                BorderWidth = 0D,
                TextColor = null,
                Align = GetNativeVmlTextAlign(element),
                VerticalAlign = GetNativeVmlVerticalAlign(element),
                FontSize = GetNativeVmlDefaultFontSize(runs),
                LineHeight = GetNativeVmlDefaultFontSize(runs) * 1.2D
            };

            ApplyNativeVmlTextBoxInset(style, element);
            return style;
        }

        private static bool TryGetNativeVmlBox(OpenXmlElement element, NativeVmlFrame frame, double pageWidth, double pageHeight, out NativeVmlBox box) {
            Dictionary<string, string> style = ParseNativeVmlStyle(GetNativeOpenXmlAttribute(element, "style"));
            double width = ResolveNativeVmlLength(style.TryGetValue("width", out string? widthValue) ? widthValue : null, frame.Width, frame.CoordWidth) ??
                           ResolveNativeVmlPercent(style, "mso-width-percent", pageWidth) ?? 0D;
            double height = ResolveNativeVmlLength(style.TryGetValue("height", out string? heightValue) ? heightValue : null, frame.Height, frame.CoordHeight) ??
                            ResolveNativeVmlPercent(style, "mso-height-percent", pageHeight) ?? 0D;

            double? xPercent = ResolveNativeVmlPercent(style, "mso-left-percent", pageWidth);
            double? yPercent = ResolveNativeVmlPercent(style, "mso-top-percent", pageHeight);
            double x = xPercent ??
                       ResolveNativeVmlPosition(style.TryGetValue("left", out string? left) ? left : null, frame.Width, frame.CoordWidth, frame.CoordOriginX) ??
                       ResolveNativeVmlPosition(style.TryGetValue("margin-left", out string? marginLeft) ? marginLeft : null, frame.Width, frame.CoordWidth, frame.CoordOriginX) ?? 0D;
            double y = yPercent ??
                       ResolveNativeVmlPosition(style.TryGetValue("top", out string? top) ? top : null, frame.Height, frame.CoordHeight, frame.CoordOriginY) ??
                       ResolveNativeVmlPosition(style.TryGetValue("margin-top", out string? marginTop) ? marginTop : null, frame.Height, frame.CoordHeight, frame.CoordOriginY) ?? 0D;

            if (element.LocalName.Equals("line", StringComparison.OrdinalIgnoreCase) &&
                (width <= 0D || height <= 0D) &&
                TryGetNativeVmlLineBounds(element, out double lineX, out double lineY, out double lineWidth, out double lineHeight)) {
                x += lineX;
                y += lineY;
                width = Math.Max(width, lineWidth);
                height = Math.Max(height, lineHeight);
            }

            if (style.TryGetValue("mso-position-vertical", out string? verticalPosition) &&
                verticalPosition.Equals("center", StringComparison.OrdinalIgnoreCase) &&
                !yPercent.HasValue &&
                !style.ContainsKey("top")) {
                y = (pageHeight - height) / 2D;
            }

            box = new NativeVmlBox(frame.X + x, frame.Y + y, width, height);
            return width > 0D && height > 0D;
        }

        private static bool TryGetNativeVmlLineBounds(OpenXmlElement element, out double x, out double y, out double width, out double height) {
            (double x1, double y1) = ParseNativeShapePoint(GetNativeOpenXmlAttribute(element, "from") ?? "0pt,0pt");
            (double x2, double y2) = ParseNativeShapePoint(GetNativeOpenXmlAttribute(element, "to") ?? "0pt,0pt");
            x = Math.Min(x1, x2);
            y = Math.Min(y1, y2);
            width = Math.Abs(x2 - x1);
            height = Math.Abs(y2 - y1);
            if (width <= 0D && height <= 0D) {
                return false;
            }

            width = Math.Max(width, 0.01D);
            height = Math.Max(height, 0.01D);
            return true;
        }

        private static Dictionary<string, string> ParseNativeVmlStyle(string? style) {
            var values = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            if (string.IsNullOrWhiteSpace(style)) {
                return values;
            }

            foreach (string part in style!.Split(';')) {
                int separator = part.IndexOf(':');
                if (separator <= 0 || separator == part.Length - 1) {
                    continue;
                }

                values[part.Substring(0, separator).Trim()] = part.Substring(separator + 1).Trim();
            }

            return values;
        }

        private static double? ResolveNativeVmlPosition(string? value, double parentSize, double parentCoord, double parentCoordOrigin) {
            if (string.IsNullOrWhiteSpace(value)) {
                return null;
            }

            string normalized = value!.Trim();
            if (HasNativeVmlLengthUnit(normalized) || normalized.EndsWith("%", StringComparison.OrdinalIgnoreCase)) {
                return ResolveNativeVmlLength(normalized, parentSize, parentCoord);
            }

            double? number = ParseNativeVmlDouble(normalized);
            if (!number.HasValue) {
                return null;
            }

            double relative = number.Value - parentCoordOrigin;
            return parentCoord > 0D && Math.Abs(parentCoord - parentSize) > 0.01D
                ? relative / parentCoord * parentSize
                : relative;
        }

        private static double? ResolveNativeVmlLength(string? value, double parentSize, double parentCoord) {
            if (string.IsNullOrWhiteSpace(value)) {
                return null;
            }

            string normalized = value!.Trim();
            if (normalized.EndsWith("pt", StringComparison.OrdinalIgnoreCase)) return ParseNativeVmlDouble(normalized.Substring(0, normalized.Length - 2));
            if (normalized.EndsWith("in", StringComparison.OrdinalIgnoreCase)) return ParseNativeVmlDouble(normalized.Substring(0, normalized.Length - 2)) * 72D;
            if (normalized.EndsWith("cm", StringComparison.OrdinalIgnoreCase)) return ParseNativeVmlDouble(normalized.Substring(0, normalized.Length - 2)) * 28.3464566929D;
            if (normalized.EndsWith("mm", StringComparison.OrdinalIgnoreCase)) return ParseNativeVmlDouble(normalized.Substring(0, normalized.Length - 2)) * 2.83464566929D;
            if (normalized.EndsWith("px", StringComparison.OrdinalIgnoreCase)) return ParseNativeVmlDouble(normalized.Substring(0, normalized.Length - 2)) * 0.75D;
            if (normalized.EndsWith("%", StringComparison.OrdinalIgnoreCase)) return ParseNativeVmlDouble(normalized.Substring(0, normalized.Length - 1)) * parentSize / 100D;

            double? number = ParseNativeVmlDouble(normalized);
            if (!number.HasValue) {
                return null;
            }

            return parentCoord > 0D && Math.Abs(parentCoord - parentSize) > 0.01D
                ? number.Value / parentCoord * parentSize
                : number.Value;
        }

        private static bool HasNativeVmlLengthUnit(string value) =>
            value.EndsWith("pt", StringComparison.OrdinalIgnoreCase) ||
            value.EndsWith("in", StringComparison.OrdinalIgnoreCase) ||
            value.EndsWith("cm", StringComparison.OrdinalIgnoreCase) ||
            value.EndsWith("mm", StringComparison.OrdinalIgnoreCase) ||
            value.EndsWith("px", StringComparison.OrdinalIgnoreCase);

        private static double? ResolveNativeVmlPercent(Dictionary<string, string> style, string key, double reference) {
            if (!style.TryGetValue(key, out string? value) ||
                !double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out double percent) ||
                percent <= 0D) {
                return null;
            }

            return reference * percent / 1000D;
        }

        private static (double Width, double Height) GetNativeVmlCoordSize(OpenXmlElement element, double fallbackWidth, double fallbackHeight) {
            string? coordSize = GetNativeOpenXmlAttribute(element, "coordsize");
            if (string.IsNullOrWhiteSpace(coordSize)) {
                return (fallbackWidth, fallbackHeight);
            }

            string[] parts = coordSize!.Split(',');
            if (parts.Length == 2 &&
                double.TryParse(parts[0], NumberStyles.Float, CultureInfo.InvariantCulture, out double width) &&
                double.TryParse(parts[1], NumberStyles.Float, CultureInfo.InvariantCulture, out double height) &&
                width > 0D &&
                height > 0D) {
                return (width, height);
            }

            return (fallbackWidth, fallbackHeight);
        }

        private static (double X, double Y) GetNativeVmlCoordOrigin(OpenXmlElement element) {
            string? coordOrigin = GetNativeOpenXmlAttribute(element, "coordorigin");
            if (string.IsNullOrWhiteSpace(coordOrigin)) {
                return (0D, 0D);
            }

            string[] parts = coordOrigin!.Split(',');
            if (parts.Length == 2 &&
                double.TryParse(parts[0], NumberStyles.Float, CultureInfo.InvariantCulture, out double x) &&
                double.TryParse(parts[1], NumberStyles.Float, CultureInfo.InvariantCulture, out double y)) {
                return (x, y);
            }

            return (0D, 0D);
        }

        private static string? NormalizeNativeVmlColor(string? value) {
            if (string.IsNullOrWhiteSpace(value)) {
                return null;
            }

            string trimmed = value!.Trim();
            int space = trimmed.IndexOf(' ');
            if (space > 0) {
                trimmed = trimmed.Substring(0, space);
            }

            return trimmed.Equals("none", StringComparison.OrdinalIgnoreCase) ? null : trimmed;
        }

        private static string? GetNativeOpenXmlAttribute(OpenXmlElement element, string localName) {
            foreach (OpenXmlAttribute attribute in element.GetAttributes()) {
                if (attribute.LocalName.Equals(localName, StringComparison.OrdinalIgnoreCase)) {
                    return attribute.Value;
                }
            }

            return null;
        }

        private static bool IsNativeVmlHidden(OpenXmlElement element) {
            Dictionary<string, string> style = ParseNativeVmlStyle(GetNativeOpenXmlAttribute(element, "style"));
            return style.TryGetValue("visibility", out string? value) &&
                   value.Equals("hidden", StringComparison.OrdinalIgnoreCase);
        }

        private static bool IsNativeVmlTextBoxShape(OpenXmlElement element) =>
            string.Equals(GetNativeOpenXmlAttribute(element, "type"), "#_x0000_t202", StringComparison.OrdinalIgnoreCase);

        private static bool IsNativeVmlPentagonShape(OpenXmlElement element) =>
            string.Equals(GetNativeOpenXmlAttribute(element, "type"), "#_x0000_t15", StringComparison.OrdinalIgnoreCase);

        private static double? ParseNativeVmlStrokeWeight(string? value) =>
            ResolveNativeVmlLength(value, 1D, 1D);

        private static double? ParseNativeVmlDouble(string value) =>
            double.TryParse(value.Trim(), NumberStyles.Float, CultureInfo.InvariantCulture, out double result)
                ? result
                : double.TryParse(value.Trim().Replace(',', '.'), NumberStyles.Float, CultureInfo.InvariantCulture, out result)
                    ? result
                    : null;

        private static double? ParseNativeVmlOpacity(string? value) {
            if (string.IsNullOrWhiteSpace(value)) {
                return null;
            }

            string normalized = value!.Trim();
            double? parsed;
            if (normalized.EndsWith("%", StringComparison.OrdinalIgnoreCase)) {
                parsed = ParseNativeVmlDouble(normalized.Substring(0, normalized.Length - 1));
                if (parsed.HasValue) {
                    parsed /= 100D;
                }
            } else {
                parsed = ParseNativeVmlDouble(normalized);
            }

            if (!parsed.HasValue) {
                return null;
            }

            return Math.Max(0D, Math.Min(1D, parsed.Value));
        }

        private static bool IsNativeVmlBoxInsidePage(NativeVmlBox box, double pageWidth, double pageHeight) =>
            box.X >= 0D &&
            box.Y >= 0D &&
            box.Width > 0D &&
            box.Height > 0D &&
            box.X + box.Width <= pageWidth + 0.01D &&
            box.Y + box.Height <= pageHeight + 0.01D;

        private static bool IsNativeVmlBoxVisibleOnPage(NativeVmlBox box, double pageWidth, double pageHeight) =>
            box.Width > 0D &&
            box.Height > 0D &&
            box.X < pageWidth &&
            box.Y < pageHeight &&
            box.X + box.Width > 0D &&
            box.Y + box.Height > 0D;

        private static void RenderNativeVmlVisible(PdfCore.PdfPageCanvas canvas, NativeVmlBox box, double pageWidth, double pageHeight, Action<PdfCore.PdfPageCanvas> render) {
            if (IsNativeVmlBoxInsidePage(box, pageWidth, pageHeight)) {
                render(canvas);
                return;
            }

            canvas.Clip(0D, 0D, pageWidth, pageHeight, render);
        }

        private static bool HasNativeOnOff(OpenXmlElement? element) {
            if (element == null) {
                return false;
            }

            string? value = GetNativeOpenXmlAttribute(element, "val");
            return value == null ||
                   value.Equals("1", StringComparison.OrdinalIgnoreCase) ||
                   value.Equals("true", StringComparison.OrdinalIgnoreCase);
        }

        private static double? GetNativeVmlRunFontSize(W.RunProperties? properties) {
            string? value = properties?.FontSize?.Val?.Value;
            return double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out double halfPoints) && halfPoints > 0D
                ? halfPoints / 2D
                : null;
        }

        private static double GetNativeVmlDefaultFontSize(IReadOnlyList<PdfCore.TextRun> runs) {
            double max = 0D;
            foreach (PdfCore.TextRun run in runs) {
                if (run.FontSize.HasValue && run.FontSize.Value > max) {
                    max = run.FontSize.Value;
                }
            }

            return max > 0D ? max : 11D;
        }

        private static PdfCore.PdfAlign GetNativeVmlTextAlign(OpenXmlElement element) {
            W.Justification? justification = element.Descendants<W.Justification>().FirstOrDefault();
            W.JustificationValues? value = justification?.Val?.Value;
            if (value == W.JustificationValues.Center) return PdfCore.PdfAlign.Center;
            if (value == W.JustificationValues.Right) return PdfCore.PdfAlign.Right;
            return PdfCore.PdfAlign.Left;
        }

        private static PdfCore.PdfVerticalAlign GetNativeVmlVerticalAlign(OpenXmlElement element) {
            Dictionary<string, string> style = ParseNativeVmlStyle(GetNativeOpenXmlAttribute(element, "style"));
            if (style.TryGetValue("v-text-anchor", out string? anchor)) {
                if (anchor.Equals("middle", StringComparison.OrdinalIgnoreCase)) return PdfCore.PdfVerticalAlign.Middle;
                if (anchor.Equals("bottom", StringComparison.OrdinalIgnoreCase)) return PdfCore.PdfVerticalAlign.Bottom;
            }

            return PdfCore.PdfVerticalAlign.Top;
        }

        private static void ApplyNativeVmlTextBoxInset(PdfCore.PdfCanvasTextBoxStyle style, OpenXmlElement element) {
            OpenXmlElement? textBox = element.Descendants().FirstOrDefault(child => child.NamespaceUri == "urn:schemas-microsoft-com:vml" && child.LocalName == "textbox");
            string? inset = textBox == null ? null : GetNativeOpenXmlAttribute(textBox, "inset");
            if (string.IsNullOrWhiteSpace(inset)) {
                style.PaddingX = 0D;
                style.PaddingY = 0D;
                return;
            }

            string[] parts = inset!.Split(',');
            if (parts.Length < 2) {
                return;
            }

            style.PaddingLeft = ResolveNativeVmlLength(parts[0], 0D, 0D) ?? 0D;
            style.PaddingTop = ResolveNativeVmlLength(parts[1], 0D, 0D) ?? 0D;
            style.PaddingRight = parts.Length > 2 ? ResolveNativeVmlLength(parts[2], 0D, 0D) ?? style.PaddingLeft : style.PaddingLeft;
            style.PaddingBottom = parts.Length > 3 ? ResolveNativeVmlLength(parts[3], 0D, 0D) ?? style.PaddingTop : style.PaddingTop;
        }

        private readonly struct NativeVmlFrame {
            public NativeVmlFrame(double x, double y, double width, double height, double coordWidth, double coordHeight, double coordOriginX, double coordOriginY) {
                X = x;
                Y = y;
                Width = width;
                Height = height;
                CoordWidth = coordWidth;
                CoordHeight = coordHeight;
                CoordOriginX = coordOriginX;
                CoordOriginY = coordOriginY;
            }

            public double X { get; }
            public double Y { get; }
            public double Width { get; }
            public double Height { get; }
            public double CoordWidth { get; }
            public double CoordHeight { get; }
            public double CoordOriginX { get; }
            public double CoordOriginY { get; }
        }

        private readonly struct NativeVmlBox {
            public NativeVmlBox(double x, double y, double width, double height) {
                X = x;
                Y = y;
                Width = width;
                Height = height;
            }

            public double X { get; }
            public double Y { get; }
            public double Width { get; }
            public double Height { get; }
        }
    }
}

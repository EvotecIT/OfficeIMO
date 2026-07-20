using System.Collections.Generic;
using System.Globalization;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Drawing;
using W = DocumentFormat.OpenXml.Wordprocessing;
using V = DocumentFormat.OpenXml.Vml;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Word.Pdf {
    public static partial class WordPdfConverterExtensions {
        internal const long NativeVmlImageMaxBytes = 32L * 1024L * 1024L;
        private const double MaxNativeVmlLengthPoints = 1_000_000D;
        private const double MaxNativeVmlTextPathFontSizePoints = 400D;

        private static bool TryRenderNativeCoverPageCanvas(INativePdfFlow pdf, WordDocument document, W.SdtBlock? sdtBlock, WordSection section, PdfSaveOptions? options) {
            W.SdtContentBlock? content = sdtBlock?.SdtContentBlock;
            if (content == null || !HasNativeVmlCoverDrawing(content)) {
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
                if ((localName == "group" || localName == "shape" || localName == "rect" || localName == "line" || localName == "oval" || localName == "roundrect") &&
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
            foreach (OpenXmlElement child in OrderNativeVmlCoverChildren(children)) {
                if (child.NamespaceUri == "urn:schemas-microsoft-com:vml") {
                    if (child.LocalName == "group") {
                        rendered = RenderNativeVmlGroup(canvas, document, child, frame, pageWidth, pageHeight) || rendered;
                    } else if (child.LocalName == "shape" || child.LocalName == "rect" || child.LocalName == "line" || child.LocalName == "oval" || child.LocalName == "roundrect") {
                        rendered = RenderNativeVmlShape(canvas, document, child, frame, pageWidth, pageHeight) || rendered;
                    }

                    continue;
                }

                rendered = RenderNativeVmlCoverChildren(canvas, document, child.ChildElements, frame, pageWidth, pageHeight) || rendered;
            }

            return rendered;
        }

        private static IEnumerable<OpenXmlElement> OrderNativeVmlCoverChildren(IEnumerable<OpenXmlElement> children) {
            List<OpenXmlElement> childList = children.SelectMany(EnumerateNativeVmlCoverRenderElements).ToList();
            if (childList.Count <= 1) {
                return childList;
            }

            return childList
                .Select((Element, Index) => new {
                    Element,
                    Index,
                    ZIndex = GetNativeVmlEffectiveZIndex(Element)
                })
                .OrderBy(item => item.ZIndex ?? 0D)
                .ThenBy(item => item.Index)
                .Select(item => item.Element);
        }

        private static IEnumerable<OpenXmlElement> EnumerateNativeVmlCoverRenderElements(OpenXmlElement element) {
            if (element.NamespaceUri == "urn:schemas-microsoft-com:vml") {
                if (element.LocalName == "group" ||
                    element.LocalName == "shape" ||
                    element.LocalName == "rect" ||
                    element.LocalName == "line" ||
                    element.LocalName == "oval" ||
                    element.LocalName == "roundrect") {
                    yield return element;
                }

                yield break;
            }

            foreach (OpenXmlElement child in element.ChildElements) {
                foreach (OpenXmlElement descendant in EnumerateNativeVmlCoverRenderElements(child)) {
                    yield return descendant;
                }
            }
        }

        private static double? GetNativeVmlEffectiveZIndex(OpenXmlElement element) {
            if (element.NamespaceUri == "urn:schemas-microsoft-com:vml") {
                return GetNativeVmlZIndex(element);
            }

            double? zIndex = null;
            foreach (OpenXmlElement descendant in element.Descendants()) {
                if (descendant.NamespaceUri != "urn:schemas-microsoft-com:vml" ||
                    !IsNativeVmlPositionedCoverElement(descendant)) {
                    continue;
                }

                double? descendantZIndex = GetNativeVmlZIndex(descendant);
                if (!descendantZIndex.HasValue) {
                    continue;
                }

                zIndex = zIndex.HasValue ? Math.Min(zIndex.Value, descendantZIndex.Value) : descendantZIndex.Value;
            }

            return zIndex;
        }

        private static double? GetNativeVmlZIndex(OpenXmlElement element) {
            if (element.NamespaceUri != "urn:schemas-microsoft-com:vml") {
                return null;
            }

            Dictionary<string, string> style = ParseNativeVmlStyle(GetNativeOpenXmlAttribute(element, "style"));
            return style.TryGetValue("z-index", out string? value) ? ParseNativeVmlDouble(value) : null;
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

            bool renderedImage = TryRenderNativeVmlImage(canvas, document, element, box, pageWidth, pageHeight);
            bool rendered = renderedImage;
            if ((ShouldRenderNativeVmlShapeFallback(element) || renderedImage) &&
                TryCreateNativeVmlShape(element, box.Width, box.Height, frame, out OfficeShape? shape) &&
                shape != null) {
                if (renderedImage) {
                    shape.FillColor = null;
                    shape.FillGradient = null;
                }

                bool hasVisibleFrame = !renderedImage || shape.StrokeColor.HasValue || shape.Shadow != null || shape.Kind == OfficeShapeKind.Line;
                if (hasVisibleFrame && IsNativeVmlBoxVisibleOnPage(box, pageWidth, pageHeight)) {
                    RenderNativeVmlVisible(canvas, box, pageWidth, pageHeight, target => target.Shape(shape, box.X, box.Y, new PdfCore.PdfDrawingStyle { Decorative = true }));
                    rendered = true;
                }
            }

            IReadOnlyList<PdfCore.TextRun> textRuns = GetNativeVmlTextRuns(document, element);
            if (textRuns.Count > 0) {
                PdfCore.PdfCanvasTextBoxStyle style = CreateNativeVmlTextBoxStyle(element, textRuns);
                double textRotation = GetNativeVmlRotationDegrees(element) ?? 0D;
                if (IsNativeVmlBoxVisibleOnPage(box, pageWidth, pageHeight)) {
                    RenderNativeVmlVisible(canvas, box, pageWidth, pageHeight, target => target.TextBox(textRuns, box.X, box.Y, box.Width, box.Height, style, textRotation));
                    rendered = true;
                }
            }

            return rendered;
        }

        private static bool ShouldRenderNativeVmlShapeFallback(OpenXmlElement element) => element.GetFirstChild<V.ImageData>() == null;

        private static bool TryRenderNativeVmlImage(PdfCore.PdfPageCanvas canvas, WordDocument document, OpenXmlElement element, NativeVmlBox box, double pageWidth, double pageHeight) {
            V.ImageData? imageData = element.GetFirstChild<V.ImageData>();
            if (imageData == null ||
                !IsNativeVmlBoxVisibleOnPage(box, pageWidth, pageHeight) ||
                !TryGetNativeVmlImageBytes(document, element, imageData, out byte[]? imageBytes) ||
                imageBytes == null ||
                !TryPrepareNativePdfImageBytes(imageBytes, out byte[] preparedBytes, out _)) {
                return false;
            }

            double rotation = GetNativeVmlRotationDegrees(element) ?? 0D;
            bool horizontalFlip = false;
            bool verticalFlip = false;
            string? flip = GetNativeVmlFlip(element);
            if (!string.IsNullOrWhiteSpace(flip)) {
                horizontalFlip = flip!.IndexOf("x", StringComparison.OrdinalIgnoreCase) >= 0;
                verticalFlip = flip.IndexOf("y", StringComparison.OrdinalIgnoreCase) >= 0;
            }

            RenderNativeVmlVisible(canvas, box, pageWidth, pageHeight, target => target.Image(
                preparedBytes,
                box.X,
                box.Y,
                box.Width,
                box.Height,
                new PdfCore.PdfImageStyle {
                    Fit = OfficeImageFit.Stretch
                },
                rotationAngle: rotation,
                horizontalFlip: horizontalFlip,
                verticalFlip: verticalFlip));
            return true;
        }

        private static bool TryGetNativeVmlImageBytes(WordDocument document, OpenXmlElement element, V.ImageData imageData, out byte[]? imageBytes) {
            imageBytes = null;
            string? relationshipId = imageData.RelationshipId?.Value ?? GetNativeOpenXmlAttribute(imageData, "id");
            if (string.IsNullOrWhiteSpace(relationshipId)) {
                return false;
            }

            if (!TryGetNativeVmlImagePart(document, element, relationshipId!, out ImagePart? imagePart) || imagePart == null) {
                return false;
            }

            using Stream stream = imagePart.GetStream(FileMode.Open, FileAccess.Read);
            return TryReadNativeVmlImageBytes(stream, NativeVmlImageMaxBytes, out imageBytes);
        }

        internal static bool TryReadNativeVmlImageBytes(Stream stream, long maxBytes, out byte[]? imageBytes) {
            imageBytes = null;
            if (stream == null || maxBytes <= 0L) {
                return false;
            }

            if (stream.CanSeek && stream.Length > maxBytes) {
                return false;
            }

            using var memory = new MemoryStream(stream.CanSeek && stream.Length > 0L && stream.Length <= int.MaxValue ? (int)stream.Length : 0);
            byte[] buffer = new byte[81920];
            long totalRead = 0L;
            while (true) {
                int bytesRead = stream.Read(buffer, 0, buffer.Length);
                if (bytesRead == 0) {
                    break;
                }

                totalRead += bytesRead;
                if (totalRead > maxBytes) {
                    return false;
                }

                memory.Write(buffer, 0, bytesRead);
            }

            if (memory.Length == 0L) {
                return false;
            }

            imageBytes = memory.ToArray();
            return true;
        }

        private static bool TryGetNativeVmlImagePart(WordDocument document, OpenXmlElement element, string relationshipId, out ImagePart? imagePart) {
            imagePart = null;
            MainDocumentPart? mainPart = document._wordprocessingDocument?.MainDocumentPart;
            if (TryGetNativeVmlImagePart(mainPart, relationshipId, out imagePart)) {
                return true;
            }

            if (mainPart == null) {
                return false;
            }

            foreach (HeaderPart headerPart in mainPart.HeaderParts) {
                if (TryGetNativeVmlImagePart(headerPart, relationshipId, out imagePart)) {
                    return true;
                }
            }

            foreach (FooterPart footerPart in mainPart.FooterParts) {
                if (TryGetNativeVmlImagePart(footerPart, relationshipId, out imagePart)) {
                    return true;
                }
            }

            OpenXmlPartRootElement? root = element.Ancestors<OpenXmlPartRootElement>().FirstOrDefault();
            if (TryGetNativeVmlImagePart(root?.OpenXmlPart, relationshipId, out imagePart)) {
                return true;
            }

            return false;
        }

        private static bool TryGetNativeVmlImagePart(OpenXmlPartContainer? container, string relationshipId, out ImagePart? imagePart) {
            imagePart = null;
            if (container == null) {
                return false;
            }

            try {
                imagePart = container.GetPartById(relationshipId) as ImagePart;
                return imagePart != null;
            } catch (ArgumentOutOfRangeException) {
                return false;
            }
        }

        private static bool TryCreateNativeVmlShape(OpenXmlElement element, double width, double height, NativeVmlFrame frame, out OfficeShape? shape) {
            shape = null;
            string localName = element.LocalName;
            if (localName == "line") {
                (double x1, double y1) = ParseNativeShapePoint(GetNativeOpenXmlAttribute(element, "from") ?? "0pt,0pt", frame);
                (double x2, double y2) = ParseNativeShapePoint(GetNativeOpenXmlAttribute(element, "to") ?? (width.ToString(CultureInfo.InvariantCulture) + "pt,0pt"), frame);
                double minX = Math.Min(x1, x2);
                double minY = Math.Min(y1, y2);
                shape = OfficeShape.Line(x1 - minX, y1 - minY, x2 - minX, y2 - minY);
            } else if (localName == "rect") {
                shape = OfficeShape.Rectangle(width, height);
            } else if (localName == "oval") {
                shape = OfficeShape.Ellipse(width, height);
            } else if (localName == "roundrect") {
                shape = OfficeShape.RoundedRectangle(width, height, GetNativeVmlRoundRectCornerRadius(element, width, height));
            } else if (IsNativeVmlTextBoxShape(element) && string.IsNullOrWhiteSpace(GetNativeOpenXmlAttribute(element, "fillcolor")) && GetNativeVmlChild(element, "fill") == null) {
                return false;
            } else if (TryCreateNativeVmlPathShape(element, width, height, out shape)) {
            } else if (TryCreateNativeVmlBuiltInShapeType(element, width, height, out shape)) {
            } else {
                shape = OfficeShape.Rectangle(width, height);
            }

            if (shape == null) {
                return false;
            }

            ApplyNativeVmlShapeStyle(shape, element);
            ApplyNativeVmlShapeTransform(shape, element);
            bool hasFill = shape.Kind != OfficeShapeKind.Line && (shape.FillColor.HasValue || shape.FillGradient != null);
            bool hasStroke = shape.StrokeColor.HasValue && shape.StrokeWidth > 0D;
            if (!hasFill && !hasStroke) {
                shape = null;
                return false;
            }

            return true;
        }

        private static bool TryCreateNativeVmlPathShape(OpenXmlElement element, double width, double height, out OfficeShape? shape) {
            shape = null;
            OpenXmlElement? shapeType = GetNativeVmlReferencedShapeTypeElement(element);
            string? path = GetNativeOpenXmlAttribute(element, "path");
            if (string.IsNullOrWhiteSpace(path) && shapeType != null) {
                path = GetNativeOpenXmlAttribute(shapeType, "path") ?? GetNativeOpenXmlAttribute(shapeType, "edgepath");
            }

            if (string.IsNullOrWhiteSpace(path)) {
                return false;
            }

            (double coordWidth, double coordHeight) = GetNativeVmlCoordSize(element, shapeType, width, height);
            IReadOnlyDictionary<string, double> formulaValues = GetNativeVmlFormulaValues(element, shapeType, coordWidth, coordHeight);
            if (TryCreateNativeVmlCommandPath(path!, coordWidth, coordHeight, width, height, formulaValues, out shape)) {
                return true;
            }

            string resolvedPath = ResolveNativeVmlFormulaReferences(path!, formulaValues);
            MatchCollection matches = Regex.Matches(resolvedPath, NativeVmlNumberPattern);
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

        private static bool TryCreateNativeVmlBuiltInShapeType(OpenXmlElement element, double width, double height, out OfficeShape? shape) {
            shape = null;
            string? type = GetNativeOpenXmlAttribute(element, "type");
            if (!string.Equals(type, "#_x0000_t15", StringComparison.OrdinalIgnoreCase)) {
                return false;
            }

            (double coordWidth, _) = GetNativeVmlCoordSize(element, 21600D, 21600D);
            double adjustment = GetNativeVmlFirstAdjustment(element, 16200D);
            double adjustedX = coordWidth > 0D
                ? Math.Max(0D, Math.Min(coordWidth, adjustment)) / coordWidth * width
                : width * 0.75D;

            shape = OfficeShape.Polygon(
                new OfficePoint(adjustedX, 0D),
                new OfficePoint(0D, 0D),
                new OfficePoint(0D, height),
                new OfficePoint(adjustedX, height),
                new OfficePoint(width, height / 2D));
            return true;
        }

        private static bool TryCreateNativeVmlCommandPath(string path, double coordWidth, double coordHeight, double width, double height, IReadOnlyDictionary<string, double> formulaValues, out OfficeShape? shape) {
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
                index = FindNativeVmlPathSegmentEnd(path, index, formulaValues);

                string segment = path.Substring(start, index - start);
                if (command == 'e') {
                    break;
                }

                if (command == 'x') {
                    commands.Add(OfficePathCommand.Close());
                    continue;
                }

                List<(double X, double Y)> pairs = ParseNativeVmlPathPairs(segment, formulaValues);
                if (command == 'm' && pairs.Count == 0) {
                    pairs.Add((0D, 0D));
                }

                if (command is 'c' or 'v' or 'q') {
                    if (!moved) {
                        commands.Add(OfficePathCommand.MoveTo(0D, 0D));
                        moved = true;
                    }

                    if (command == 'c') {
                        for (int i = 0; i + 2 < pairs.Count; i += 3) {
                            (double control1X, double control1Y) = pairs[i];
                            (double control2X, double control2Y) = pairs[i + 1];
                            (double endX, double endY) = pairs[i + 2];
                            commands.Add(OfficePathCommand.CubicBezierTo(
                                new OfficePoint(ScaleNativeVmlPathX(control1X, coordWidth, width), ScaleNativeVmlPathY(control1Y, coordHeight, height)),
                                new OfficePoint(ScaleNativeVmlPathX(control2X, coordWidth, width), ScaleNativeVmlPathY(control2Y, coordHeight, height)),
                                new OfficePoint(ScaleNativeVmlPathX(endX, coordWidth, width), ScaleNativeVmlPathY(endY, coordHeight, height))));
                            currentX = endX;
                            currentY = endY;
                        }
                    } else if (command == 'v') {
                        for (int i = 0; i + 2 < pairs.Count; i += 3) {
                            (double control1OffsetX, double control1OffsetY) = pairs[i];
                            (double control2OffsetX, double control2OffsetY) = pairs[i + 1];
                            (double endOffsetX, double endOffsetY) = pairs[i + 2];
                            double control1X = currentX + control1OffsetX;
                            double control1Y = currentY + control1OffsetY;
                            double control2X = currentX + control2OffsetX;
                            double control2Y = currentY + control2OffsetY;
                            double endX = currentX + endOffsetX;
                            double endY = currentY + endOffsetY;
                            commands.Add(OfficePathCommand.CubicBezierTo(
                                new OfficePoint(ScaleNativeVmlPathX(control1X, coordWidth, width), ScaleNativeVmlPathY(control1Y, coordHeight, height)),
                                new OfficePoint(ScaleNativeVmlPathX(control2X, coordWidth, width), ScaleNativeVmlPathY(control2Y, coordHeight, height)),
                                new OfficePoint(ScaleNativeVmlPathX(endX, coordWidth, width), ScaleNativeVmlPathY(endY, coordHeight, height))));
                            currentX = endX;
                            currentY = endY;
                        }
                    } else {
                        for (int i = 0; i + 1 < pairs.Count; i += 2) {
                            (double controlX, double controlY) = pairs[i];
                            (double endX, double endY) = pairs[i + 1];
                            double control1X = currentX + (controlX - currentX) * 2D / 3D;
                            double control1Y = currentY + (controlY - currentY) * 2D / 3D;
                            double control2X = endX + (controlX - endX) * 2D / 3D;
                            double control2Y = endY + (controlY - endY) * 2D / 3D;

                            commands.Add(OfficePathCommand.CubicBezierTo(
                                new OfficePoint(ScaleNativeVmlPathX(control1X, coordWidth, width), ScaleNativeVmlPathY(control1Y, coordHeight, height)),
                                new OfficePoint(ScaleNativeVmlPathX(control2X, coordWidth, width), ScaleNativeVmlPathY(control2Y, coordHeight, height)),
                                new OfficePoint(ScaleNativeVmlPathX(endX, coordWidth, width), ScaleNativeVmlPathY(endY, coordHeight, height))));
                            currentX = endX;
                            currentY = endY;
                        }
                    }

                    continue;
                }

                foreach ((double x, double y) in pairs) {
                    if (command == 'm') {
                        currentX = x;
                        currentY = y;
                        commands.Add(OfficePathCommand.MoveTo(
                            ScaleNativeVmlPathX(currentX, coordWidth, width),
                            ScaleNativeVmlPathY(currentY, coordHeight, height)));
                        moved = true;
                    } else if (command == 't') {
                        currentX += x;
                        currentY += y;
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
            if (drawingCommands < 2) {
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
            value is 'm' or 't' or 'l' or 'r' or 'c' or 'v' or 'q' or 'x' or 'e';

        private static List<(double X, double Y)> ParseNativeVmlPathPairs(string segment, IReadOnlyDictionary<string, double> formulaValues) {
            var pairs = new List<(double X, double Y)>();
            IReadOnlyList<string> parts = TokenizeNativeVmlPathParts(segment, formulaValues);
            for (int i = 0; i < parts.Count; i += 2) {
                double x = ParseNativeVmlPathPart(parts[i], formulaValues);
                double y = i + 1 < parts.Count ? ParseNativeVmlPathPart(parts[i + 1], formulaValues) : 0D;
                pairs.Add((x, y));
            }

            return pairs;
        }

        private static double ParseNativeVmlPathPart(string value, IReadOnlyDictionary<string, double> formulaValues) =>
            formulaValues.TryGetValue(value, out double resolved)
                ? resolved
                : ParseNativeVmlDouble(value) ?? 0D;

        private static double ScaleNativeVmlPathX(double value, double coordWidth, double width) =>
            coordWidth > 0D ? value / coordWidth * width : value;

        private static double ScaleNativeVmlPathY(double value, double coordHeight, double height) =>
            coordHeight > 0D ? value / coordHeight * height : value;

        private static void ApplyNativeVmlShapeStyle(OfficeShape shape, OpenXmlElement element) {
            if (shape.Kind != OfficeShapeKind.Line) {
                OpenXmlElement? fillElement = GetNativeVmlChild(element, "fill");
                bool fillEnabled = IsNativeVmlSwitchEnabled(GetNativeOpenXmlAttribute(element, "filled")) &&
                                   (fillElement == null || IsNativeVmlSwitchEnabled(GetNativeOpenXmlAttribute(fillElement, "on")));
                if (fillEnabled) {
                    string? childFillColor = fillElement is not null ? GetNativeOpenXmlAttribute(fillElement, "color") : null;
                    string? fillColor = GetNativeOpenXmlAttribute(element, "fillcolor") ?? childFillColor;
                    bool explicitNoFill = IsNativeVmlNoColor(fillColor);
                    PdfCore.PdfColor? fill = explicitNoFill ? null : ParseNativeColor(NormalizeNativeVmlColor(fillColor));
                    if (!explicitNoFill) {
                        shape.FillColor = (fill ?? PdfCore.PdfColor.White).ToOfficeColor();
                    }

                    if (!explicitNoFill && TryGetNativeVmlGradientFill(fillElement, fill, out OfficeLinearGradient? gradient)) {
                        shape.FillGradient = gradient;
                    }

                    double? fillOpacity = fillElement is not null ? ParseNativeVmlOpacity(GetNativeOpenXmlAttribute(fillElement, "opacity")) : null;
                    if (fillOpacity.HasValue) {
                        shape.FillOpacity = fillOpacity.Value;
                    }
                }
            }

            ApplyNativeVmlShapeShadow(shape, element);
            OpenXmlElement? strokeElement = GetNativeVmlChild(element, "stroke");
            string? childStrokeColor = strokeElement is not null ? GetNativeOpenXmlAttribute(strokeElement, "color") : null;
            string? rawStrokeColor = GetNativeOpenXmlAttribute(element, "strokecolor") ?? childStrokeColor;
            string? strokeColor = NormalizeNativeVmlColor(rawStrokeColor);
            bool stroked = IsNativeVmlSwitchEnabled(GetNativeOpenXmlAttribute(element, "stroked")) &&
                           (strokeElement == null || IsNativeVmlSwitchEnabled(GetNativeOpenXmlAttribute(strokeElement, "on")));

            if (!stroked || IsNativeVmlNoColor(rawStrokeColor)) {
                shape.StrokeColor = null;
                shape.StrokeWidth = 0D;
                return;
            }

            shape.StrokeColor = (ParseNativeColor(strokeColor) ?? PdfCore.PdfColor.Black).ToOfficeColor();
            shape.StrokeWidth = ParseNativeVmlStrokeWeight(GetNativeOpenXmlAttribute(element, "strokeweight")) ?? 1D;
            shape.StrokeDashStyle = MapNativeVmlStrokeDashStyle(GetNativeOpenXmlAttribute(strokeElement ?? element, "dashstyle"));
            shape.StrokeLineCap = MapNativeVmlStrokeLineCap(GetNativeOpenXmlAttribute(strokeElement ?? element, "endcap"));
            shape.StrokeLineJoin = MapNativeVmlStrokeLineJoin(GetNativeOpenXmlAttribute(strokeElement ?? element, "joinstyle"));
            double? strokeOpacity = strokeElement is not null ? ParseNativeVmlOpacity(GetNativeOpenXmlAttribute(strokeElement, "opacity")) : null;
            if (strokeOpacity.HasValue) {
                shape.StrokeOpacity = strokeOpacity.Value;
            }
        }

        private static OfficeStrokeDashStyle MapNativeVmlStrokeDashStyle(string? value) {
            if (string.IsNullOrWhiteSpace(value)) {
                return OfficeStrokeDashStyle.Solid;
            }

            string normalized = value!.Replace(" ", string.Empty).Replace("-", string.Empty);
            bool hasDash = normalized.IndexOf("dash", StringComparison.OrdinalIgnoreCase) >= 0;
            bool hasDot = normalized.IndexOf("dot", StringComparison.OrdinalIgnoreCase) >= 0;
            if (hasDash && hasDot) {
                return OfficeStrokeDashStyle.DashDot;
            }

            if (hasDot) {
                return OfficeStrokeDashStyle.Dot;
            }

            return hasDash ? OfficeStrokeDashStyle.Dash : OfficeStrokeDashStyle.Solid;
        }

        private static OfficeStrokeLineCap? MapNativeVmlStrokeLineCap(string? value) {
            if (string.IsNullOrWhiteSpace(value)) {
                return null;
            }

            if (value!.Equals("round", StringComparison.OrdinalIgnoreCase)) {
                return OfficeStrokeLineCap.Round;
            }

            if (value.Equals("square", StringComparison.OrdinalIgnoreCase)) {
                return OfficeStrokeLineCap.Square;
            }

            return value.Equals("flat", StringComparison.OrdinalIgnoreCase) ||
                   value.Equals("butt", StringComparison.OrdinalIgnoreCase)
                ? OfficeStrokeLineCap.Butt
                : null;
        }

        private static OfficeStrokeLineJoin? MapNativeVmlStrokeLineJoin(string? value) {
            if (string.IsNullOrWhiteSpace(value)) {
                return null;
            }

            if (value!.Equals("round", StringComparison.OrdinalIgnoreCase)) {
                return OfficeStrokeLineJoin.Round;
            }

            if (value.Equals("bevel", StringComparison.OrdinalIgnoreCase)) {
                return OfficeStrokeLineJoin.Bevel;
            }

            return value.Equals("miter", StringComparison.OrdinalIgnoreCase)
                ? OfficeStrokeLineJoin.Miter
                : null;
        }

        private static void ApplyNativeVmlShapeShadow(OfficeShape shape, OpenXmlElement element) {
            OpenXmlElement? shadowElement = GetNativeVmlChild(element, "shadow");
            if (shadowElement == null ||
                !IsNativeVmlSwitchEnabled(GetNativeOpenXmlAttribute(shadowElement, "on"))) {
                return;
            }

            PdfCore.PdfColor shadowColor = ParseNativeColor(NormalizeNativeVmlColor(GetNativeOpenXmlAttribute(shadowElement, "color"))) ?? PdfCore.PdfColor.Black;
            double opacity = ParseNativeVmlOpacity(GetNativeOpenXmlAttribute(shadowElement, "opacity")) ?? 0.5D;
            (double offsetX, double offsetY) = ParseNativeVmlOffset(GetNativeOpenXmlAttribute(shadowElement, "offset")) ?? (2D, 2D);
            shape.Shadow = new OfficeShadow(shadowColor.ToOfficeColor(), opacity, offsetX, offsetY);
        }

        private static (double X, double Y)? ParseNativeVmlOffset(string? value) {
            if (string.IsNullOrWhiteSpace(value)) {
                return null;
            }

            string[] parts = value!.Split(',');
            if (parts.Length != 2) {
                return null;
            }

            double? x = ResolveNativeVmlLength(parts[0], 1D, 1D);
            double? y = ResolveNativeVmlLength(parts[1], 1D, 1D);
            return x.HasValue && y.HasValue ? (x.Value, y.Value) : null;
        }

        private static void ApplyNativeVmlShapeTransform(OfficeShape shape, OpenXmlElement element) {
            OfficeTransform? transform = null;

            string? flip = GetNativeVmlFlip(element);
            if (!string.IsNullOrWhiteSpace(flip)) {
                bool horizontal = flip!.IndexOf("x", StringComparison.OrdinalIgnoreCase) >= 0;
                bool vertical = flip.IndexOf("y", StringComparison.OrdinalIgnoreCase) >= 0;
                if (horizontal || vertical) {
                    OfficeTransform flipTransform = CreateNativeVmlCenterScaleTransform(
                        shape.Width,
                        shape.Height,
                        horizontal ? -1D : 1D,
                        vertical ? -1D : 1D);
                    transform = transform.HasValue ? transform.Value.Then(flipTransform) : flipTransform;
                }
            }

            double? rotation = GetNativeVmlRotationDegrees(element);
            if (rotation.HasValue && Math.Abs(rotation.Value) > 0.0001D) {
                OfficeTransform rotationTransform = OfficeTransform.RotateDegrees(rotation.Value, shape.Width / 2D, shape.Height / 2D);
                transform = transform.HasValue ? transform.Value.Then(rotationTransform) : rotationTransform;
            }

            if (transform.HasValue) {
                shape.Transform = shape.Transform.HasValue ? shape.Transform.Value.Then(transform.Value) : transform.Value;
            }
        }

        private static OfficeTransform CreateNativeVmlCenterScaleTransform(double width, double height, double scaleX, double scaleY) {
            double centerX = width / 2D;
            double centerY = height / 2D;
            return OfficeTransform.Translate(-centerX, -centerY)
                .Then(OfficeTransform.Scale(scaleX, scaleY))
                .Then(OfficeTransform.Translate(centerX, centerY));
        }

        private static double? GetNativeVmlRotationDegrees(OpenXmlElement element) {
            Dictionary<string, string> style = ParseNativeVmlStyle(GetNativeOpenXmlAttribute(element, "style"));
            string? value = style.TryGetValue("rotation", out string? styleRotation)
                ? styleRotation
                : GetNativeOpenXmlAttribute(element, "rotation");

            if (string.IsNullOrWhiteSpace(value)) {
                return null;
            }

            string normalized = value!.Trim();
            if (normalized.EndsWith("fd", StringComparison.OrdinalIgnoreCase)) {
                double? fixedPoint = ParseNativeVmlDouble(normalized.Substring(0, normalized.Length - 2));
                return fixedPoint.HasValue ? fixedPoint.Value / 65536D : null;
            }

            if (normalized.EndsWith("deg", StringComparison.OrdinalIgnoreCase)) {
                normalized = normalized.Substring(0, normalized.Length - 3);
            }

            return ParseNativeVmlDouble(normalized);
        }

        private static string? GetNativeVmlFlip(OpenXmlElement element) {
            Dictionary<string, string> style = ParseNativeVmlStyle(GetNativeOpenXmlAttribute(element, "style"));
            return style.TryGetValue("flip", out string? styleFlip)
                ? styleFlip
                : GetNativeOpenXmlAttribute(element, "flip");
        }

        private static bool TryGetNativeVmlGradientFill(OpenXmlElement? fillElement, PdfCore.PdfColor? startFill, out OfficeLinearGradient? gradient) {
            gradient = null;
            if (fillElement == null) {
                return false;
            }

            string? type = GetNativeOpenXmlAttribute(fillElement, "type");
            if (string.IsNullOrWhiteSpace(type) ||
                type!.IndexOf("gradient", StringComparison.OrdinalIgnoreCase) < 0) {
                return false;
            }

            string? color2Value = NormalizeNativeVmlColor(GetNativeOpenXmlAttribute(fillElement, "color2"));
            TryGetNativeVmlGradientStopColors(fillElement, out PdfCore.PdfColor? stopStartFill, out PdfCore.PdfColor? stopEndFill);
            PdfCore.PdfColor? gradientStartFill = startFill ?? stopStartFill;
            PdfCore.PdfColor? endFill = ParseNativeColor(color2Value);
            if (!endFill.HasValue) {
                endFill = stopEndFill;
            }

            if (!gradientStartFill.HasValue || !endFill.HasValue) {
                return false;
            }

            gradient = CreateNativeVmlLinearGradient(
                gradientStartFill.Value.ToOfficeColor(),
                endFill.Value.ToOfficeColor(),
                GetNativeOpenXmlAttribute(fillElement, "angle"));
            return true;
        }

        private static bool TryGetNativeVmlGradientStopColors(OpenXmlElement fillElement, out PdfCore.PdfColor? startFill, out PdfCore.PdfColor? endFill) {
            startFill = null;
            endFill = null;
            string? colors = GetNativeOpenXmlAttribute(fillElement, "colors");
            if (string.IsNullOrWhiteSpace(colors)) {
                return false;
            }

            var stops = new List<(double Offset, PdfCore.PdfColor Color)>();
            foreach (string entry in colors!.Split(new[] { ';' }, StringSplitOptions.RemoveEmptyEntries)) {
                Match match = Regex.Match(entry.Trim(), "^(" + NativeVmlNumberPattern + "(?:%|f)?)\\s+(.+)$", RegexOptions.IgnoreCase);
                if (!match.Success) {
                    continue;
                }

                string offsetValue = match.Groups[1].Value;
                string offsetText = offsetValue.TrimEnd('%', 'f', 'F');
                if (!double.TryParse(offsetText, NumberStyles.Float, CultureInfo.InvariantCulture, out double offset)) {
                    continue;
                }

                if (offsetValue.EndsWith("%", StringComparison.OrdinalIgnoreCase)) {
                    offset /= 100D;
                } else if (offsetValue.EndsWith("f", StringComparison.OrdinalIgnoreCase)) {
                    offset /= 65536D;
                }

                string? colorValue = NormalizeNativeVmlColor(match.Groups[2].Value);
                PdfCore.PdfColor? color = ParseNativeColor(colorValue);
                if (!color.HasValue) {
                    continue;
                }

                stops.Add((Math.Max(0D, Math.Min(1D, offset)), color.Value));
            }

            if (stops.Count < 2) {
                return false;
            }

            stops.Sort((left, right) => left.Offset.CompareTo(right.Offset));
            startFill = stops[0].Color;
            endFill = stops[stops.Count - 1].Color;
            return true;
        }

        private static OfficeLinearGradient CreateNativeVmlLinearGradient(OfficeColor startColor, OfficeColor endColor, string? angleValue) {
            double angle = NormalizeNativeVmlAngleDegrees(ParseNativeVmlDouble(angleValue ?? string.Empty));
            if (IsNativeVmlAngleNear(angle, 0D) || IsNativeVmlAngleNear(angle, 180D)) {
                return OfficeLinearGradient.Horizontal(startColor, endColor);
            }

            return IsNativeVmlAngleNear(angle, 90D) || IsNativeVmlAngleNear(angle, 270D)
                ? OfficeLinearGradient.Vertical(startColor, endColor)
                : OfficeLinearGradient.DiagonalDown(startColor, endColor);
        }

        private static double NormalizeNativeVmlAngleDegrees(double? angle) {
            if (!angle.HasValue) {
                return 0D;
            }

            double degrees = angle.Value % 360D;
            return degrees < 0D ? degrees + 360D : degrees;
        }

        private static bool IsNativeVmlAngleNear(double angle, double target) {
            double distance = Math.Abs(angle - target);
            distance = Math.Min(distance, 360D - distance);
            return distance <= 22.5D;
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
                    foreach (OpenXmlElement child in run.ChildElements) {
                        if (child is W.Text text) {
                            AddNativeVmlTextRun(runs, document, run, properties, text);
                        } else if (child is W.Break) {
                            runs.Add(PdfCore.TextRun.LineBreak());
                        } else if (child is W.TabChar) {
                            runs.Add(PdfCore.TextRun.Tab());
                        } else {
                            foreach (W.Text nestedText in child.Descendants<W.Text>()) {
                                AddNativeVmlTextRun(runs, document, run, properties, nestedText);
                            }
                        }
                    }
                }
            }

            AddNativeVmlTextPathRuns(runs, document, element);

            while (runs.Count > 0 && runs[0].Text == "\n") {
                runs.RemoveAt(0);
            }

            while (runs.Count > 0 && runs[runs.Count - 1].Text == "\n") {
                runs.RemoveAt(runs.Count - 1);
            }

            return runs;
        }

        private static void AddNativeVmlTextRun(List<PdfCore.TextRun> runs, WordDocument document, W.Run run, W.RunProperties? properties, W.Text text) {
            string value = ResolveNativeVmlRunText(document, run, text);
            if (string.IsNullOrEmpty(value)) {
                return;
            }

            runs.Add(new PdfCore.TextRun(
                value,
                bold: HasNativeOnOff(properties?.Bold),
                underline: HasNativeVmlUnderline(properties?.Underline),
                color: ParseNativeColor(properties?.Color?.Val?.Value),
                italic: HasNativeOnOff(properties?.Italic),
                strike: HasNativeOnOff(properties?.Strike),
                fontSize: GetNativeVmlRunFontSize(properties),
                font: GetNativeVmlRunFont(properties)));
        }

        private static void AddNativeVmlTextPathRuns(List<PdfCore.TextRun> runs, WordDocument document, OpenXmlElement element) {
            foreach (V.TextPath textPath in element.Descendants<V.TextPath>()) {
                if (textPath.On != null && textPath.On.Value == false) {
                    continue;
                }

                string? raw = textPath.String?.Value;
                if (string.IsNullOrWhiteSpace(raw)) {
                    continue;
                }

                string value = ResolveNativeBuiltInPropertyPlaceholders(document, raw!);
                if (runs.Count > 0 && runs[runs.Count - 1].Text != "\n") {
                    runs.Add(PdfCore.TextRun.LineBreak());
                }

                runs.Add(new PdfCore.TextRun(
                    value,
                    bold: false,
                    underline: false,
                    color: null,
                    italic: false,
                    strike: false,
                    fontSize: GetNativeVmlTextPathFontSize(textPath),
                    font: GetNativeVmlTextPathFont(textPath)));
            }
        }

        private static double? GetNativeVmlTextPathFontSize(V.TextPath textPath) {
            Dictionary<string, string> style = ParseNativeVmlStyle(textPath.Style?.Value);
            double? fontSize = ResolveNativeVmlLength(style.TryGetValue("font-size", out string? value) ? value : null, 1D, 1D);
            return fontSize.HasValue && fontSize.Value > 0D && fontSize.Value <= MaxNativeVmlTextPathFontSizePoints
                ? fontSize
                : null;
        }

        private static PdfCore.PdfStandardFont? GetNativeVmlTextPathFont(V.TextPath textPath) {
            Dictionary<string, string> style = ParseNativeVmlStyle(textPath.Style?.Value);
            if (!style.TryGetValue("font-family", out string? family)) {
                return null;
            }

            family = family.Trim().Trim('"', '\'');
            return PdfCore.PdfStandardFontMapper.TryMapFontFamily(family, out PdfCore.PdfStandardFont font)
                ? font
                : null;
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
            bool hasHorizontalPosition = style.ContainsKey("mso-position-horizontal");
            bool hasVerticalPosition = style.ContainsKey("mso-position-vertical");
            bool hasExplicitX = xPercent.HasValue ||
                                HasNativeVmlExplicitPosition(style, "left", hasHorizontalPosition) ||
                                HasNativeVmlExplicitPosition(style, "margin-left", hasHorizontalPosition);
            bool hasExplicitY = yPercent.HasValue ||
                                HasNativeVmlExplicitPosition(style, "top", hasVerticalPosition) ||
                                HasNativeVmlExplicitPosition(style, "margin-top", hasVerticalPosition);
            double x = xPercent ??
                       ResolveNativeVmlPosition(style.TryGetValue("left", out string? left) ? left : null, frame.Width, frame.CoordWidth, frame.CoordOriginX) ??
                       ResolveNativeVmlPosition(style.TryGetValue("margin-left", out string? marginLeft) ? marginLeft : null, frame.Width, frame.CoordWidth, frame.CoordOriginX) ?? 0D;
            double y = yPercent ??
                       ResolveNativeVmlPosition(style.TryGetValue("top", out string? top) ? top : null, frame.Height, frame.CoordHeight, frame.CoordOriginY) ??
                       ResolveNativeVmlPosition(style.TryGetValue("margin-top", out string? marginTop) ? marginTop : null, frame.Height, frame.CoordHeight, frame.CoordOriginY) ?? 0D;

            if (element.LocalName.Equals("line", StringComparison.OrdinalIgnoreCase) &&
                (width <= 0D || height <= 0D) &&
                TryGetNativeVmlLineBounds(element, frame, out double lineX, out double lineY, out double lineWidth, out double lineHeight)) {
                x += lineX;
                y += lineY;
                width = Math.Max(width, lineWidth);
                height = Math.Max(height, lineHeight);
            }

            if (style.TryGetValue("mso-position-horizontal", out string? horizontalPosition) && !hasExplicitX) {
                if (horizontalPosition.Equals("center", StringComparison.OrdinalIgnoreCase)) {
                    x = (pageWidth - width) / 2D;
                } else if (horizontalPosition.Equals("right", StringComparison.OrdinalIgnoreCase)) {
                    x = pageWidth - width;
                }
            }

            if (style.TryGetValue("mso-position-vertical", out string? verticalPosition) &&
                !hasExplicitY) {
                if (verticalPosition.Equals("center", StringComparison.OrdinalIgnoreCase)) {
                    y = (pageHeight - height) / 2D;
                } else if (verticalPosition.Equals("bottom", StringComparison.OrdinalIgnoreCase)) {
                    y = pageHeight - height;
                }
            }

            box = new NativeVmlBox(frame.X + x, frame.Y + y, width, height);
            return width > 0D && height > 0D;
        }

        private static bool TryGetNativeVmlLineBounds(OpenXmlElement element, NativeVmlFrame frame, out double x, out double y, out double width, out double height) {
            (double x1, double y1) = ParseNativeShapePoint(GetNativeOpenXmlAttribute(element, "from") ?? "0pt,0pt", frame);
            (double x2, double y2) = ParseNativeShapePoint(GetNativeOpenXmlAttribute(element, "to") ?? "0pt,0pt", frame);
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

        private static (double X, double Y) ParseNativeShapePoint(string value, NativeVmlFrame frame) {
            string[] parts = value.Split(',');
            if (parts.Length != 2) {
                return (0D, 0D);
            }

            double x = ResolveNativeVmlPosition(parts[0], frame.Width, frame.CoordWidth, frame.CoordOriginX) ?? 0D;
            double y = ResolveNativeVmlPosition(parts[1], frame.Height, frame.CoordHeight, frame.CoordOriginY) ?? 0D;
            return (x, y);
        }

        private static bool HasNativeVmlExplicitPosition(Dictionary<string, string> style, string key, bool hasRelativePosition) {
            if (!style.TryGetValue(key, out string? value)) {
                return false;
            }

            if (hasRelativePosition) {
                double? resolved = ResolveNativeVmlPosition(value, 1D, 1D, 0D);
                if (resolved.HasValue && Math.Abs(resolved.Value) < 0.001D) {
                    return false;
                }
            }

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
            if (normalized.EndsWith("pt", StringComparison.OrdinalIgnoreCase)) return NormalizeNativeVmlLength(ParseNativeVmlDouble(normalized.Substring(0, normalized.Length - 2)));
            if (normalized.EndsWith("in", StringComparison.OrdinalIgnoreCase)) return MultiplyNativeVmlLength(ParseNativeVmlDouble(normalized.Substring(0, normalized.Length - 2)), 72D);
            if (normalized.EndsWith("cm", StringComparison.OrdinalIgnoreCase)) return MultiplyNativeVmlLength(ParseNativeVmlDouble(normalized.Substring(0, normalized.Length - 2)), 28.3464566929D);
            if (normalized.EndsWith("mm", StringComparison.OrdinalIgnoreCase)) return MultiplyNativeVmlLength(ParseNativeVmlDouble(normalized.Substring(0, normalized.Length - 2)), 2.83464566929D);
            if (normalized.EndsWith("px", StringComparison.OrdinalIgnoreCase)) return MultiplyNativeVmlLength(ParseNativeVmlDouble(normalized.Substring(0, normalized.Length - 2)), 0.75D);
            if (normalized.EndsWith("%", StringComparison.OrdinalIgnoreCase)) return MultiplyNativeVmlLength(ParseNativeVmlDouble(normalized.Substring(0, normalized.Length - 1)), parentSize / 100D);

            double? number = ParseNativeVmlDouble(normalized);
            if (!number.HasValue) {
                return null;
            }

            double resolved = parentCoord > 0D && Math.Abs(parentCoord - parentSize) > 0.01D
                ? number.Value / parentCoord * parentSize
                : number.Value;
            return NormalizeNativeVmlLength(resolved);
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

        private static bool IsNativeVmlNoColor(string? value) =>
            value?.Trim().Equals("none", StringComparison.OrdinalIgnoreCase) == true;

        private static string? GetNativeOpenXmlAttribute(OpenXmlElement element, string localName) {
            foreach (OpenXmlAttribute attribute in element.GetAttributes()) {
                if (attribute.LocalName.Equals(localName, StringComparison.OrdinalIgnoreCase)) {
                    return attribute.Value;
                }
            }

            return null;
        }

        private static bool IsNativeVmlSwitchEnabled(string? value) =>
            string.IsNullOrWhiteSpace(value) ||
            (!value!.Trim().Equals("f", StringComparison.OrdinalIgnoreCase) &&
             !value.Trim().Equals("false", StringComparison.OrdinalIgnoreCase) &&
             !value.Trim().Equals("0", StringComparison.OrdinalIgnoreCase));

        private static bool IsNativeVmlHidden(OpenXmlElement element) {
            Dictionary<string, string> style = ParseNativeVmlStyle(GetNativeOpenXmlAttribute(element, "style"));
            return style.TryGetValue("visibility", out string? value) &&
                   value.Equals("hidden", StringComparison.OrdinalIgnoreCase);
        }

        private static bool IsNativeVmlTextBoxShape(OpenXmlElement element) =>
            string.Equals(GetNativeOpenXmlAttribute(element, "type"), "#_x0000_t202", StringComparison.OrdinalIgnoreCase);

        private static double? ParseNativeVmlStrokeWeight(string? value) =>
            ResolveNativeVmlLength(value, 1D, 1D);

        private static double GetNativeVmlFirstAdjustment(OpenXmlElement element, double fallback) {
            string? value = GetNativeOpenXmlAttribute(element, "adj");
            if (string.IsNullOrWhiteSpace(value)) {
                return fallback;
            }

            string[] parts = value!.Split(new[] { ',', ';', ' ' }, StringSplitOptions.RemoveEmptyEntries);
            double? adjustment = parts.Length > 0 ? ParseNativeVmlDouble(parts[0]) : null;
            return adjustment.HasValue
                ? adjustment.Value
                : fallback;
        }

        private static double GetNativeVmlRoundRectCornerRadius(OpenXmlElement element, double width, double height) {
            const double defaultArcSize = 0.2D;
            double fraction = ParseNativeVmlOpacity(GetNativeOpenXmlAttribute(element, "arcsize")) ?? defaultArcSize;
            fraction = Math.Max(0D, Math.Min(0.5D, fraction));
            return Math.Min(width, height) * fraction;
        }

        private static double? ParseNativeVmlDouble(string value) {
            string normalized = value.Trim();
            if (double.TryParse(normalized, NumberStyles.Float, CultureInfo.InvariantCulture, out double result) && IsNativeVmlFinite(result)) {
                return result;
            }

            if (double.TryParse(normalized.Replace(',', '.'), NumberStyles.Float, CultureInfo.InvariantCulture, out result) && IsNativeVmlFinite(result)) {
                return result;
            }

            return null;
        }

        private static double? MultiplyNativeVmlLength(double? value, double factor) =>
            value.HasValue ? NormalizeNativeVmlLength(value.Value * factor) : null;

        private static double? NormalizeNativeVmlLength(double? value) {
            if (!value.HasValue || !IsNativeVmlFinite(value.Value) || Math.Abs(value.Value) > MaxNativeVmlLengthPoints) {
                return null;
            }

            return value.Value;
        }

        private static bool IsNativeVmlFinite(double value) =>
            !double.IsNaN(value) && !double.IsInfinity(value);

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
            } else if (normalized.EndsWith("f", StringComparison.OrdinalIgnoreCase)) {
                parsed = ParseNativeVmlDouble(normalized.Substring(0, normalized.Length - 1));
                if (parsed.HasValue) {
                    parsed /= 65536D;
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

        private static bool HasNativeVmlUnderline(W.Underline? underline) {
            if (underline == null) {
                return false;
            }

            if (underline.Val != null && underline.Val.Value == W.UnderlineValues.None) {
                return false;
            }

            string value = underline.Val?.Value.ToString() ?? string.Empty;
            if (value.Length == 0) {
                return true;
            }

            return !value.Equals("none", StringComparison.OrdinalIgnoreCase) &&
                   !value.Equals("0", StringComparison.OrdinalIgnoreCase) &&
                   !value.Equals("false", StringComparison.OrdinalIgnoreCase);
        }

        private static double? GetNativeVmlRunFontSize(W.RunProperties? properties) {
            string? value = properties?.FontSize?.Val?.Value;
            return double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out double halfPoints) && halfPoints > 0D
                ? halfPoints / 2D
                : null;
        }

        private static PdfCore.PdfStandardFont? GetNativeVmlRunFont(W.RunProperties? properties) {
            W.RunFonts? fonts = properties?.RunFonts;
            if (fonts == null) {
                return null;
            }

            foreach (string? family in new[] {
                fonts.Ascii?.Value,
                fonts.HighAnsi?.Value,
                fonts.EastAsia?.Value,
                fonts.ComplexScript?.Value
            }) {
                if (PdfCore.PdfStandardFontMapper.TryMapFontFamily(family, out PdfCore.PdfStandardFont font)) {
                    return font;
                }
            }

            return null;
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

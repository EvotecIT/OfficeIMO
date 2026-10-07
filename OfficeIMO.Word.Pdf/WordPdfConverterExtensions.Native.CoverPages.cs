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

        private static bool TryRenderNativeCoverPageCanvas(INativePdfFlow pdf, WordDocument document, W.SdtBlock? sdtBlock, WordSection section, WordToPdfOptions? options) {
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
            if (group.Ancestors<V.Group>().Count() >= 32)
                throw new InvalidDataException("Word shape group nesting exceeds the PDF export limit of 32 levels.");
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

            IReadOnlyList<PdfCore.PdfTextRun> textRuns = GetNativeVmlTextRuns(document, element);
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

        private static OpenXmlElement? GetNativeVmlChild(OpenXmlElement? element, string localName) {
            return element?.ChildElements.FirstOrDefault(child =>
                child.NamespaceUri == "urn:schemas-microsoft-com:vml" &&
                child.LocalName.Equals(localName, StringComparison.OrdinalIgnoreCase));
        }

    }
}

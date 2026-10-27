using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.PowerPoint {
    public partial class PowerPointSmartArt {
        private const int MaximumCachedDrawingShapes = 4096;

        /// <summary>
        /// Projects the SmartArt drawing Office caches beside the diagram data (<c>dsp:drawing</c>) into
        /// a detached group of ordinary slide shapes, positioned in the SmartArt frame.
        /// </summary>
        /// <remarks>
        /// Office writes the cached drawing so that consumers without the SmartArt layout engine can draw
        /// the diagram exactly as Office laid it out. The group is never added to the slide. The drawing is
        /// used only while its node text still matches the diagram data; a stale drawing reports a reason.
        /// </remarks>
        internal bool TryCreateCachedDrawing(out PowerPointGroupShape group, out string? failureReason) {
            group = null!;
            failureReason = null;
            try {
                DiagramDataPart dataPart = GetDiagramDataPart();
                XDocument dataDocument = LoadDiagramXDocument(dataPart);
                DiagramPersistLayoutPart? drawingPart = FindCachedDrawingPart(dataDocument);
                if (drawingPart == null) {
                    failureReason = "The SmartArt diagram has no cached drawing from its producer.";
                    return false;
                }
                OpenXmlCompositeElement? tree = drawingPart.Drawing?.ChildElements.OfType<OpenXmlCompositeElement>()
                    .FirstOrDefault(element => element.LocalName == "spTree");
                if (tree == null) {
                    failureReason = "The SmartArt cached drawing has no shape tree.";
                    return false;
                }
                if (!MatchesDiagramText(dataDocument, tree, out _)) {
                    failureReason = "The SmartArt cached drawing no longer matches the diagram text.";
                    return false;
                }
                Transform? frame = GraphicFrame.Transform;
                long x = frame?.Offset?.X?.Value ?? 0L, y = frame?.Offset?.Y?.Value ?? 0L;
                long cx = frame?.Extents?.Cx?.Value ?? 0L, cy = frame?.Extents?.Cy?.Value ?? 0L;
                if (cx <= 0L || cy <= 0L) {
                    failureReason = "The SmartArt frame has no size.";
                    return false;
                }
                if (OwnerSlide == null) {
                    failureReason = "The SmartArt diagram is not attached to a slide.";
                    return false;
                }
                var colors = new CachedDrawingColors(OwnerSlide);
                uint nextId = 1U;
                var rootTransform = new A.TransformGroup(
                    new A.Offset { X = x, Y = y }, new A.Extents { Cx = cx, Cy = cy },
                    new A.ChildOffset { X = 0L, Y = 0L }, new A.ChildExtents { Cx = cx, Cy = cy });
                if (frame?.Rotation?.Value is int frameRotation && frameRotation != 0) rootTransform.Rotation = frameRotation;
                if (frame?.HorizontalFlip?.Value == true) rootTransform.HorizontalFlip = true;
                if (frame?.VerticalFlip?.Value == true) rootTransform.VerticalFlip = true;
                GroupShape root = CreateGroup(Name ?? string.Empty, ref nextId, rootTransform);
                int count = 0;
                AppendShapes(root, tree, colors, ref nextId, ref count);
                group = (PowerPointGroupShape)new PowerPointGroupShape(root, drawingPart).AttachTo(OwnerSlide, drawingPart);
                return true;
            } catch (Exception exception) when (exception is InvalidOperationException or InvalidDataException or
                FormatException or System.Xml.XmlException or OpenXmlPackageException) {
                failureReason = "The SmartArt cached drawing could not be read: " + exception.Message;
                return false;
            }
        }

        private static void AppendShapes(GroupShape target, OpenXmlCompositeElement source, CachedDrawingColors colors,
            ref uint nextId, ref int count) {
            foreach (OpenXmlCompositeElement child in source.ChildElements.OfType<OpenXmlCompositeElement>()) {
                if (++count > MaximumCachedDrawingShapes)
                    throw new InvalidDataException("The SmartArt cached drawing exceeds the supported shape count.");
                if (child.LocalName == "sp") {
                    AppendShape(target, child, colors, ref nextId);
                } else if (child.LocalName == "grpSp") {
                    A.TransformGroup? transform = child.ChildElements.FirstOrDefault(element => element.LocalName == "grpSpPr")?
                        .GetFirstChild<A.TransformGroup>();
                    GroupShape nested = CreateGroup(string.Empty, ref nextId,
                        (A.TransformGroup?)transform?.CloneNode(true) ?? new A.TransformGroup());
                    AppendShapes(nested, child, colors, ref nextId, ref count);
                    target.Append(nested);
                }
            }
        }

        private static void AppendShape(GroupShape target, OpenXmlCompositeElement source, CachedDrawingColors colors,
            ref uint nextId) {
            OpenXmlElement? properties = source.ChildElements.FirstOrDefault(element => element.LocalName == "spPr");
            OpenXmlElement? style = source.ChildElements.FirstOrDefault(element => element.LocalName == "style");
            OpenXmlElement? text = source.ChildElements.FirstOrDefault(element => element.LocalName == "txBody");
            OpenXmlElement? textTransform = source.ChildElements.FirstOrDefault(element => element.LocalName == "txXfrm");
            A.Transform2D? shapeTransform = properties?.GetFirstChild<A.Transform2D>();
            bool separateText = text != null && textTransform != null && (
                ReadLong(textTransform, "off", "x") != (shapeTransform?.Offset?.X?.Value ?? 0L) ||
                ReadLong(textTransform, "off", "y") != (shapeTransform?.Offset?.Y?.Value ?? 0L) ||
                ReadLong(textTransform, "ext", "cx") != (shapeTransform?.Extents?.Cx?.Value ?? 0L) ||
                ReadLong(textTransform, "ext", "cy") != (shapeTransform?.Extents?.Cy?.Value ?? 0L) ||
                GetRotation(textTransform) != (shapeTransform?.Rotation?.Value ?? 0));
            var shape = new Shape(CreateShapeNames(ref nextId), new ShapeProperties(CloneChildren(properties)));
            if (style != null) shape.Append(new ShapeStyle(CloneChildren(style)));
            if (text != null && !separateText) shape.Append(colors.WithTextColor(new TextBody(CloneChildren(text)), style));
            target.Append(colors.Resolve(shape));
            if (!separateText) return;
            // The cached text rectangle (dsp:txXfrm) can differ from the shape, for example for the
            // label area of a chevron. Draw the text in its own unfilled rectangle at that position.
            var rectangle = new A.Transform2D(
                new A.Offset { X = ReadLong(textTransform!, "off", "x"), Y = ReadLong(textTransform!, "off", "y") },
                new A.Extents { Cx = ReadLong(textTransform!, "ext", "cx"), Cy = ReadLong(textTransform!, "ext", "cy") });
            int rotation = GetRotation(textTransform!);
            if (rotation != 0) rectangle.Rotation = rotation;
            var label = new Shape(CreateShapeNames(ref nextId), new ShapeProperties(rectangle,
                new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle },
                new A.NoFill(), new A.Outline(new A.NoFill())));
            if (style != null) label.Append(new ShapeStyle(CloneChildren(style)));
            label.Append(colors.WithTextColor(new TextBody(CloneChildren(text!)), style));
            target.Append(colors.Resolve(label));
        }

        /// <summary>
        /// Resolves the theme colours of the detached shapes against the slide's theme and colour map,
        /// so that every renderer sees the same literal colours Office drew.
        /// </summary>
        private sealed class CachedDrawingColors {
            private readonly A.ColorScheme? _scheme;
            private readonly OpenXmlElement? _colorMap;

            internal CachedDrawingColors(PowerPointSlide slide) {
                SlidePart part = slide.SlidePart;
                _scheme = part.ThemeOverridePart?.ThemeOverride?.ColorScheme
                    ?? part.SlideLayoutPart?.ThemeOverridePart?.ThemeOverride?.ColorScheme
                    ?? part.SlideLayoutPart?.SlideMasterPart?.ThemePart?.Theme?.ThemeElements?.ColorScheme;
                _colorMap = part.Slide?.ColorMapOverride?.GetFirstChild<A.OverrideColorMapping>()
                    ?? part.SlideLayoutPart?.SlideLayout?.ColorMapOverride?.GetFirstChild<A.OverrideColorMapping>()
                    ?? (OpenXmlElement?)part.SlideLayoutPart?.SlideMasterPart?.SlideMaster?.ColorMap;
            }

            internal Shape Resolve(Shape shape) {
                foreach (OpenXmlElement color in shape.Descendants().Where(IsColor).ToList()) {
                    if (color.Parent is A.LineReference or A.FillReference or A.EffectReference or A.FontReference) continue;
                    OfficeIMO.OfficeColor? resolved = OfficeIMO.OpenXml.Internal.OfficeOpenXmlThemeColorResolver
                        .ResolveColor(color, _scheme, colorMap: _colorMap);
                    if (resolved is not OfficeIMO.OfficeColor value) continue;
                    var literal = new A.RgbColorModelHex { Val = value.ToRgbHex() };
                    if (value.A < 255) literal.Append(new A.Alpha { Val = (int)Math.Round(value.A * 100000D / 255D) });
                    color.Parent!.ReplaceChild(literal, color);
                }
                return shape;
            }

            // SmartArt text takes its colour from the style's font reference unless a run sets one.
            internal TextBody WithTextColor(TextBody body, OpenXmlElement? style) {
                OpenXmlElement? fontColor = style?.ChildElements.OfType<A.FontReference>().FirstOrDefault()?
                    .ChildElements.FirstOrDefault(IsColor);
                if (fontColor == null) return body;
                foreach (OpenXmlCompositeElement properties in body.Descendants<OpenXmlCompositeElement>()
                    .Where(element => element is A.RunProperties or A.EndParagraphRunProperties).ToList()) {
                    if (properties.ChildElements.Any(child => child is A.SolidFill or A.NoFill or A.GradientFill)) continue;
                    properties.InsertAt(new A.SolidFill(fontColor.CloneNode(true)), 0);
                }
                return body;
            }

            private static bool IsColor(OpenXmlElement element) =>
                element is A.SchemeColor or A.RgbColorModelHex or A.RgbColorModelPercentage or A.HslColor or
                    A.SystemColor or A.PresetColor;
        }

        private static GroupShape CreateGroup(string name, ref uint nextId, A.TransformGroup transform) =>
            new GroupShape(
                new NonVisualGroupShapeProperties(
                    new NonVisualDrawingProperties { Id = nextId++, Name = name },
                    new NonVisualGroupShapeDrawingProperties(),
                    new ApplicationNonVisualDrawingProperties()),
                new GroupShapeProperties(transform));

        private static NonVisualShapeProperties CreateShapeNames(ref uint nextId) =>
            new NonVisualShapeProperties(
                new NonVisualDrawingProperties { Id = nextId, Name = "SmartArt shape " + nextId++ },
                new NonVisualShapeDrawingProperties(),
                new ApplicationNonVisualDrawingProperties());

        private static IEnumerable<OpenXmlElement> CloneChildren(OpenXmlElement? element) =>
            element?.ChildElements.Select(child => child.CloneNode(true)) ?? Enumerable.Empty<OpenXmlElement>();

        private static int GetRotation(OpenXmlElement transform) =>
            int.TryParse(transform.GetAttributes().FirstOrDefault(attribute => attribute.LocalName == "rot").Value,
                System.Globalization.NumberStyles.Integer, System.Globalization.CultureInfo.InvariantCulture, out int rotation)
                ? rotation : 0;

        private static long ReadLong(OpenXmlElement transform, string child, string attribute) {
            string? value = transform.ChildElements.FirstOrDefault(element => element.LocalName == child)?
                .GetAttributes().FirstOrDefault(item => item.LocalName == attribute).Value;
            return long.TryParse(value, System.Globalization.NumberStyles.Integer,
                System.Globalization.CultureInfo.InvariantCulture, out long result)
                ? result
                : throw new InvalidDataException("The SmartArt cached text rectangle is incomplete.");
        }

        private DiagramPersistLayoutPart? FindCachedDrawingPart(XDocument dataDocument) {
            XElement? extension = dataDocument.Descendants().FirstOrDefault(element =>
                element.Name.LocalName == "dataModelExt" && element.Name.NamespaceName == PersistedDiagramNamespace);
            string? id = (string?)extension?.Attribute("relId");
            return !string.IsNullOrWhiteSpace(id) && _slidePart.TryGetPartById(id!, out OpenXmlPart? part)
                ? part as DiagramPersistLayoutPart : null;
        }

        // Cached shapes reference presentation points, which map back to source nodes.
        // Comparing a bag of words loses node placement and significant whitespace.
        private static bool MatchesDiagramText(XDocument dataDocument, OpenXmlCompositeElement tree,
            out Dictionary<string, string> associations) {
            XNamespace a = "http://schemas.openxmlformats.org/drawingml/2006/main";
            XNamespace dgm = "http://schemas.openxmlformats.org/drawingml/2006/diagram";
            var nodes = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            associations = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            foreach (XElement point in dataDocument.Descendants(dgm + "pt")) {
                string? id = (string?)point.Attribute("modelId");
                if (string.IsNullOrEmpty(id)) return false;
                string type = (string?)point.Attribute("type") ?? "node";
                if (type == "node") {
                    if (nodes.ContainsKey(id!)) return false;
                    nodes.Add(id!, ReadVisibleText(point.Element(dgm + "t"), a));
                } else if (type == "pres") {
                    string? source = (string?)point.Element(dgm + "prSet")?.Attribute("presAssocID");
                    if (!string.IsNullOrEmpty(source)) associations[id!] = source!;
                }
            }
            // Some producers express the association through presOf connections.
            foreach (XElement connection in dataDocument.Descendants(dgm + "cxn")
                .Where(item => (string?)item.Attribute("type") == "presOf")) {
                string? source = (string?)connection.Attribute("srcId");
                string? target = (string?)connection.Attribute("destId");
                if (string.IsNullOrEmpty(source) || string.IsNullOrEmpty(target)) return false;
                if (associations.TryGetValue(target!, out string? previous) &&
                    !string.Equals(previous, source, StringComparison.OrdinalIgnoreCase)) return false;
                associations[target!] = source!;
            }
            var covered = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (OpenXmlCompositeElement shape in tree.Descendants<OpenXmlCompositeElement>()
                .Where(item => item.LocalName == "sp")) {
                OpenXmlElement? body = shape.ChildElements.FirstOrDefault(item => item.LocalName == "txBody");
                string text = body == null ? string.Empty : ReadVisibleText(XElement.Parse(body.OuterXml), a);
                string? id = shape.GetAttributes().FirstOrDefault(item => item.LocalName == "modelId").Value;
                if (id != null && associations.TryGetValue(id, out string? source)) id = source;
                if (id != null && nodes.TryGetValue(id, out string? expected)) {
                    if (!string.Equals(expected, text, StringComparison.Ordinal)) return false;
                    covered.Add(id);
                } else if (text.Length != 0) return false;
            }
            return nodes.All(node => node.Value.Length == 0 || covered.Contains(node.Key));
        }

        private static string ReadVisibleText(XElement? body, XNamespace drawingNamespace) => body == null
            ? string.Empty
            : string.Join("\n", body.Elements(drawingNamespace + "p").Select(paragraph =>
                string.Concat(paragraph.Descendants().Where(item => item.Name == drawingNamespace + "t" ||
                    item.Name == drawingNamespace + "br" || item.Name == drawingNamespace + "tab")
                    .Select(item => item.Name.LocalName == "br" ? "\n" : item.Name.LocalName == "tab" ? "\t" : item.Value))));
    }
}

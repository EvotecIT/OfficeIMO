using System.Globalization;
using System.Xml;
using System.Xml.Linq;
using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.OpenXml.Internal;
using OfficeIMO.PowerPoint.OpenXml;

namespace OfficeIMO.PowerPoint {
    public partial class PowerPointSmartArt {
        private const int MaximumCachedDrawingShapes = 4096;
        private const int MaximumCachedDrawingElements = 500_000;
        private const int MaximumCachedDrawingDepth = 128;
        private const int P = OfficeXmlNamespaces.PresentationML;
        private const int D = OfficeXmlNamespaces.DrawingML;

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
                PowerPointPartRef dataPart = GetDiagramDataPart();
                XDocument dataDocument = LoadDiagramXDocument(dataPart);
                if (!TryFindCachedDrawingPart(dataDocument, out PowerPointPartRef drawingPart)) {
                    failureReason = "The SmartArt diagram has no cached drawing from its producer.";
                    return false;
                }
                // Admission limits apply while the part streams in, before it is read whole.
                byte[]? drawingBytes = null;
                if (drawingPart.Current.IsSdk) {
                    using Stream stream = drawingPart.OpenRead();
                    ValidateCachedDrawing(stream);
                } else {
                    // The package's bytes for the part, refused unread when they are stored longer than the limits allow.
                    using var stream = (MemoryStream)drawingPart.OpenRead(MaximumCachedDrawingBytes);
                    ValidateCachedDrawing(stream);
                    drawingBytes = stream.ToArray();
                }
                OfficeXmlTree drawing = ParseCachedDrawing(ReadCachedDrawingMarkup(drawingPart, drawingBytes));
                int tree = FindShapeTree(drawing);
                if (tree == OfficeXmlTree.NoNode) {
                    failureReason = "The SmartArt cached drawing has no shape tree.";
                    return false;
                }
                if (!MatchesDiagramText(dataDocument, drawing, tree, out _)) {
                    failureReason = "The SmartArt cached drawing no longer matches the diagram text.";
                    return false;
                }
                OfficeXmlNode frame = XmlElement.Child(P, "xfrm");
                OfficeXmlNode offset = frame.Child(D, "off"), extents = frame.Child(D, "ext");
                long x = ReadInt64(offset, "x") ?? 0L, y = ReadInt64(offset, "y") ?? 0L;
                long cx = ReadInt64(extents, "cx") ?? 0L, cy = ReadInt64(extents, "cy") ?? 0L;
                if (cx <= 0L || cy <= 0L) {
                    failureReason = "The SmartArt frame has no size.";
                    return false;
                }
                if (OwnerSlide == null) {
                    failureReason = "The SmartArt diagram is not attached to a slide.";
                    return false;
                }
                var colors = new CachedDrawingColors(OwnerSlide);
                // The group is built in a tree of its own that shares the drawing's names, so copies need no mapping.
                OfficeXmlTree output = OfficeXmlTree.Create(drawing.Namespaces, P, "grpSp", 256);
                uint nextId = 1U;
                int rootTransform = output.CreateElement(D, "xfrm");
                int? frameRotation = ReadInt32(frame, "rot");
                if (frameRotation is int rotation && rotation != 0) output.SetAttribute(rootTransform, OfficeXmlNamespaces.None, "rot", Format(rotation));
                if (ReadBoolean(frame, "flipH") == true) output.SetAttribute(rootTransform, OfficeXmlNamespaces.None, "flipH", "1");
                if (ReadBoolean(frame, "flipV") == true) output.SetAttribute(rootTransform, OfficeXmlNamespaces.None, "flipV", "1");
                AppendPoint(output, rootTransform, "off", "x", x, "y", y);
                AppendPoint(output, rootTransform, "ext", "cx", cx, "cy", cy);
                AppendPoint(output, rootTransform, "chOff", "x", 0L, "y", 0L);
                AppendPoint(output, rootTransform, "chExt", "cx", cx, "cy", cy);
                FillGroup(output, output.Root, Name ?? string.Empty, ref nextId, rootTransform);
                int count = 0;
                AppendShapes(output, output.Root, drawing, tree, colors, ref nextId, ref count);
                RemoveFormattingWhitespace(output, output.Root);
                group = CreateCachedDrawingGroup(OwnerSlide, output.WriteElement(output.Root), drawingPart);
                return true;
            } catch (Exception exception) when (exception is InvalidOperationException or InvalidDataException or
                FormatException or XmlException) {
                failureReason = "The SmartArt cached drawing could not be read: " + exception.Message;
                return false;
            }
        }

        // The most bytes the character limit allows: four UTF-8 bytes a character, and a byte order mark.
        private const long MaximumCachedDrawingBytes = 4L * PowerPointXmlReader.MaximumPackageXmlCharacters + 4L;

        private static void ValidateCachedDrawing(Stream stream) {
            using XmlReader reader = XmlReader.Create(stream, new XmlReaderSettings {
                DtdProcessing = DtdProcessing.Prohibit,
                XmlResolver = null,
                MaxCharactersInDocument = PowerPointXmlReader.MaximumPackageXmlCharacters
            });
            int elements = 0;
            int shapes = 0;
            try {
                while (reader.Read()) {
                    if (reader.NodeType != XmlNodeType.Element) continue;
                    if (++elements > MaximumCachedDrawingElements || reader.Depth + 1 > MaximumCachedDrawingDepth)
                        throw new InvalidDataException("The SmartArt cached drawing exceeds the supported XML structure limits.");
                    if ((reader.LocalName is "sp" or "grpSp") && ++shapes > MaximumCachedDrawingShapes)
                        throw new InvalidDataException("The SmartArt cached drawing exceeds the supported shape count.");
                }
            } catch (XmlException exception) {
                throw new InvalidDataException("The SmartArt cached drawing failed XML admission limits or parsing.", exception);
            }
        }

        /// <summary>
        /// The cached drawing's markup: the SDK's element once the SDK holds the part, which is what the model reads and
        /// writes back, else the part's bytes. Empty when the part has no root element.
        /// </summary>
        private static byte[] ReadCachedDrawingMarkup(PowerPointPartRef part, byte[]? partBytes) {
            PowerPointPartRef current = part.Current;
            if (!current.IsSdk) return partBytes ?? ReadPartBytes(current);
            OfficeXmlNode root = current.Root;
            return root.IsNone ? Array.Empty<byte>() : root.OuterXml();
        }

        /// <summary>The cached drawing as a tree of its own; a part without a root element reads as an empty tree.</summary>
        private static OfficeXmlTree ParseCachedDrawing(byte[] markup) =>
            markup.Length == 0
                ? OfficeXmlTree.Create(new OfficeXmlNamespaces.Table(), OfficeXmlNamespaces.None, "empty", 1)
                : OfficeXmlTree.Parse(markup, markup.Length, new OfficeXmlNamespaces.Table());

        /// <summary>
        /// Removes whitespace-only text beside child elements, which formats the markup but isn't content, so a drawing
        /// read from indented bytes builds the same shapes as one read without formatting.
        /// </summary>
        private static void RemoveFormattingWhitespace(OfficeXmlTree tree, int element) {
            bool hasElements = false;
            for (int child = tree.FirstChildOf(element); child != OfficeXmlTree.NoNode; child = tree.NextSiblingOf(child)) {
                if (tree.KindOf(child) != OfficeXmlTreeNodeKind.Element) continue;
                hasElements = true;
                RemoveFormattingWhitespace(tree, child);
            }
            if (!hasElements) return;
            for (int child = tree.FirstChildOf(element); child != OfficeXmlTree.NoNode;) {
                int next = tree.NextSiblingOf(child);
                if (tree.KindOf(child) == OfficeXmlTreeNodeKind.Text && !tree.IsCData(child) && IsXmlWhitespace(tree.GetText(child))) tree.Remove(child);
                child = next;
            }
        }

        // The drawing root's first child element named spTree, in any namespace.
        private static int FindShapeTree(OfficeXmlTree drawing) {
            if (drawing.Root == OfficeXmlTree.NoNode) return OfficeXmlTree.NoNode;
            for (int child = drawing.FirstChildOf(drawing.Root); child != OfficeXmlTree.NoNode; child = drawing.NextSiblingOf(child)) {
                if (drawing.KindOf(child) == OfficeXmlTreeNodeKind.Element && drawing.LocalNameOf(child) == "spTree") return child;
            }
            return OfficeXmlTree.NoNode;
        }

        private static void AppendShapes(OfficeXmlTree output, int target, OfficeXmlTree drawing, int source, CachedDrawingColors colors,
            ref uint nextId, ref int count) {
            for (int child = drawing.FirstChildOf(source); child != OfficeXmlTree.NoNode; child = drawing.NextSiblingOf(child)) {
                if (drawing.KindOf(child) != OfficeXmlTreeNodeKind.Element) continue;
                if (++count > MaximumCachedDrawingShapes)
                    throw new InvalidDataException("The SmartArt cached drawing exceeds the supported shape count.");
                string localName = drawing.LocalNameOf(child);
                if (localName == "sp") {
                    AppendShape(output, target, drawing, child, colors, ref nextId);
                } else if (localName == "grpSp") {
                    int properties = FirstChildNamed(drawing, child, "grpSpPr");
                    int transform = properties == OfficeXmlTree.NoNode ? OfficeXmlTree.NoNode : drawing.FindChild(properties, D, "xfrm");
                    int nested = output.CreateElement(P, "grpSp");
                    FillGroup(output, nested, string.Empty, ref nextId,
                        transform == OfficeXmlTree.NoNode ? output.CreateElement(D, "xfrm") : output.ImportElement(drawing, transform));
                    AppendShapes(output, nested, drawing, child, colors, ref nextId, ref count);
                    output.Append(target, nested);
                }
            }
        }

        private static void AppendShape(OfficeXmlTree output, int target, OfficeXmlTree drawing, int source, CachedDrawingColors colors,
            ref uint nextId) {
            int properties = FirstChildNamed(drawing, source, "spPr");
            int style = FirstChildNamed(drawing, source, "style");
            int text = FirstChildNamed(drawing, source, "txBody");
            int textTransform = FirstChildNamed(drawing, source, "txXfrm");
            OfficeXmlNode shapeTransform = properties == OfficeXmlTree.NoNode
                ? default
                : new OfficeXmlNode(drawing, properties).Child(D, "xfrm");
            bool separateText = text != OfficeXmlTree.NoNode && textTransform != OfficeXmlTree.NoNode && (
                ReadLong(drawing, textTransform, "off", "x") != (ReadInt64(shapeTransform.Child(D, "off"), "x") ?? 0L) ||
                ReadLong(drawing, textTransform, "off", "y") != (ReadInt64(shapeTransform.Child(D, "off"), "y") ?? 0L) ||
                ReadLong(drawing, textTransform, "ext", "cx") != (ReadInt64(shapeTransform.Child(D, "ext"), "cx") ?? 0L) ||
                ReadLong(drawing, textTransform, "ext", "cy") != (ReadInt64(shapeTransform.Child(D, "ext"), "cy") ?? 0L) ||
                GetRotation(drawing, textTransform) != (ReadInt32(shapeTransform, "rot") ?? 0));
            int shape = CreateShape(output, ref nextId);
            int shapeProperties = output.CreateElement(P, "spPr");
            ImportChildren(output, shapeProperties, drawing, properties);
            output.Append(shape, shapeProperties);
            if (style != OfficeXmlTree.NoNode) output.Append(shape, CloneStyle(output, drawing, style));
            if (text != OfficeXmlTree.NoNode && !separateText) output.Append(shape, colors.WithTextColor(output, CloneTextBody(output, drawing, text), drawing, style));
            output.Append(target, colors.Resolve(output, shape));
            if (!separateText) return;
            // The cached text rectangle (dsp:txXfrm) can differ from the shape, for example for the
            // label area of a chevron. Draw the text in its own unfilled rectangle at that position.
            int rectangle = output.CreateElement(D, "xfrm");
            int labelRotation = GetRotation(drawing, textTransform);
            if (labelRotation != 0) output.SetAttribute(rectangle, OfficeXmlNamespaces.None, "rot", Format(labelRotation));
            AppendPoint(output, rectangle, "off", "x", ReadLong(drawing, textTransform, "off", "x"), "y", ReadLong(drawing, textTransform, "off", "y"));
            AppendPoint(output, rectangle, "ext", "cx", ReadLong(drawing, textTransform, "ext", "cx"), "cy", ReadLong(drawing, textTransform, "ext", "cy"));
            int label = CreateShape(output, ref nextId);
            int labelProperties = output.CreateElement(P, "spPr");
            output.Append(labelProperties, rectangle);
            int geometry = output.CreateElement(D, "prstGeom");
            output.SetAttribute(geometry, OfficeXmlNamespaces.None, "prst", "rect");
            output.Append(geometry, output.CreateElement(D, "avLst"));
            output.Append(labelProperties, geometry);
            output.Append(labelProperties, output.CreateElement(D, "noFill"));
            int outline = output.CreateElement(D, "ln");
            output.Append(outline, output.CreateElement(D, "noFill"));
            output.Append(labelProperties, outline);
            output.Append(label, labelProperties);
            if (style != OfficeXmlTree.NoNode) output.Append(label, CloneStyle(output, drawing, style));
            output.Append(label, colors.WithTextColor(output, CloneTextBody(output, drawing, text), drawing, style));
            output.Append(target, colors.Resolve(output, label));
        }

        /// <summary>
        /// Resolves the theme colours of the detached shapes against the slide's theme and colour map,
        /// so that every renderer sees the same literal colours Office drew.
        /// </summary>
        private sealed class CachedDrawingColors {
            private readonly OfficeXmlNode _scheme;
            private readonly OfficeXmlNode _colorMap;

            internal CachedDrawingColors(PowerPointSlide slide) {
                PowerPointPartRef part = slide.PartRef;
                PowerPointThemeScope scope = PowerPointThemeScope.For(part);
                _scheme = scope.ColorScheme;
                // The slide's override mapping, else its layout's, else the master's map.
                PowerPointPartRef layout = part.RelatedPart(PowerPointPartRef.SlideLayoutRelationshipType);
                OfficeXmlNode map = part.Root.Child(P, "clrMapOvr").Child(D, "overrideClrMapping");
                if (map.IsNone && !layout.IsNone) map = layout.Root.Child(P, "clrMapOvr").Child(D, "overrideClrMapping");
                if (map.IsNone && !scope.Master.IsNone) map = scope.Master.Root.Child(P, "clrMap");
                _colorMap = map;
            }

            internal int Resolve(OfficeXmlTree output, int shape) {
                var found = new List<int>();
                CollectColors(output, shape, found);
                foreach (int color in found) {
                    int parent = output.ParentOf(color);
                    if (parent != OfficeXmlTree.NoNode && output.NamespaceOf(parent) == D
                        && output.LocalNameOf(parent) is "lnRef" or "fillRef" or "effectRef" or "fontRef") continue;
                    OfficeIMO.OfficeColor? resolved = OfficeXmlThemeColors.ResolveColor(new OfficeXmlNode(output, color), _scheme, colorMap: _colorMap);
                    if (resolved is not OfficeIMO.OfficeColor value) continue;
                    int literal = output.CreateElement(D, "srgbClr");
                    output.SetAttribute(literal, OfficeXmlNamespaces.None, "val", value.ToRgbHex());
                    if (value.A < 255) {
                        int alpha = output.CreateElement(D, "alpha");
                        output.SetAttribute(alpha, OfficeXmlNamespaces.None, "val", Format((int)Math.Round(value.A * 100000D / 255D)));
                        output.Append(literal, alpha);
                    }
                    output.InsertBefore(parent, literal, color);
                    output.Remove(color);
                }
                return shape;
            }

            // SmartArt text takes its colour from the style's font reference unless a run sets one.
            internal int WithTextColor(OfficeXmlTree output, int body, OfficeXmlTree drawing, int style) {
                int fontReference = style == OfficeXmlTree.NoNode ? OfficeXmlTree.NoNode : drawing.FindChild(style, D, "fontRef");
                int fontColor = OfficeXmlTree.NoNode;
                for (int child = fontReference == OfficeXmlTree.NoNode ? OfficeXmlTree.NoNode : drawing.FirstChildOf(fontReference);
                     child != OfficeXmlTree.NoNode; child = drawing.NextSiblingOf(child)) {
                    if (IsColor(drawing, child)) {
                        fontColor = child;
                        break;
                    }
                }
                if (fontColor == OfficeXmlTree.NoNode) return body;
                var runProperties = new List<int>();
                foreach (int element in Descendants(output, body)) {
                    if (output.NamespaceOf(element) == D && output.LocalNameOf(element) is "rPr" or "endParaRPr") runProperties.Add(element);
                }
                foreach (int properties in runProperties) {
                    if (output.FindChild(properties, D, "solidFill") != OfficeXmlTree.NoNode
                        || output.FindChild(properties, D, "noFill") != OfficeXmlTree.NoNode
                        || output.FindChild(properties, D, "gradFill") != OfficeXmlTree.NoNode) continue;
                    int fill = output.CreateElement(D, "solidFill");
                    output.Append(fill, output.ImportElement(drawing, fontColor));
                    output.InsertBefore(properties, fill, output.FirstChildOf(properties));
                }
                return body;
            }

            private static void CollectColors(OfficeXmlTree tree, int element, List<int> found) {
                foreach (int node in Descendants(tree, element)) {
                    if (IsColor(tree, node)) found.Add(node);
                }
            }

            private static bool IsColor(OfficeXmlTree tree, int node) =>
                tree.KindOf(node) == OfficeXmlTreeNodeKind.Element && tree.NamespaceOf(node) == D
                && tree.LocalNameOf(node) is "schemeClr" or "srgbClr" or "scrgbClr" or "hslClr" or "sysClr" or "prstClr";
        }

        // Fills a detached p:grpSp with its non-visual properties and group properties around a transform.
        private static void FillGroup(OfficeXmlTree output, int group, string name, ref uint nextId, int transform) {
            int nonVisual = output.CreateElement(P, "nvGrpSpPr");
            int drawingProperties = output.CreateElement(P, "cNvPr");
            output.SetAttribute(drawingProperties, OfficeXmlNamespaces.None, "id", Format(nextId++));
            output.SetAttribute(drawingProperties, OfficeXmlNamespaces.None, "name", name);
            output.Append(nonVisual, drawingProperties);
            output.Append(nonVisual, output.CreateElement(P, "cNvGrpSpPr"));
            output.Append(nonVisual, output.CreateElement(P, "nvPr"));
            output.Append(group, nonVisual);
            int properties = output.CreateElement(P, "grpSpPr");
            output.Append(properties, transform);
            output.Append(group, properties);
        }

        private static int CreateShape(OfficeXmlTree output, ref uint nextId) {
            int shape = output.CreateElement(P, "sp");
            int nonVisual = output.CreateElement(P, "nvSpPr");
            int drawingProperties = output.CreateElement(P, "cNvPr");
            output.SetAttribute(drawingProperties, OfficeXmlNamespaces.None, "id", Format(nextId));
            output.SetAttribute(drawingProperties, OfficeXmlNamespaces.None, "name", "SmartArt shape " + Format(nextId++));
            output.Append(nonVisual, drawingProperties);
            output.Append(nonVisual, output.CreateElement(P, "cNvSpPr"));
            output.Append(nonVisual, output.CreateElement(P, "nvPr"));
            output.Append(shape, nonVisual);
            return shape;
        }

        private static int CloneStyle(OfficeXmlTree output, OfficeXmlTree drawing, int style) {
            int clone = output.CreateElement(P, "style");
            ImportChildren(output, clone, drawing, style);
            return clone;
        }

        /// <summary>Clones a cached text body, keeping text that a CDATA section splits within one a:t element.</summary>
        private static int CloneTextBody(OfficeXmlTree output, OfficeXmlTree drawing, int source) {
            int body = output.CreateElement(P, "txBody");
            ImportChildren(output, body, drawing, source);
            List<int> sourceTexts = Descendants(drawing, source).Where(node => IsDrawingText(drawing, node)).ToList();
            List<int> clonedTexts = Descendants(output, body).Where(node => IsDrawingText(output, node)).ToList();
            for (int index = 0; index < sourceTexts.Count && index < clonedTexts.Count; index++) {
                if (!HasCData(drawing, sourceTexts[index])) continue;
                // The copy holds every segment as ordinary text; one that reads differently from the source's text is
                // written anew, without the attributes the split one had.
                string value = ReadFullText(drawing, sourceTexts[index]);
                if (string.Equals(value, output.GetInnerText(clonedTexts[index]), StringComparison.Ordinal)) continue;
                int replacement = output.CreateElement(D, "t");
                if (value.Length > 0) output.Append(replacement, output.CreateText(value));
                output.InsertBefore(output.ParentOf(clonedTexts[index]), replacement, clonedTexts[index]);
                output.Remove(clonedTexts[index]);
            }
            return body;
        }

        // Copies the element children of a drawing element; text and comments between them are not content.
        private static void ImportChildren(OfficeXmlTree output, int target, OfficeXmlTree drawing, int source) {
            if (source == OfficeXmlTree.NoNode) return;
            for (int child = drawing.FirstChildOf(source); child != OfficeXmlTree.NoNode; child = drawing.NextSiblingOf(child)) {
                if (drawing.KindOf(child) == OfficeXmlTreeNodeKind.Element) output.Append(target, output.ImportElement(drawing, child));
            }
        }

        private static void AppendPoint(OfficeXmlTree output, int parent, string name, string first, long firstValue, string second, long secondValue) {
            int point = output.CreateElement(D, name);
            output.SetAttribute(point, OfficeXmlNamespaces.None, first, Format(firstValue));
            output.SetAttribute(point, OfficeXmlNamespaces.None, second, Format(secondValue));
            output.Append(parent, point);
        }

        /// <summary>The element descendants of an element in document order, the element itself excluded.</summary>
        private static List<int> Descendants(OfficeXmlTree tree, int element) {
            var result = new List<int>();
            var pending = new Stack<int>();
            for (int child = tree.LastChildOf(element); child != OfficeXmlTree.NoNode; child = tree.PreviousSiblingOf(child)) pending.Push(child);
            while (pending.Count > 0) {
                int node = pending.Pop();
                if (tree.KindOf(node) != OfficeXmlTreeNodeKind.Element) continue;
                result.Add(node);
                for (int child = tree.LastChildOf(node); child != OfficeXmlTree.NoNode; child = tree.PreviousSiblingOf(child)) pending.Push(child);
            }
            return result;
        }

        private static int FirstChildNamed(OfficeXmlTree tree, int parent, string localName) {
            for (int child = tree.FirstChildOf(parent); child != OfficeXmlTree.NoNode; child = tree.NextSiblingOf(child)) {
                if (tree.KindOf(child) == OfficeXmlTreeNodeKind.Element && tree.LocalNameOf(child) == localName) return child;
            }
            return OfficeXmlTree.NoNode;
        }

        private static bool IsDrawingText(OfficeXmlTree tree, int node) => tree.NamespaceOf(node) == D && tree.LocalNameOf(node) == "t";

        private static string Format(long value) => value.ToString(CultureInfo.InvariantCulture);

        private static string Format(uint value) => value.ToString(CultureInfo.InvariantCulture);

        private static string Format(int value) => value.ToString(CultureInfo.InvariantCulture);

        // Typed attribute values as the schema reads them: a value that isn't one is a format error, a missing one null.
        private static long? ReadInt64(OfficeXmlNode element, string localName) =>
            element.Attribute(localName) is { } value ? XmlConvert.ToInt64(value) : null;

        private static int? ReadInt32(OfficeXmlNode element, string localName) =>
            element.Attribute(localName) is { } value ? XmlConvert.ToInt32(value) : null;

        private static bool? ReadBoolean(OfficeXmlNode element, string localName) =>
            element.Attribute(localName) is { } value ? XmlConvert.ToBoolean(value) : null;

        private static int GetRotation(OfficeXmlTree tree, int transform) =>
            int.TryParse(tree.GetAttribute(transform, OfficeXmlNamespaces.None, "rot"), NumberStyles.Integer, CultureInfo.InvariantCulture, out int rotation)
                ? rotation : 0;

        private static long ReadLong(OfficeXmlTree tree, int transform, string child, string attribute) {
            int element = FirstChildNamed(tree, transform, child);
            string? value = element == OfficeXmlTree.NoNode ? null : tree.GetAttribute(element, OfficeXmlNamespaces.None, attribute);
            return long.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out long result)
                ? result
                : throw new InvalidDataException("The SmartArt cached text rectangle is incomplete.");
        }

        private bool TryFindCachedDrawingPart(XDocument dataDocument, out PowerPointPartRef part) {
            part = default;
            XElement? extension = dataDocument.Descendants().FirstOrDefault(element =>
                element.Name.LocalName == "dataModelExt" && element.Name.NamespaceName == PersistedDiagramNamespace);
            string? id = (string?)extension?.Attribute("relId");
            return !string.IsNullOrWhiteSpace(id) && TryGetDiagramPart(id, DiagramDrawingRelationshipType, out part);
        }

        // Cached shapes reference presentation points, which map back to source nodes.
        // Comparing a bag of words loses node placement and significant whitespace.
        private static bool MatchesDiagramText(XDocument dataDocument, OfficeXmlTree drawing, int tree,
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
            foreach (int shape in Descendants(drawing, tree)) {
                if (drawing.LocalNameOf(shape) != "sp") continue;
                int body = FirstChildNamed(drawing, shape, "txBody");
                string text = body == OfficeXmlTree.NoNode ? string.Empty : ReadVisibleText(drawing, body);
                string? id = drawing.GetAttribute(shape, OfficeXmlNamespaces.None, "modelId");
                if (id != null && associations.TryGetValue(id, out string? source)) id = source;
                if (id != null && nodes.TryGetValue(id, out string? expected)) {
                    if (!string.Equals(expected, text, StringComparison.Ordinal)) return false;
                    covered.Add(id);
                } else if (text.Length != 0) return false;
            }
            return nodes.All(node => node.Value.Length == 0 || covered.Contains(node.Key));
        }

        /// <summary>
        ///     Reads cached-drawing text from the drawing tree with the same result as
        ///     <see cref="ReadVisibleText(XElement?, XNamespace)" /> over a parsed copy, without serializing the shape.
        /// </summary>
        private static string ReadVisibleText(OfficeXmlTree drawing, int body) {
            var text = new System.Text.StringBuilder();
            bool firstParagraph = true;
            for (int paragraph = drawing.FirstChildOf(body); paragraph != OfficeXmlTree.NoNode; paragraph = drawing.NextSiblingOf(paragraph)) {
                if (!drawing.IsElement(paragraph, D, "p")) continue;
                if (!firstParagraph) text.Append('\n');
                firstParagraph = false;
                foreach (int item in Descendants(drawing, paragraph)) {
                    if (drawing.NamespaceOf(item) != D) continue;
                    switch (drawing.LocalNameOf(item)) {
                        case "br": text.Append('\n'); break;
                        case "tab": text.Append('\t'); break;
                        case "t": text.Append(ParsedTextValue(drawing, item)); break;
                    }
                }
            }
            return text.ToString();
        }

        /// <summary>A parse without whitespace preservation drops whitespace-only text unless xml:space keeps it.</summary>
        private static string ParsedTextValue(OfficeXmlTree drawing, int textElement) {
            // A parse applies whitespace handling to each text node, and it always keeps CDATA content.
            if (HasCData(drawing, textElement)) return ReadFullText(drawing, textElement);
            string value = drawing.GetInnerText(textElement);
            if (value.Length == 0 || !IsXmlWhitespace(value)) return value;
            return PreservesWhitespace(drawing, textElement) ? value : string.Empty;
        }

        /// <summary>
        /// The text of a text element whose content a CDATA section splits, as a parse of the element reads it: every
        /// CDATA section and every other text node, except whitespace-only ones outside an <c>xml:space="preserve"</c> scope.
        /// </summary>
        private static string ReadFullText(OfficeXmlTree drawing, int textElement) {
            bool preserve = PreservesWhitespace(drawing, textElement);
            var text = new System.Text.StringBuilder();
            for (int child = drawing.FirstChildOf(textElement); child != OfficeXmlTree.NoNode; child = drawing.NextSiblingOf(child)) {
                if (drawing.KindOf(child) != OfficeXmlTreeNodeKind.Text) continue;
                string value = drawing.GetText(child);
                if (drawing.IsCData(child) || preserve || !IsXmlWhitespace(value)) text.Append(value);
            }
            return text.ToString();
        }

        private static bool HasCData(OfficeXmlTree drawing, int element) {
            for (int child = drawing.FirstChildOf(element); child != OfficeXmlTree.NoNode; child = drawing.NextSiblingOf(child)) {
                if (drawing.KindOf(child) == OfficeXmlTreeNodeKind.Text && drawing.IsCData(child)) return true;
            }
            return false;
        }

        private static bool IsXmlWhitespace(string value) {
            foreach (char character in value) {
                if (character != ' ' && character != '\t' && character != '\r' && character != '\n') return false;
            }
            return true;
        }

        private static bool PreservesWhitespace(OfficeXmlTree drawing, int element) {
            for (int scope = element; scope != OfficeXmlTree.NoNode; scope = drawing.ParentOf(scope)) {
                string? space = drawing.GetAttribute(scope, OfficeXmlNamespaces.Xml, "space");
                if (space == "preserve") return true;
                if (space == "default") return false;
            }
            return false;
        }

        private static string ReadVisibleText(XElement? body, XNamespace drawingNamespace) => body == null
            ? string.Empty
            : string.Join("\n", body.Elements(drawingNamespace + "p").Select(paragraph =>
                string.Concat(paragraph.Descendants().Where(item => item.Name == drawingNamespace + "t" ||
                    item.Name == drawingNamespace + "br" || item.Name == drawingNamespace + "tab")
                    .Select(item => item.Name.LocalName == "br" ? "\n" : item.Name.LocalName == "tab" ? "\t" : item.Value))));
    }
}

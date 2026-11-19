using System.Globalization;
using System.IO.Packaging;
using System.Text;
using System.Xml;
using DocumentFormat.OpenXml;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace OfficeIMO.PowerPoint {
    public sealed partial class PowerPointSlideReader {
        // The same shape and nesting limits PowerPointSlide.EnumerateShapesDeep applies to a loaded slide.
        internal const int MaximumShapeCount = 10000;
        internal const int MaximumGroupDepth = 128;

        private const string PresentationNamespace = "http://schemas.openxmlformats.org/presentationml/2006/main";
        private const string StrictPresentationNamespace = "http://purl.oclc.org/ooxml/presentationml/main";
        private const string DrawingNamespace = "http://schemas.openxmlformats.org/drawingml/2006/main";
        private const string StrictDrawingNamespace = "http://purl.oclc.org/ooxml/drawingml/main";
        private const string RelationshipNamespace = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        private const string StrictRelationshipNamespace = "http://purl.oclc.org/ooxml/officeDocument/relationships";
        private const string ChartNamespace = "http://schemas.openxmlformats.org/drawingml/2006/chart";
        private const string StrictChartNamespace = "http://purl.oclc.org/ooxml/drawingml/chart";
        private const string DiagramNamespace = "http://schemas.openxmlformats.org/drawingml/2006/diagram";
        private const string StrictDiagramNamespace = "http://purl.oclc.org/ooxml/drawingml/diagram";

        private static IReadOnlyList<SlideEntry> ReadSlideList(Package package, PackagePart presentationPart,
            XmlReaderSettings settings, CancellationToken cancellationToken) {
            var slides = new List<SlideEntry>();
            using Stream stream = presentationPart.GetStream(FileMode.Open, FileAccess.Read);
            using XmlReader reader = XmlReader.Create(stream, settings);
            if (reader.MoveToContent() != XmlNodeType.Element) return slides;
            bool listSeen = false;
            ForEachChild(reader, () => {
                if (listSeen || !IsPresentation(reader, "sldIdLst")) return false;
                listSeen = true;
                ForEachChild(reader, () => {
                    if (!IsPresentation(reader, "sldId")) return false;
                    cancellationToken.ThrowIfCancellationRequested();
                    string? relationshipId = reader.GetAttribute("id", RelationshipNamespace);
                    if (string.IsNullOrEmpty(relationshipId)) relationshipId = reader.GetAttribute("id", StrictRelationshipNamespace);
                    if (string.IsNullOrEmpty(relationshipId)) return false;
                    PackagePart? slidePart = presentationPart.RelationshipExists(relationshipId!)
                        && presentationPart.GetRelationship(relationshipId!) is { } relationship
                        && IsRelationshipType(relationship, "slide")
                        ? GetTargetPart(package, relationship)
                        : null;
                    if (slidePart == null) {
                        throw new InvalidDataException($"The slide list entry '{relationshipId}' does not target a slide part.");
                    }
                    slides.Add(new SlideEntry(slidePart, reader.GetAttribute("show")));
                    return false;
                });
                return true;
            });
            return slides;
        }

        private PowerPointSlideContent ReadSlide(int index, bool includeNotes, CancellationToken cancellationToken) {
            SlideEntry entry = _slides[index];
            var context = new SlideContext(this, entry.Part, cancellationToken);
            bool? show = null;
            using (Stream stream = entry.Part.GetStream(FileMode.Open, FileAccess.Read))
            using (XmlReader reader = XmlReader.Create(stream, _xmlSettings)) {
                if (reader.MoveToContent() == XmlNodeType.Element) {
                    show = ReadBoolean(reader.GetAttribute("show"));
                    bool commonSlideDataSeen = false;
                    ForEachChild(reader, () => {
                        if (commonSlideDataSeen || !IsPresentation(reader, "cSld")) return false;
                        commonSlideDataSeen = true;
                        bool treeSeen = false;
                        ForEachChild(reader, () => {
                            if (treeSeen || !IsPresentation(reader, "spTree")) return false;
                            treeSeen = true;
                            ReadShapeTree(reader, context, parent: null, depth: 0);
                            return true;
                        });
                        return true;
                    });
                }
            }

            bool hidden = show.HasValue ? !show.Value : IsHiddenShowValue(entry.ListShowValue);
            string? notes = includeNotes ? ReadNotes(_package, entry.Part, _xmlSettings) : null;
            return new PowerPointSlideContent(index + 1, hidden, context.Build(), notes);
        }

        private static bool IsHiddenShowValue(string? value) =>
            string.Equals(value, "0", StringComparison.Ordinal) || string.Equals(value, "false", StringComparison.OrdinalIgnoreCase);

        // Reads the children of a shape tree or group in drawing order; the caller is positioned on its start element.
        private static void ReadShapeTree(XmlReader reader, SlideContext context, ShapeBuilder? parent, int depth) {
            ForEachChild(reader, () => {
                if (!IsPresentationNamespace(reader.NamespaceURI)) return false;
                switch (reader.LocalName) {
                    case "sp":
                    case "cxnSp":
                    case "pic":
                    case "grpSp":
                    case "graphicFrame":
                        ReadShape(reader, context, parent, depth);
                        return true;
                    default:
                        return false;
                }
            });
        }

        private static void ReadShape(XmlReader reader, SlideContext context, ShapeBuilder? parent, int depth) {
            context.CancellationToken.ThrowIfCancellationRequested();
            if (depth > MaximumGroupDepth) {
                throw new InvalidOperationException("The grouped-shape nesting depth exceeds the configured limit.");
            }
            string element = reader.LocalName;
            var shape = new ShapeBuilder(parent, depth);
            int position = context.Reserve(shape);
            switch (element) {
                case "sp":
                    ForEachChild(reader, () => {
                        if (IsPresentation(reader, "nvSpPr")) return ReadNonVisualProperties(reader, shape, readPlaceholder: true);
                        if (IsPresentation(reader, "txBody") && shape.Paragraphs == null) {
                            shape.Paragraphs = ReadTextBody(reader, context);
                            return true;
                        }
                        return false;
                    });
                    shape.Kind = shape.Paragraphs != null ? PowerPointShapeContentType.TextBox : PowerPointShapeContentType.AutoShape;
                    break;
                case "cxnSp":
                    ForEachChild(reader, () => IsPresentation(reader, "nvCxnSpPr") && ReadNonVisualProperties(reader, shape, readPlaceholder: true));
                    shape.Kind = PowerPointShapeContentType.Connector;
                    break;
                case "pic":
                    bool fillSeen = false;
                    ForEachChild(reader, () => {
                        if (IsPresentation(reader, "nvPicPr")) return ReadNonVisualProperties(reader, shape, readPlaceholder: true, detectMedia: true);
                        if (IsPresentation(reader, "blipFill") && !fillSeen) {
                            fillSeen = true;
                            ReadBlipFill(reader, context, shape);
                            return true;
                        }
                        return false;
                    });
                    shape.Kind = shape.IsMedia ? PowerPointShapeContentType.Media : PowerPointShapeContentType.Picture;
                    break;
                case "grpSp":
                    shape.Kind = PowerPointShapeContentType.Group;
                    ForEachChild(reader,
                        () => IsPresentation(reader, "nvGrpSpPr") && ReadNonVisualProperties(reader, shape, readPlaceholder: false),
                        child => IsPresentationNamespace(child.NamespaceURI) && child.LocalName is "sp" or "cxnSp" or "pic" or "grpSp" or "graphicFrame",
                    () => ReadShape(reader, context, shape, depth + 1));
                    break;
                case "graphicFrame":
                    ForEachChild(reader, () => {
                        if (IsPresentation(reader, "nvGraphicFramePr")) return ReadNonVisualProperties(reader, shape, readPlaceholder: true);
                        if (IsDrawing(reader, "graphic") && !shape.GraphicSeen) {
                            shape.GraphicSeen = true;
                            ReadGraphic(reader, context, shape);
                            return true;
                        }
                        return false;
                    });
                    // A frame holding none of these is not a slide shape in the presentation model either.
                    if (shape.FrameKind == null) {
                        context.Release(position);
                        return;
                    }
                    shape.Kind = shape.FrameKind.Value;
                    break;
            }
            context.Count();
        }

        // Reads the first non-visual drawing properties and placeholder of a shape's non-visual property group.
        private static bool ReadNonVisualProperties(XmlReader reader, ShapeBuilder shape, bool readPlaceholder,
            bool detectMedia = false) {
            if (shape.NonVisualSeen) return false;
            shape.NonVisualSeen = true;
            bool drawingSeen = false, applicationSeen = false;
            ForEachChild(reader, () => {
                if (IsPresentation(reader, "cNvPr") && !drawingSeen) {
                    drawingSeen = true;
                    string? id = reader.GetAttribute("id");
                    shape.Id = id == null ? null : ParseUInt32(id);
                    shape.Name = reader.GetAttribute("name");
                    shape.AltText = reader.GetAttribute("descr");
                    shape.Hidden = ReadBoolean(reader.GetAttribute("hidden")) == true;
                    return false;
                }
                if (IsPresentation(reader, "nvPr") && !applicationSeen) {
                    applicationSeen = true;
                    bool placeholderSeen = false;
                    int propertiesDepth = reader.Depth;
                    ReadDescendants(reader, () => {
                        if (readPlaceholder && !placeholderSeen && reader.Depth == propertiesDepth + 1 && IsPresentation(reader, "ph")) {
                            placeholderSeen = true;
                            string? type = reader.GetAttribute("type");
                            if (type != null) shape.PlaceholderType = ParsePlaceholderType(type);
                        }
                        if (detectMedia && IsDrawingNamespace(reader.NamespaceURI)
                            && reader.LocalName is "audioFile" or "videoFile") {
                            shape.IsMedia = true;
                        }
                    });
                    return true;
                }
                return false;
            });
            return true;
        }

        private static void ReadBlipFill(XmlReader reader, SlideContext context, ShapeBuilder shape) {
            bool blipSeen = false;
            ForEachChild(reader, () => {
                if (!IsDrawing(reader, "blip") || blipSeen) return false;
                blipSeen = true;
                string? embed = GetRelationshipAttribute(reader, "embed");
                string? link = GetRelationshipAttribute(reader, "link");
                if (!string.IsNullOrEmpty(embed)) {
                    shape.Image = context.ResolveImage(embed!);
                } else if (!string.IsNullOrEmpty(link)) {
                    shape.Image = context.ResolveImage(link!);
                }
                return false;
            });
        }

        // Classifies a graphic frame the way the presentation model does: table, chart, SmartArt, then OLE object.
        private static void ReadGraphic(XmlReader reader, SlideContext context, ShapeBuilder shape) {
            bool dataSeen = false;
            ForEachChild(reader, () => {
                if (!IsDrawing(reader, "graphicData") || dataSeen) return false;
                dataSeen = true;
                bool chart = false, smartArt = false, ole = false;
                ForEachChild(reader, () => {
                    if (IsDrawing(reader, "tbl")) {
                        if (shape.Table != null) return false;
                        shape.Table = ReadTable(reader, context, shape);
                        return true;
                    }
                    if (IsChartNamespace(reader.NamespaceURI) && reader.LocalName == "chart") chart = true;
                    else if (IsDiagramNamespace(reader.NamespaceURI) && reader.LocalName == "relIds") smartArt = true;
                    else if (IsPresentation(reader, "oleObj")) ole = true;
                    return false;
                });
                shape.FrameKind = shape.Table != null ? PowerPointShapeContentType.Table
                    : chart ? PowerPointShapeContentType.Chart
                    : smartArt ? PowerPointShapeContentType.SmartArt
                    : ole ? PowerPointShapeContentType.OleObject
                    : null;
                return true;
            });
        }

        private static PowerPointTableContent ReadTable(XmlReader reader, SlideContext context, ShapeBuilder shape) {
            bool propertiesSeen = false, gridSeen = false, headerRow = false;
            int columnCount = 0;
            var rows = new List<PowerPointTableRowContent>();
            ForEachChild(reader, () => {
                if (IsDrawing(reader, "tblPr") && !propertiesSeen) {
                    propertiesSeen = true;
                    headerRow = ReadBoolean(reader.GetAttribute("firstRow")) == true;
                    return false;
                }
                if (IsDrawing(reader, "tblGrid") && !gridSeen) {
                    gridSeen = true;
                    ForEachChild(reader, () => {
                        if (IsDrawing(reader, "gridCol")) columnCount++;
                        return false;
                    });
                    return true;
                }
                if (IsDrawing(reader, "tr")) {
                    context.CancellationToken.ThrowIfCancellationRequested();
                    var cells = new List<PowerPointTableCellContent>();
                    ForEachChild(reader, () => {
                        if (!IsDrawing(reader, "tc")) return false;
                        IReadOnlyList<PowerPointParagraphContent>? paragraphs = null;
                        ForEachChild(reader, () => {
                            if (!IsDrawing(reader, "txBody") || paragraphs != null) return false;
                            paragraphs = ReadTextBody(reader, context);
                            return true;
                        });
                        cells.Add(new PowerPointTableCellContent(paragraphs ?? Array.Empty<PowerPointParagraphContent>()));
                        return true;
                    });
                    rows.Add(new PowerPointTableRowContent(cells));
                    return true;
                }
                return false;
            });
            shape.TableHasGrid = gridSeen;
            return new PowerPointTableContent(null, columnCount, headerRow, rows);
        }

        private static IReadOnlyList<PowerPointParagraphContent> ReadTextBody(XmlReader reader, SlideContext context) {
            var paragraphs = new List<PowerPointParagraphContent>();
            ForEachChild(reader, () => {
                if (!IsDrawing(reader, "p")) return false;
                paragraphs.Add(ReadParagraph(reader, context));
                return true;
            });
            return paragraphs;
        }

        private static PowerPointParagraphContent ReadParagraph(XmlReader reader, SlideContext context) {
            var inlines = new List<PowerPointInlineContent>();
            bool propertiesSeen = false, bulletSeen = false, numberingSeen = false;
            int? level = null, startAt = null;
            string? bullet = null;
            bool numbered = false;
            PowerPointNumberingScheme? scheme = null;
            ForEachChild(reader, () => {
                if (!IsDrawingNamespace(reader.NamespaceURI)) return false;
                switch (reader.LocalName) {
                    case "pPr":
                        if (propertiesSeen) return false;
                        propertiesSeen = true;
                        string? levelValue = reader.GetAttribute("lvl");
                        level = levelValue == null ? null : ParseInt32(levelValue);
                        ForEachChild(reader, () => {
                            if (IsDrawing(reader, "buChar") && !bulletSeen) {
                                bulletSeen = true;
                                bullet = reader.GetAttribute("char");
                            } else if (IsDrawing(reader, "buAutoNum") && !numberingSeen) {
                                numberingSeen = true;
                                numbered = true;
                                string? type = reader.GetAttribute("type");
                                scheme = type == null ? null : ParseNumberingScheme(type);
                                string? start = reader.GetAttribute("startAt");
                                startAt = start == null ? null : ParseInt32(start);
                            }
                            return false;
                        });
                        return true;
                    case "r":
                        inlines.Add(ReadRun(reader, context));
                        return true;
                    case "br":
                        inlines.Add(PowerPointInlineContent.LineBreak);
                        return false;
                    case "fld":
                        inlines.Add(ReadField(reader));
                        return true;
                    default:
                        return false;
                }
            });
            return new PowerPointParagraphContent(inlines, level, bullet, numbered, scheme, startAt);
        }

        private static PowerPointInlineContent ReadRun(XmlReader reader, SlideContext context) {
            string? text = null;
            bool propertiesSeen = false;
            Uri? hyperlink = null;
            ForEachChild(reader, () => {
                if (IsDrawing(reader, "rPr") && !propertiesSeen) {
                    propertiesSeen = true;
                    bool linkSeen = false;
                    ForEachChild(reader, () => {
                        if (!IsDrawing(reader, "hlinkClick") || linkSeen) return false;
                        linkSeen = true;
                        hyperlink = context.ResolveHyperlink(reader.GetAttribute("action"), GetRelationshipAttribute(reader, "id"));
                        return false;
                    });
                    return true;
                }
                if (IsDrawing(reader, "t") && text == null) {
                    text = ReadTextContent(reader);
                    return true;
                }
                return false;
            });
            return new PowerPointInlineContent(PowerPointParagraphInlineKind.Run, text ?? string.Empty, hyperlink, null);
        }

        private static PowerPointInlineContent ReadField(XmlReader reader) {
            string? text = null;
            StringBuilder? other = null;
            ForEachChild(reader, () => {
                if (IsDrawing(reader, "t") && text == null) {
                    text = ReadTextContent(reader);
                    return true;
                }
                string content = ReadTextContent(reader);
                if (content.Length > 0) (other ??= new StringBuilder()).Append(content);
                return true;
            });
            return new PowerPointInlineContent(PowerPointParagraphInlineKind.Field, text ?? string.Empty, null,
                text != null ? null : other?.ToString() ?? string.Empty);
        }

        // Matches PowerPointNotes.TryGetExistingText: the text of every top-level notes shape, blank ones skipped.
        private static string? ReadNotes(Package package, PackagePart slidePart, XmlReaderSettings settings) {
            try {
                PackagePart? notesPart = null;
                foreach (PackageRelationship relationship in slidePart.GetRelationships()) {
                    if (IsRelationshipType(relationship, "notesSlide") && (notesPart = GetTargetPart(package, relationship)) != null) break;
                }
                if (notesPart == null) return null;
                var shapes = new List<string>();
                bool treeFound = false;
                using (Stream stream = notesPart.GetStream(FileMode.Open, FileAccess.Read))
                using (XmlReader reader = XmlReader.Create(stream, settings)) {
                    if (reader.MoveToContent() != XmlNodeType.Element) return null;
                    bool commonSlideDataSeen = false;
                    ForEachChild(reader, () => {
                        if (commonSlideDataSeen || !IsPresentation(reader, "cSld")) return false;
                        commonSlideDataSeen = true;
                        ForEachChild(reader, () => {
                            if (treeFound || !IsPresentation(reader, "spTree")) return false;
                            treeFound = true;
                            ForEachChild(reader, () => {
                                if (!IsPresentation(reader, "sp")) return false;
                                string text = ReadNotesShapeText(reader);
                                if (!string.IsNullOrWhiteSpace(text)) shapes.Add(text);
                                return true;
                            });
                            return true;
                        });
                        return true;
                    });
                }
                if (!treeFound) return null;
                string notes = string.Join(Environment.NewLine + Environment.NewLine, shapes).Trim();
                return notes.Length > 0 ? notes : null;
            } catch (Exception exception) when (exception is not OutOfMemoryException) {
                return null;
            }
        }

        private static string ReadNotesShapeText(XmlReader reader) {
            var paragraphs = new List<string>();
            bool bodySeen = false;
            ForEachChild(reader, () => {
                if (!IsPresentation(reader, "txBody") || bodySeen) return false;
                bodySeen = true;
                ForEachChild(reader, () => {
                    if (!IsDrawing(reader, "p")) return false;
                    var text = new StringBuilder();
                    ForEachChild(reader, () => {
                        if (IsDrawing(reader, "br")) {
                            text.AppendLine();
                            return false;
                        }
                        if (!IsDrawing(reader, "r") && !IsDrawing(reader, "fld")) return false;
                        string? value = null;
                        ForEachChild(reader, () => {
                            if (!IsDrawing(reader, "t") || value != null) return false;
                            value = ReadTextContent(reader);
                            return true;
                        });
                        text.Append(value);
                        return true;
                    });
                    string paragraph = text.ToString();
                    if (!string.IsNullOrWhiteSpace(paragraph)) paragraphs.Add(paragraph);
                    return true;
                });
                return true;
            });
            return string.Join(Environment.NewLine, paragraphs);
        }

        /// <summary>
        ///     Visits the direct child elements of the element the reader is positioned on. The visitor returns true when
        ///     it consumed the child through its end; otherwise the child is skipped. Returns with the reader on the
        ///     parent's end (or on the parent itself when it is empty).
        /// </summary>
        private static void ForEachChild(XmlReader reader, Func<bool> visit) =>
            ForEachChild(reader, visit, null, null);

        private static void ForEachChild(XmlReader reader, Func<bool> visit, Func<XmlReader, bool>? isNested, Action? readNested) {
            if (reader.IsEmptyElement) return;
            int depth = reader.Depth;
            while (reader.Read()) {
                if (reader.NodeType == XmlNodeType.EndElement && reader.Depth == depth) return;
                if (reader.NodeType != XmlNodeType.Element) continue;
                if (isNested != null && isNested(reader)) {
                    readNested!();
                    continue;
                }
                if (!visit()) SkipElement(reader);
            }
        }

        // Visits every descendant element; the reader ends on the element's end.
        private static void ReadDescendants(XmlReader reader, Action visit) {
            if (reader.IsEmptyElement) return;
            int depth = reader.Depth;
            while (reader.Read()) {
                if (reader.NodeType == XmlNodeType.EndElement && reader.Depth == depth) return;
                if (reader.NodeType == XmlNodeType.Element) visit();
            }
        }

        private static void SkipElement(XmlReader reader) {
            if (reader.IsEmptyElement) return;
            int depth = reader.Depth;
            while (reader.Read() && !(reader.NodeType == XmlNodeType.EndElement && reader.Depth == depth)) { }
        }

        // Concatenates the element's descendant text; the reader ends on the element's end.
        private static string ReadTextContent(XmlReader reader) {
            if (reader.IsEmptyElement) return string.Empty;
            int depth = reader.Depth;
            string? single = null;
            StringBuilder? builder = null;
            while (reader.Read()) {
                if (reader.NodeType == XmlNodeType.EndElement && reader.Depth == depth) break;
                if (reader.NodeType is XmlNodeType.Text or XmlNodeType.CDATA or XmlNodeType.Whitespace or XmlNodeType.SignificantWhitespace) {
                    if (single == null && builder == null) {
                        single = reader.Value;
                    } else {
                        builder ??= new StringBuilder(single);
                        builder.Append(reader.Value);
                    }
                }
            }
            return builder?.ToString() ?? single ?? string.Empty;
        }

        private static string? GetRelationshipAttribute(XmlReader reader, string localName) {
            string? value = reader.GetAttribute(localName, RelationshipNamespace);
            return string.IsNullOrEmpty(value) ? reader.GetAttribute(localName, StrictRelationshipNamespace) : value;
        }

        private static bool? ReadBoolean(string? value) {
            if (value == null) return null;
            try {
                return XmlConvert.ToBoolean(value);
            } catch (FormatException exception) {
                throw new InvalidDataException($"'{value}' is not a valid boolean value.", exception);
            }
        }

        private static uint ParseUInt32(string value) {
            try {
                return XmlConvert.ToUInt32(value);
            } catch (Exception exception) when (exception is FormatException or OverflowException) {
                throw new InvalidDataException($"'{value}' is not a valid shape identifier.", exception);
            }
        }

        private static int ParseInt32(string value) {
            try {
                return XmlConvert.ToInt32(value);
            } catch (Exception exception) when (exception is FormatException or OverflowException) {
                throw new InvalidDataException($"'{value}' is not a valid integer value.", exception);
            }
        }

        private static PowerPointPlaceholderType ParsePlaceholderType(string value) {
            var parsed = new P.PlaceholderValues(value);
            if (!((IEnumValue)parsed).IsValid) throw new InvalidDataException($"'{value}' is not a valid placeholder type.");
            return parsed.ToOfficeEnum();
        }

        private static PowerPointNumberingScheme ParseNumberingScheme(string value) {
            var parsed = new A.TextAutoNumberSchemeValues(value);
            if (!((IEnumValue)parsed).IsValid) throw new InvalidDataException($"'{value}' is not a valid numbering scheme.");
            return parsed.ToOfficeEnum();
        }

        private const string RelationshipTypeBase = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/";
        private const string StrictRelationshipTypeBase = "http://purl.oclc.org/ooxml/officeDocument/relationships/";
        private const string SlideJumpAction = "ppaction://hlinksldjump";
        private const string NextSlideAction = "ppaction://hlinkshowjump?jump=nextslide";

        private static PackagePart GetPresentationPart(Package package) {
            foreach (PackageRelationship relationship in package.GetRelationships()) {
                if (IsRelationshipType(relationship, "officeDocument") && GetTargetPart(package, relationship) is { } part) return part;
            }
            throw new InvalidDataException("The package has no presentation part.");
        }

        private static bool IsRelationshipType(PackageRelationship relationship, string name) {
            string type = relationship.RelationshipType;
            return type == RelationshipTypeBase + name || type == StrictRelationshipTypeBase + name;
        }

        private static PackagePart? GetTargetPart(Package package, PackageRelationship relationship) {
            if (relationship.TargetMode == TargetMode.External) return null;
            Uri target = PackUriHelper.ResolvePartUri(relationship.SourceUri, relationship.TargetUri);
            return package.PartExists(target) ? package.GetPart(target) : null;
        }

        private static bool IsPresentation(XmlReader reader, string localName) =>
            reader.LocalName == localName && IsPresentationNamespace(reader.NamespaceURI);

        private static bool IsDrawing(XmlReader reader, string localName) =>
            reader.LocalName == localName && IsDrawingNamespace(reader.NamespaceURI);

        private static bool IsPresentationNamespace(string value) =>
            value == PresentationNamespace || value == StrictPresentationNamespace;

        private static bool IsDrawingNamespace(string value) =>
            value == DrawingNamespace || value == StrictDrawingNamespace;

        private static bool IsChartNamespace(string value) => value == ChartNamespace || value == StrictChartNamespace;

        private static bool IsDiagramNamespace(string value) => value == DiagramNamespace || value == StrictDiagramNamespace;

        private sealed class SlideContext {
            private readonly PowerPointSlideReader _reader;
            private readonly PackagePart _slidePart;
            private readonly List<ShapeBuilder?> _shapes = new();
            private int _count;

            internal SlideContext(PowerPointSlideReader reader, PackagePart slidePart, CancellationToken cancellationToken) {
                _reader = reader;
                _slidePart = slidePart;
                CancellationToken = cancellationToken;
            }

            internal CancellationToken CancellationToken { get; }

            internal int Reserve(ShapeBuilder shape) {
                _shapes.Add(shape);
                return _shapes.Count - 1;
            }

            // A frame that turned out not to be a shape leaves its slot, and any children it recorded, unused.
            internal void Release(int position) {
                for (int index = position; index < _shapes.Count; index++) _shapes[index] = null;
            }

            internal void Count() {
                if (++_count > MaximumShapeCount) {
                    throw new InvalidOperationException("The shape count exceeds the configured inspection limit.");
                }
            }

            private PackageRelationship? GetRelationship(string relationshipId) =>
                _slidePart.RelationshipExists(relationshipId) ? _slidePart.GetRelationship(relationshipId) : null;

            internal PowerPointImageReference ResolveImage(string relationshipId) {
                PackageRelationship? relationship = GetRelationship(relationshipId);
                if (relationship?.TargetMode == TargetMode.External) {
                    return new PowerPointImageReference(relationshipId, null, null, relationship.TargetUri);
                }
                PackagePart? part = relationship == null ? null : GetTargetPart(_reader._package, relationship);
                return new PowerPointImageReference(relationshipId, part?.Uri, part?.ContentType, null);
            }

            // Resolves click hyperlinks as PowerPointHyperlinkResolver does for a loaded slide: external hyperlink
            // targets as they are, and slide jumps and next-slide actions as #slide-N fragments.
            internal Uri? ResolveHyperlink(string? action, string? relationshipId) {
                if (string.Equals(action, NextSlideAction, StringComparison.OrdinalIgnoreCase)) {
                    return _reader._slideNumbers.TryGetValue(_slidePart.Uri.OriginalString, out int position)
                        ? SlideFragment(position + 1)
                        : null;
                }
                if (string.IsNullOrWhiteSpace(relationshipId)) return null;
                PackageRelationship? relationship = GetRelationship(relationshipId!);
                bool hyperlink = relationship != null && IsRelationshipType(relationship, "hyperlink");
                if (hyperlink && relationship!.TargetMode == TargetMode.External) return relationship.TargetUri;
                if (!string.Equals(action, SlideJumpAction, StringComparison.OrdinalIgnoreCase)
                    || relationship == null || relationship.TargetMode == TargetMode.External) {
                    return null;
                }
                Uri target = PackUriHelper.ResolvePartUri(relationship.SourceUri, relationship.TargetUri);
                if (!hyperlink && !_reader._package.PartExists(target)) return null;
                return _reader._slideNumbers.TryGetValue(target.OriginalString, out int number) ? SlideFragment(number) : null;
            }

            private static Uri SlideFragment(int number) =>
                new("#slide-" + number.ToString(CultureInfo.InvariantCulture), UriKind.Relative);

            internal IReadOnlyList<PowerPointShapeContent> Build() {
                var shapes = new List<PowerPointShapeContent>(_shapes.Count);
                foreach (ShapeBuilder? shape in _shapes) {
                    if (shape != null) shapes.Add(shape.Build());
                }
                return shapes;
            }
        }

        private sealed class ShapeBuilder {
            private PowerPointShapeContent? _built;

            internal ShapeBuilder(ShapeBuilder? parent, int depth) {
                Parent = parent;
                Depth = depth;
            }

            internal ShapeBuilder? Parent { get; }
            internal int Depth { get; }
            internal PowerPointShapeContentType Kind { get; set; }
            internal uint? Id { get; set; }
            internal string? Name { get; set; }
            internal string? AltText { get; set; }
            internal bool Hidden { get; set; }
            internal PowerPointPlaceholderType? PlaceholderType { get; set; }
            internal IReadOnlyList<PowerPointParagraphContent>? Paragraphs { get; set; }
            internal PowerPointTableContent? Table { get; set; }
            internal bool TableHasGrid { get; set; }
            internal PowerPointImageReference? Image { get; set; }
            internal bool IsMedia { get; set; }
            internal bool NonVisualSeen { get; set; }
            internal bool GraphicSeen { get; set; }
            internal PowerPointShapeContentType? FrameKind { get; set; }

            internal bool IsVisible => !Hidden && (Parent == null || Parent.IsVisible);

            internal PowerPointShapeContent Build() {
                if (_built != null) return _built;
                PowerPointTableContent? table = Table == null ? null
                    : new PowerPointTableContent(Name, Table.ColumnCount, Table.HeaderRow, Table.Rows, TableHasGrid);
                _built = new PowerPointShapeContent(Kind, Id, Name, AltText, Hidden, IsVisible, Depth, PlaceholderType,
                    Kind == PowerPointShapeContentType.TextBox ? Paragraphs! : Array.Empty<PowerPointParagraphContent>(),
                    table, Image);
                return _built;
            }
        }
    }
}

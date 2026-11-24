using System.Globalization;
using System.Text;
using System.Xml;
using OfficeIMO.Core.Internal.OpenXml;

namespace OfficeIMO.PowerPoint {
    public sealed partial class PowerPointSlideReader {
        // The same shape and nesting limits PowerPointSlide.EnumerateShapesDeep applies to a loaded slide.
        internal const int MaximumShapeCount = 10000;
        internal const int MaximumGroupDepth = 128;

        // Strict namespaces read as their transitional equivalents, so one id covers both spellings.
        private const int PresentationNamespace = OfficeXmlNamespaces.PresentationML;
        private const int DrawingNamespace = OfficeXmlNamespaces.DrawingML;
        private const int RelationshipNamespace = OfficeXmlNamespaces.Relationships;

        private static IReadOnlyList<SlideEntry> ReadSlideList(OfficeOpcPackage package, string presentationPart,
            long maximumCharacters, CancellationToken cancellationToken) {
            var slides = new List<SlideEntry>();
            OfficeXmlPullReader reader = OpenPart(package, presentationPart, maximumCharacters, out OfficePooledBytes content);
            using OfficePooledBytes owned = content;
            if (!MoveToRoot(reader) || !IsPresentation(reader, "presentation")) {
                // A model load rejects a package whose main part is another document, such as a renamed DOCX.
                throw new InvalidDataException("The package is not a PowerPoint presentation.");
            }
            bool listSeen = false;
            for (var children29 = new Children(reader); children29.Next(reader);) {
                if (listSeen || !IsPresentation(reader, "sldIdLst")) continue;
                listSeen = true;
                for (var children28 = new Children(reader); children28.Next(reader);) {
                    if (!IsPresentation(reader, "sldId")) continue;
                    cancellationToken.ThrowIfCancellationRequested();
                    string? relationshipId = GetRelationshipAttribute(reader, "id");
                    if (string.IsNullOrEmpty(relationshipId)) continue;
                    OfficeOpcRelationship? relationship = FindRelationship(package, presentationPart, relationshipId!);
                    string? slidePart = relationship != null && IsRelationshipType(relationship, "slide")
                        ? GetTargetPart(package, relationship)
                        : null;
                    if (slidePart == null) {
                        throw new InvalidDataException($"The slide list entry '{relationshipId}' does not target a slide part.");
                    }
                    slides.Add(new SlideEntry(slidePart, GetAttribute(reader, "show")));
                    continue;
                }
                continue;
            }
            return slides;
        }

        private PowerPointSlideContent ReadSlide(int index, bool includeNotes, CancellationToken cancellationToken) {
            SlideEntry entry = _slides[index];
            var context = new SlideContext(this, entry.PartName, cancellationToken);
            bool? show = null;
            OfficeXmlPullReader reader = OpenPart(_package, entry.PartName, _maximumCharacters, out OfficePooledBytes content);
            using (content) {
                if (MoveToRoot(reader)) {
                    show = ReadBoolean(GetAttribute(reader, "show"));
                    bool commonSlideDataSeen = false;
                    for (var children27 = new Children(reader); children27.Next(reader);) {
                        if (commonSlideDataSeen || !IsPresentation(reader, "cSld")) continue;
                        commonSlideDataSeen = true;
                        bool treeSeen = false;
                        for (var children26 = new Children(reader); children26.Next(reader);) {
                            if (treeSeen || !IsPresentation(reader, "spTree")) continue;
                            treeSeen = true;
                            ReadShapeTree(reader, context, parent: null, depth: 0);
                            continue;
                        }
                        continue;
                    }
                }
            }

            bool hidden = show.HasValue ? !show.Value : IsHiddenShowValue(entry.ListShowValue);
            string? notes = includeNotes ? ReadNotes(_package, entry.PartName, _maximumCharacters, cancellationToken) : null;
            cancellationToken.ThrowIfCancellationRequested();
            return new PowerPointSlideContent(index + 1, hidden, context.Build(), notes);
        }

        private static bool IsHiddenShowValue(string? value) =>
            string.Equals(value, "0", StringComparison.Ordinal) || string.Equals(value, "false", StringComparison.OrdinalIgnoreCase);

        // Reads the children of a shape tree or group in drawing order; the caller is positioned on its start element.
        private static void ReadShapeTree(OfficeXmlPullReader reader, SlideContext context, ShapeBuilder? parent, int depth) {
            for (var children25 = new Children(reader); children25.Next(reader);) {
                if (reader.NamespaceId != PresentationNamespace) continue;
                switch (reader.LocalName) {
                    case "sp":
                    case "cxnSp":
                    case "pic":
                    case "grpSp":
                    case "graphicFrame":
                        ReadShape(reader, context, parent, depth);
                        continue;
                    default:
                        continue;
                }
            }
        }

        private static void ReadShape(OfficeXmlPullReader reader, SlideContext context, ShapeBuilder? parent, int depth) {
            context.CancellationToken.ThrowIfCancellationRequested();
            if (depth > MaximumGroupDepth) {
                throw new InvalidOperationException("The grouped-shape nesting depth exceeds the configured limit.");
            }
            string element = reader.LocalName;
            var shape = new ShapeBuilder(parent, depth);
            int position = context.Reserve(shape);
            switch (element) {
                case "sp":
                    for (var children24 = new Children(reader); children24.Next(reader);) {
                        if (IsPresentation(reader, "nvSpPr")) { _ = ReadNonVisualProperties(reader, shape, readPlaceholder: true); continue; }
                        if (IsPresentation(reader, "txBody") && shape.Paragraphs == null) {
                            shape.Paragraphs = ReadTextBody(reader, context);
                            continue;
                        }
                        continue;
                    }
                    shape.Kind = shape.Paragraphs != null ? PowerPointShapeContentType.TextBox : PowerPointShapeContentType.AutoShape;
                    break;
                case "cxnSp":
                    for (var connectorChildren = new Children(reader); connectorChildren.Next(reader);) {
                        if (IsPresentation(reader, "nvCxnSpPr")) ReadNonVisualProperties(reader, shape, readPlaceholder: true);
                    }
                    shape.Kind = PowerPointShapeContentType.Connector;
                    break;
                case "pic":
                    bool fillSeen = false;
                    for (var children23 = new Children(reader); children23.Next(reader);) {
                        if (IsPresentation(reader, "nvPicPr")) { _ = ReadNonVisualProperties(reader, shape, readPlaceholder: true, detectMedia: true); continue; }
                        if (IsPresentation(reader, "blipFill") && !fillSeen) {
                            fillSeen = true;
                            ReadBlipFill(reader, context, shape);
                            continue;
                        }
                        continue;
                    }
                    shape.Kind = shape.IsMedia ? PowerPointShapeContentType.Media : PowerPointShapeContentType.Picture;
                    break;
                case "grpSp":
                    shape.Kind = PowerPointShapeContentType.Group;
                    for (var groupChildren = new Children(reader); groupChildren.Next(reader);) {
                        if (reader.NamespaceId == PresentationNamespace && reader.LocalName is "sp" or "cxnSp" or "pic" or "grpSp" or "graphicFrame") {
                            ReadShape(reader, context, shape, depth + 1);
                        } else if (IsPresentation(reader, "nvGrpSpPr")) {
                            ReadNonVisualProperties(reader, shape, readPlaceholder: false);
                        }
                    }
                    break;
                case "graphicFrame":
                    for (var children22 = new Children(reader); children22.Next(reader);) {
                        if (IsPresentation(reader, "nvGraphicFramePr")) { _ = ReadNonVisualProperties(reader, shape, readPlaceholder: true); continue; }
                        if (IsDrawing(reader, "graphic") && !shape.GraphicSeen) {
                            shape.GraphicSeen = true;
                            ReadGraphic(reader, context, shape);
                            continue;
                        }
                        continue;
                    }
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
        private static bool ReadNonVisualProperties(OfficeXmlPullReader reader, ShapeBuilder shape, bool readPlaceholder,
            bool detectMedia = false) {
            if (shape.NonVisualSeen) return false;
            shape.NonVisualSeen = true;
            bool drawingSeen = false, applicationSeen = false;
            for (var children21 = new Children(reader); children21.Next(reader);) {
                if (IsPresentation(reader, "cNvPr") && !drawingSeen) {
                    drawingSeen = true;
                    string? id = GetAttribute(reader, "id");
                    shape.Id = id == null ? null : ParseUInt32(id);
                    shape.Name = GetAttribute(reader, "name");
                    shape.AltText = GetAttribute(reader, "descr");
                    shape.Hidden = ReadBoolean(GetAttribute(reader, "hidden")) == true;
                    continue;
                }
                if (IsPresentation(reader, "nvPr") && !applicationSeen) {
                    applicationSeen = true;
                    bool placeholderSeen = false;
                    int propertiesDepth = reader.Depth;
                    bool propertiesEmpty = reader.IsEmptyElement;
                    while (!propertiesEmpty && reader.Read()
                           && !(reader.NodeKind == OfficeXmlNodeKind.EndElement && reader.Depth == propertiesDepth)) {
                        if (reader.NodeKind != OfficeXmlNodeKind.StartElement) continue;
                        if (readPlaceholder && !placeholderSeen && reader.Depth == propertiesDepth + 1 && IsPresentation(reader, "ph")) {
                            placeholderSeen = true;
                            string? type = GetAttribute(reader, "type");
                            if (type != null) shape.PlaceholderType = ParsePlaceholderType(type);
                        }
                        if (detectMedia && reader.NamespaceId == DrawingNamespace
                            && reader.LocalName is "audioFile" or "videoFile") {
                            shape.IsMedia = true;
                        }
                    }
                    continue;
                }
                continue;
            }
            return true;
        }

        private static void ReadBlipFill(OfficeXmlPullReader reader, SlideContext context, ShapeBuilder shape) {
            bool blipSeen = false;
            for (var children20 = new Children(reader); children20.Next(reader);) {
                if (!IsDrawing(reader, "blip") || blipSeen) continue;
                blipSeen = true;
                string? embed = GetRelationshipAttribute(reader, "embed");
                string? link = GetRelationshipAttribute(reader, "link");
                if (!string.IsNullOrEmpty(embed)) {
                    shape.Image = context.ResolveImage(embed!);
                } else if (!string.IsNullOrEmpty(link)) {
                    shape.Image = context.ResolveImage(link!);
                }
                continue;
            }
        }

        // Classifies a graphic frame the way the presentation model does: table, chart, SmartArt, then OLE object.
        private static void ReadGraphic(OfficeXmlPullReader reader, SlideContext context, ShapeBuilder shape) {
            bool dataSeen = false;
            for (var children19 = new Children(reader); children19.Next(reader);) {
                if (!IsDrawing(reader, "graphicData") || dataSeen) continue;
                dataSeen = true;
                bool chart = false, smartArt = false, ole = false;
                for (var children18 = new Children(reader); children18.Next(reader);) {
                    if (IsDrawing(reader, "tbl")) {
                        if (shape.Table != null) continue;
                        shape.Table = ReadTable(reader, context, shape);
                        continue;
                    }
                    if (reader.IsElement(OfficeXmlNamespaces.Chart, "chart")) chart = true;
                    else if (reader.IsElement(OfficeXmlNamespaces.Diagram, "relIds")) smartArt = true;
                    else if (IsPresentation(reader, "oleObj")) ole = true;
                    continue;
                }
                shape.FrameKind = shape.Table != null ? PowerPointShapeContentType.Table
                    : chart ? PowerPointShapeContentType.Chart
                    : smartArt ? PowerPointShapeContentType.SmartArt
                    : ole ? PowerPointShapeContentType.OleObject
                    : null;
                continue;
            }
        }

        private static PowerPointTableContent ReadTable(OfficeXmlPullReader reader, SlideContext context, ShapeBuilder shape) {
            bool propertiesSeen = false, gridSeen = false, headerRow = false;
            int columnCount = 0;
            var rows = new List<PowerPointTableRowContent>();
            for (var children17 = new Children(reader); children17.Next(reader);) {
                if (IsDrawing(reader, "tblPr") && !propertiesSeen) {
                    propertiesSeen = true;
                    headerRow = ReadBoolean(GetAttribute(reader, "firstRow")) == true;
                    continue;
                }
                if (IsDrawing(reader, "tblGrid") && !gridSeen) {
                    gridSeen = true;
                    for (var children16 = new Children(reader); children16.Next(reader);) {
                        if (IsDrawing(reader, "gridCol")) columnCount++;
                        continue;
                    }
                    continue;
                }
                if (IsDrawing(reader, "tr")) {
                    context.CancellationToken.ThrowIfCancellationRequested();
                    var cells = new List<PowerPointTableCellContent>();
                    for (var children15 = new Children(reader); children15.Next(reader);) {
                        if (!IsDrawing(reader, "tc")) continue;
                        IReadOnlyList<PowerPointParagraphContent>? paragraphs = null;
                        for (var children14 = new Children(reader); children14.Next(reader);) {
                            if (!IsDrawing(reader, "txBody") || paragraphs != null) continue;
                            paragraphs = ReadTextBody(reader, context);
                            continue;
                        }
                        cells.Add(new PowerPointTableCellContent(paragraphs ?? Array.Empty<PowerPointParagraphContent>()));
                        continue;
                    }
                    rows.Add(new PowerPointTableRowContent(cells));
                    continue;
                }
                continue;
            }
            shape.TableHasGrid = gridSeen;
            return new PowerPointTableContent(null, columnCount, headerRow, rows);
        }

        private static IReadOnlyList<PowerPointParagraphContent> ReadTextBody(OfficeXmlPullReader reader, SlideContext context) {
            var paragraphs = new List<PowerPointParagraphContent>();
            for (var children13 = new Children(reader); children13.Next(reader);) {
                if (!IsDrawing(reader, "p")) continue;
                paragraphs.Add(ReadParagraph(reader, context));
                continue;
            }
            return paragraphs;
        }

        private static PowerPointParagraphContent ReadParagraph(OfficeXmlPullReader reader, SlideContext context) {
            var inlines = new List<PowerPointInlineContent>();
            bool propertiesSeen = false, bulletSeen = false, numberingSeen = false;
            int? level = null, startAt = null;
            string? bullet = null;
            bool numbered = false;
            PowerPointNumberingScheme? scheme = null;
            for (var children12 = new Children(reader); children12.Next(reader);) {
                if (reader.NamespaceId != DrawingNamespace) continue;
                switch (reader.LocalName) {
                    case "pPr":
                        if (propertiesSeen) continue;
                        propertiesSeen = true;
                        string? levelValue = GetAttribute(reader, "lvl");
                        level = levelValue == null ? null : ParseInt32(levelValue);
                        for (var children11 = new Children(reader); children11.Next(reader);) {
                            if (IsDrawing(reader, "buChar") && !bulletSeen) {
                                bulletSeen = true;
                                bullet = GetAttribute(reader, "char");
                            } else if (IsDrawing(reader, "buAutoNum") && !numberingSeen) {
                                numberingSeen = true;
                                numbered = true;
                                string? type = GetAttribute(reader, "type");
                                scheme = ParseNumberingScheme(type);
                                string? start = GetAttribute(reader, "startAt");
                                startAt = start == null ? null : ParseInt32(start);
                            }
                            continue;
                        }
                        continue;
                    case "r":
                        inlines.Add(ReadRun(reader, context));
                        continue;
                    case "br":
                        inlines.Add(PowerPointInlineContent.LineBreak);
                        continue;
                    case "fld":
                        inlines.Add(ReadField(reader));
                        continue;
                    default:
                        continue;
                }
            }
            return new PowerPointParagraphContent(inlines, level, bullet, numbered, scheme, startAt);
        }

        private static PowerPointInlineContent ReadRun(OfficeXmlPullReader reader, SlideContext context) {
            string? text = null;
            bool propertiesSeen = false;
            Uri? hyperlink = null;
            for (var children10 = new Children(reader); children10.Next(reader);) {
                if (IsDrawing(reader, "rPr") && !propertiesSeen) {
                    propertiesSeen = true;
                    bool linkSeen = false;
                    for (var children9 = new Children(reader); children9.Next(reader);) {
                        if (!IsDrawing(reader, "hlinkClick") || linkSeen) continue;
                        linkSeen = true;
                        hyperlink = context.ResolveHyperlink(GetAttribute(reader, "action"), GetRelationshipAttribute(reader, "id"));
                        continue;
                    }
                    continue;
                }
                if (IsDrawing(reader, "t") && text == null) {
                    text = ReadTextContent(reader);
                    continue;
                }
                continue;
            }
            return new PowerPointInlineContent(PowerPointParagraphInlineKind.Run, text ?? string.Empty, hyperlink, null);
        }

        private static PowerPointInlineContent ReadField(OfficeXmlPullReader reader) {
            string? text = null;
            StringBuilder? other = null;
            for (var children8 = new Children(reader); children8.Next(reader);) {
                if (IsDrawing(reader, "t") && text == null) {
                    text = ReadTextContent(reader);
                    continue;
                }
                string content = ReadTextContent(reader);
                if (content.Length > 0) (other ??= new StringBuilder()).Append(content);
                continue;
            }
            return new PowerPointInlineContent(PowerPointParagraphInlineKind.Field, text ?? string.Empty, null,
                text != null ? null : other?.ToString() ?? string.Empty);
        }

        // Matches PowerPointNotes.TryGetExistingText: the text of every top-level notes shape, blank ones skipped.
        private static string? ReadNotes(OfficeOpcPackage package, string slidePart, long maximumCharacters,
            CancellationToken cancellationToken) {
            try {
                string? notesPart = null;
                foreach (OfficeOpcRelationship relationship in package.GetRelationships(slidePart)) {
                    if (IsRelationshipType(relationship, "notesSlide") && (notesPart = GetTargetPart(package, relationship)) != null) break;
                }
                if (notesPart == null) return null;
                var shapes = new List<string>();
                bool treeFound = false;
                OfficeXmlPullReader reader = OpenPart(package, notesPart, maximumCharacters, out OfficePooledBytes content);
                using (content) {
                    if (!MoveToRoot(reader)) return null;
                    bool commonSlideDataSeen = false;
                    for (var children7 = new Children(reader); children7.Next(reader);) {
                        if (commonSlideDataSeen || !IsPresentation(reader, "cSld")) continue;
                        commonSlideDataSeen = true;
                        for (var children6 = new Children(reader); children6.Next(reader);) {
                            if (treeFound || !IsPresentation(reader, "spTree")) continue;
                            treeFound = true;
                            for (var children5 = new Children(reader); children5.Next(reader);) {
                                if (!IsPresentation(reader, "sp")) continue;
                                cancellationToken.ThrowIfCancellationRequested();
                                string text = ReadNotesShapeText(reader);
                                if (!string.IsNullOrWhiteSpace(text)) shapes.Add(text);
                                continue;
                            }
                            continue;
                        }
                        continue;
                    }
                }
                if (!treeFound) return null;
                string notes = string.Join(Environment.NewLine + Environment.NewLine, shapes).Trim();
                return notes.Length > 0 ? notes : null;
            } catch (Exception exception) when (exception is not OutOfMemoryException and not OperationCanceledException
                                                  && !IsPartSizeLimit(exception)) {
                // Like the model, malformed notes count as no notes; a part over the size limit does not.
                return null;
            }
        }

        private const string PartSizeLimitKey = "OfficeIMO.PartSizeLimit";

        // A part over the character limit is reported as an XmlException, as XmlReader reports MaxCharactersInDocument.
        private static bool IsPartSizeLimit(Exception exception) => exception is XmlException && exception.Data.Contains(PartSizeLimitKey);

        private static XmlException PartSizeLimit(string partName, long maximumCharacters) {
            var exception = new XmlException($"The part '{partName}' has more than {maximumCharacters} characters, the configured maximum.");
            exception.Data[PartSizeLimitKey] = true;
            return exception;
        }

        // Comments and processing instructions are not reported, and compatibility markup is read as written: text in
        // markup-compatibility fallbacks is not read, as the remarks on the class say.
        private static OfficeXmlPullReader OpenPart(OfficeOpcPackage package, string partName, long maximumCharacters,
            out OfficePooledBytes content) {
            // A member that declares more bytes than the limit allows fails before it is inflated.
            if (package.GetPartLength(partName) > OfficeXmlPullReader.MaximumBytesFor(maximumCharacters)) {
                content = default;
                throw PartSizeLimit(partName, maximumCharacters);
            }
            content = package.ReadPart(partName);
            // A part can only exceed the limit when it has more bytes than the limit, as UTF-8 never encodes a
            // character in less than a byte; only then are its characters counted.
            if (content.Length > maximumCharacters && OfficeXmlPullReader.CountCharacters(content.Array, content.Length) > maximumCharacters) {
                content.Dispose();
                throw PartSizeLimit(partName, maximumCharacters);
            }
            return new OfficeXmlPullReader(content.Array, content.Length, package.Namespaces, processMarkupCompatibility: false);
        }

        private static bool MoveToRoot(OfficeXmlPullReader reader) {
            while (reader.Read()) {
                if (reader.NodeKind == OfficeXmlNodeKind.StartElement) return true;
            }
            return false;
        }

        private static string ReadNotesShapeText(OfficeXmlPullReader reader) {
            var paragraphs = new List<string>();
            bool bodySeen = false;
            for (var children4 = new Children(reader); children4.Next(reader);) {
                if (!IsPresentation(reader, "txBody") || bodySeen) continue;
                bodySeen = true;
                for (var children3 = new Children(reader); children3.Next(reader);) {
                    if (!IsDrawing(reader, "p")) continue;
                    var text = new StringBuilder();
                    for (var children2 = new Children(reader); children2.Next(reader);) {
                        if (IsDrawing(reader, "br")) {
                            text.AppendLine();
                            continue;
                        }
                        if (!IsDrawing(reader, "r") && !IsDrawing(reader, "fld")) continue;
                        string? value = null;
                        for (var children1 = new Children(reader); children1.Next(reader);) {
                            if (!IsDrawing(reader, "t") || value != null) continue;
                            value = ReadTextContent(reader);
                            continue;
                        }
                        text.Append(value);
                        continue;
                    }
                    string paragraph = text.ToString();
                    if (!string.IsNullOrWhiteSpace(paragraph)) paragraphs.Add(paragraph);
                    continue;
                }
                continue;
            }
            return string.Join(Environment.NewLine, paragraphs);
        }

        /// <summary>
        ///     Steps through the direct child elements of the element the reader is positioned on. A child the loop body
        ///     does not read is passed over; one it reads must be left on its end. Returns false on the parent's end.
        /// </summary>
        private struct Children {
            private readonly int _depth;
            private bool _done;

            internal Children(OfficeXmlPullReader reader) {
                _depth = reader.Depth;
                _done = reader.IsEmptyElement;
            }

            internal bool Next(OfficeXmlPullReader reader) {
                while (!_done && reader.Read()) {
                    if (reader.NodeKind == OfficeXmlNodeKind.EndElement && reader.Depth == _depth) break;
                    if (reader.NodeKind == OfficeXmlNodeKind.StartElement && reader.Depth == _depth + 1) return true;
                }
                _done = true;
                return false;
            }
        }

        // Concatenates the element's descendant text; the reader ends on the element's end.
        private static string ReadTextContent(OfficeXmlPullReader reader) {
            if (reader.IsEmptyElement) return string.Empty;
            int depth = reader.Depth;
            string? single = null;
            OfficeCharBuffer? builder = null;
            while (reader.Read()) {
                if (reader.NodeKind == OfficeXmlNodeKind.EndElement && reader.Depth == depth) break;
                if (reader.NodeKind != OfficeXmlNodeKind.Text) continue;
                if (single == null && builder == null) {
                    single = reader.GetText();
                } else {
                    if (builder == null) {
                        builder = new OfficeCharBuffer(Math.Max(256, single!.Length * 2));
                        foreach (char value in single) builder.Append(value);
                    }
                    reader.AppendText(builder);
                }
            }
            return builder?.ToString() ?? single ?? string.Empty;
        }

        private static string? GetAttribute(OfficeXmlPullReader reader, string localName) =>
            reader.TryGetAttribute(OfficeXmlNamespaces.None, localName, out string value) ? value : null;

        private static string? GetRelationshipAttribute(OfficeXmlPullReader reader, string localName) =>
            reader.TryGetAttribute(RelationshipNamespace, localName, out string value) ? value : null;

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

        private static PowerPointPlaceholderType ParsePlaceholderType(string value) =>
            PowerPointOpenXmlValueTypeExtensions.TryParseOpenXmlValue(value, out PowerPointPlaceholderType type)
                ? type
                : throw new InvalidDataException($"'{value}' is not a valid placeholder type.");

        // Like the Markdown renderer for an unset scheme, an unknown automatic-numbering type falls back to plain
        // numbers instead of failing the slide; both the streamed and the loaded projection read it this way.
        internal static PowerPointNumberingScheme? ParseNumberingScheme(string? value) =>
            value != null && PowerPointOpenXmlValueTypeExtensions.TryParseOpenXmlValue(value, out PowerPointNumberingScheme scheme)
                ? scheme
                : null;

        private const string RelationshipTypeBase = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/";
        private const string SlideJumpAction = "ppaction://hlinksldjump";
        private const string NextSlideAction = "ppaction://hlinkshowjump?jump=nextslide";

        private static string GetPresentationPart(OfficeOpcPackage package) {
            foreach (OfficeOpcRelationship relationship in package.GetRelationships(null)) {
                if (IsRelationshipType(relationship, "officeDocument") && GetTargetPart(package, relationship) is { } part) {
                    package.RequireContentType(part,
                        "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml",
                        "application/vnd.openxmlformats-officedocument.presentationml.template.main+xml",
                        "application/vnd.openxmlformats-officedocument.presentationml.slideshow.main+xml",
                        "application/vnd.ms-powerpoint.presentation.macroEnabled.main+xml",
                        "application/vnd.ms-powerpoint.template.macroEnabled.main+xml",
                        "application/vnd.ms-powerpoint.slideshow.macroEnabled.main+xml",
                        "application/vnd.ms-powerpoint.addin.macroEnabled.main+xml");
                    return part;
                }
            }
            throw new InvalidDataException("The package has no presentation part.");
        }

        // Relationship types are reported with their transitional spelling, Strict ones included.
        private static bool IsRelationshipType(OfficeOpcRelationship relationship, string name) =>
            relationship.Type.Length == RelationshipTypeBase.Length + name.Length
            && relationship.Type.StartsWith(RelationshipTypeBase, StringComparison.Ordinal)
            && relationship.Type.EndsWith(name, StringComparison.Ordinal);

        private static string? GetTargetPart(OfficeOpcPackage package, OfficeOpcRelationship relationship) =>
            relationship.TargetPartName != null && package.PartExists(relationship.TargetPartName) ? relationship.TargetPartName : null;

        private static OfficeOpcRelationship? FindRelationship(OfficeOpcPackage package, string sourcePart, string relationshipId) =>
            package.FindRelationship(sourcePart, relationshipId);

        private static Uri GetTargetUri(OfficeOpcRelationship relationship) => new(relationship.Target, UriKind.RelativeOrAbsolute);

        private static bool IsPresentation(OfficeXmlPullReader reader, string localName) =>
            reader.IsElement(PresentationNamespace, localName);

        private static bool IsDrawing(OfficeXmlPullReader reader, string localName) =>
            reader.IsElement(DrawingNamespace, localName);

        private sealed class SlideContext {
            private readonly PowerPointSlideReader _reader;
            private readonly string _slidePart;
            private readonly List<ShapeBuilder?> _shapes = new();
            private int _count;

            internal SlideContext(PowerPointSlideReader reader, string slidePart, CancellationToken cancellationToken) {
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

            private OfficeOpcRelationship? GetRelationship(string relationshipId) =>
                FindRelationship(_reader._package, _slidePart, relationshipId);

            internal PowerPointImageReference ResolveImage(string relationshipId) {
                OfficeOpcRelationship? relationship = GetRelationship(relationshipId);
                if (relationship?.IsExternal == true) {
                    return new PowerPointImageReference(relationshipId, null, null, GetTargetUri(relationship));
                }
                string? part = relationship == null ? null : GetTargetPart(_reader._package, relationship);
                return part == null
                    ? new PowerPointImageReference(relationshipId, null, null, null)
                    : new PowerPointImageReference(relationshipId, new Uri(part, UriKind.Relative), _reader._package.GetContentType(part), null);
            }

            // Resolves click hyperlinks as PowerPointHyperlinkResolver does for a loaded slide: external hyperlink
            // targets as they are, and slide jumps and next-slide actions as #slide-N fragments.
            internal Uri? ResolveHyperlink(string? action, string? relationshipId) {
                if (string.Equals(action, NextSlideAction, StringComparison.OrdinalIgnoreCase)) {
                    return _reader._slideNumbers.TryGetValue(OfficeOpcPackage.GetPartIdentity(_slidePart), out int position)
                        ? SlideFragment(position + 1)
                        : null;
                }
                if (string.IsNullOrWhiteSpace(relationshipId)) return null;
                OfficeOpcRelationship? relationship = GetRelationship(relationshipId!);
                bool hyperlink = relationship != null && IsRelationshipType(relationship, "hyperlink");
                if (hyperlink && relationship!.IsExternal) return GetTargetUri(relationship);
                if (!string.Equals(action, SlideJumpAction, StringComparison.OrdinalIgnoreCase)
                    || relationship == null || relationship.IsExternal || relationship.TargetPartName == null) {
                    return null;
                }
                string target = relationship.TargetPartName;
                if (!hyperlink && !_reader._package.PartExists(target)) return null;
                return _reader._slideNumbers.TryGetValue(OfficeOpcPackage.GetPartIdentity(target), out int number) ? SlideFragment(number) : null;
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

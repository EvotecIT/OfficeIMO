using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Text;
using System.Xml.Linq;

namespace OfficeIMO.Excel.LegacyXls.Write {
    internal static class LegacyXlsCommentWriter {
        private const int BiffMaxRecordDataLength = 8224;
        private const int CommentShapeContainerPayloadLength = 126;
        private const int CommentShapeContainerLength = 8 + CommentShapeContainerPayloadLength;
        private const int MaximumCommentsPerDrawing = 1023;
        private const int MaximumDrawingIdentifier = 0x0ffe;

        internal static bool SupportsCommentCount(int commentCount) =>
            commentCount >= 0 && commentCount <= MaximumCommentsPerDrawing;

        internal static bool SupportsCommentDrawingSheetIndex(int zeroBasedSheetIndex) =>
            zeroBasedSheetIndex >= 0 && zeroBasedSheetIndex < MaximumDrawingIdentifier;

        internal static bool SupportsWorksheetComments(ExcelSheet sheet, LegacyXlsFontTable fontTable, out string? reason) {
            reason = null;
            if (sheet.WorksheetPart.WorksheetThreadedCommentsParts.Any()) {
                reason = "threaded comments";
                return false;
            }

            int declaredCommentCount = sheet.WorksheetPart.WorksheetCommentsPart?
                .Comments?
                .CommentList?
                .Elements<Comment>()
                .Count() ?? 0;
            if (!SupportsCommentCount(declaredCommentCount)) {
                reason = "comment counts outside BIFF8 limits";
                return false;
            }

            IReadOnlyList<CommentInfo> comments = GetWorksheetComments(sheet, fontTable, out reason);
            if (reason != null) {
                return false;
            }

            if (!SupportsCommentCount(comments.Count)) {
                reason = "comment counts outside BIFF8 limits";
                return false;
            }

            foreach (CommentInfo comment in comments) {
                if (!SupportsComment(comment, out reason)) {
                    return false;
                }
            }

            return true;
        }

        internal static IReadOnlyList<CommentRecordSet> CreateCommentRecordSets(
            ExcelSheet sheet,
            LegacyXlsFontTable fontTable,
            ushort drawingId) {
            string? reason;
            var records = new List<CommentRecordSet>();
            CommentInfo[] comments = GetWorksheetComments(sheet, fontTable, out reason)
                .Where(comment => SupportsComment(comment, out _))
                .ToArray();
            int shapeCount = comments.Length;
            uint shapeIdBase = checked((uint)drawingId << 10);
            uint lastShapeId = checked(shapeIdBase + (uint)shapeCount);
            ushort objectId = 1;
            foreach (CommentInfo comment in comments) {
                records.Add(BuildCommentRecordSet(
                    comment,
                    objectId,
                    drawingId,
                    shapeIdBase,
                    shapeCount,
                    lastShapeId,
                    objectId == 1));
                objectId++;
            }

            return records;
        }

        internal static byte[] BuildWorkbookDrawingGroupPayload(
            IReadOnlyList<ExcelSheet> sheets,
            LegacyXlsFontTable fontTable) {
            var drawingGroups = new List<DrawingGroupInfo>();
            for (int index = 0; index < sheets.Count; index++) {
                ExcelSheet sheet = sheets[index];
                string? reason;
                int count = 0;
                foreach (CommentInfo comment in GetWorksheetComments(sheet, fontTable, out reason)) {
                    if (SupportsComment(comment, out _)) {
                        count++;
                    }
                }

                if (count != 0) {
                    drawingGroups.Add(new DrawingGroupInfo(checked((ushort)(index + 1)), count));
                }
            }

            if (drawingGroups.Count == 0) {
                return Array.Empty<byte>();
            }

            DrawingGroupInfo lastDrawingGroup = drawingGroups[drawingGroups.Count - 1];
            uint maxShapeId = checked(((uint)lastDrawingGroup.DrawingId << 10) + (uint)lastDrawingGroup.CommentCount + 1U);
            uint savedShapeCount = checked((uint)drawingGroups.Sum(group => group.CommentCount + 1));
            using var drawingGroup = new MemoryStream();
            WriteUInt32(drawingGroup, maxShapeId);
            WriteUInt32(drawingGroup, checked((uint)drawingGroups.Count + 1U));
            WriteUInt32(drawingGroup, savedShapeCount);
            WriteUInt32(drawingGroup, checked((uint)drawingGroups.Count));
            foreach (DrawingGroupInfo group in drawingGroups) {
                WriteUInt32(drawingGroup, group.DrawingId);
                WriteUInt32(drawingGroup, checked((uint)group.CommentCount + 1U));
            }

            byte[] drawingGroupBlock = BuildOfficeArtRecord(0xf006, instance: 0, version: 0x00, drawingGroup.ToArray());
            byte[] drawingProperties = BuildOfficeArtRecord(0xf00b, instance: 3, version: 0x03, [
                0xbf, 0x00, 0x08, 0x00, 0x08, 0x00,
                0x81, 0x01, 0x41, 0x00, 0x00, 0x08,
                0xc0, 0x01, 0x40, 0x00, 0x00, 0x08
            ]);
            byte[] splitMenuColors = BuildOfficeArtRecord(0xf11e, instance: 4, version: 0x00, [
                0x0d, 0x00, 0x00, 0x08,
                0x0c, 0x00, 0x00, 0x08,
                0x17, 0x00, 0x00, 0x08,
                0xf7, 0x00, 0x00, 0x10
            ]);
            return BuildOfficeArtRecord(
                0xf000,
                instance: 0,
                version: 0x0f,
                Combine(drawingGroupBlock, drawingProperties, splitMenuColors));
        }

        private static IReadOnlyList<CommentInfo> GetWorksheetComments(ExcelSheet sheet, LegacyXlsFontTable fontTable, out string? reason) {
            reason = null;
            WorksheetCommentsPart? commentsPart = sheet.WorksheetPart.WorksheetCommentsPart;
            Comments? comments = commentsPart?.Comments;
            if (comments?.CommentList == null) {
                if (HasLegacyVmlDrawingContent(sheet)) {
                    reason = "legacy VML drawings or shapes";
                }

                return Array.Empty<CommentInfo>();
            }

            var commentReferences = new HashSet<string>(
                comments.CommentList.Elements<Comment>()
                    .Select(comment => comment.Reference?.Value)
                    .Where(reference => !string.IsNullOrWhiteSpace(reference))
                    .Select(reference => reference!),
                StringComparer.OrdinalIgnoreCase);
            Dictionary<string, CommentShapeInfo> shapes = ReadCommentShapes(sheet, commentReferences, out reason);
            if (reason != null) {
                return Array.Empty<CommentInfo>();
            }

            List<string> authors = comments.Authors?.Elements<Author>().Select(author => author.Text ?? string.Empty).ToList()
                ?? new List<string>();
            var results = new List<CommentInfo>();
            foreach (Comment comment in comments.CommentList.Elements<Comment>()) {
                string? reference = comment.Reference?.Value;
                if (string.IsNullOrWhiteSpace(reference)) {
                    reason = "comments without cell references";
                    return Array.Empty<CommentInfo>();
                }

                (int row, int column) = A1.ParseCellRef(reference!);
                if (row < 1 || column < 1 || row > 65536 || column > 256) {
                    reason = "comments outside BIFF8 worksheet limits";
                    return Array.Empty<CommentInfo>();
                }

                string author = "OfficeIMO";
                if (comment.AuthorId?.Value != null && comment.AuthorId.Value < authors.Count) {
                    author = string.IsNullOrWhiteSpace(authors[(int)comment.AuthorId.Value])
                        ? "OfficeIMO"
                        : authors[(int)comment.AuthorId.Value];
                }

                if (!TryExtractCommentTextAndRuns(comment.CommentText, fontTable, out string? text, out IReadOnlyList<CommentFormattingRun> formattingRuns, out reason)) {
                    return Array.Empty<CommentInfo>();
                }

                shapes.TryGetValue(reference!, out CommentShapeInfo shape);
                results.Add(new CommentInfo(
                    checked((ushort)(row - 1)),
                    checked((ushort)(column - 1)),
                    text!,
                    author,
                    formattingRuns,
                    shape.Visible,
                    shape.Anchor));
            }

            return results;
        }

        private static bool SupportsComment(CommentInfo comment, out string? reason) {
            reason = null;
            if (comment.Text.Length == 0 || comment.Text.Length > ushort.MaxValue) {
                reason = "comment text lengths outside BIFF8 limits";
                return false;
            }

            if (comment.Author.Length > ushort.MaxValue) {
                reason = "comment author lengths outside BIFF8 limits";
                return false;
            }

            if (GetStringContinuePayloadLength(comment.Text) > BiffMaxRecordDataLength) {
                reason = "comment text payload lengths outside BIFF8 limits";
                return false;
            }

            if (GetNotePayloadLength(comment.Author) > BiffMaxRecordDataLength) {
                reason = "comment author payload lengths outside BIFF8 limits";
                return false;
            }

            int formattingByteCount = checked((comment.FormattingRuns.Count + 1) * 8);
            if (comment.FormattingRuns.Count == 0 || formattingByteCount > BiffMaxRecordDataLength) {
                reason = "comment rich-text formatting runs outside BIFF8 limits";
                return false;
            }

            return true;
        }

        private static bool TryExtractCommentTextAndRuns(
            CommentText? commentText,
            LegacyXlsFontTable fontTable,
            out string? text,
            out IReadOnlyList<CommentFormattingRun> formattingRuns,
            out string? reason) {
            text = null;
            formattingRuns = Array.Empty<CommentFormattingRun>();
            reason = null;
            if (commentText == null) {
                reason = "comments without text";
                return false;
            }

            List<Run> runs = commentText.Elements<Run>().ToList();
            if (runs.Count == 0) {
                text = string.Concat(commentText.Descendants<Text>().Select(item => item.Text ?? string.Empty));
                if (string.IsNullOrEmpty(text)) {
                    reason = "comments without text";
                    return false;
                }

                formattingRuns = new[] { new CommentFormattingRun(0, 0) };
                return true;
            }

            var builder = new StringBuilder();
            var collectedRuns = new List<CommentFormattingRun>();
            foreach (Run run in runs) {
                if (!SupportsCommentRunMetadata(run, out reason)) {
                    return false;
                }

                string runText = run.Text?.Text ?? string.Empty;
                if (runText.Length == 0) {
                    continue;
                }

                if (builder.Length > ushort.MaxValue) {
                    reason = "comment text lengths outside BIFF8 limits";
                    return false;
                }

                if (!fontTable.TryGetFontIndex(run.RunProperties, out ushort fontIndex, out reason)) {
                    return false;
                }

                ushort startCharacter = checked((ushort)builder.Length);
                if (collectedRuns.Count == 0 || collectedRuns[collectedRuns.Count - 1].FontIndex != fontIndex) {
                    collectedRuns.Add(new CommentFormattingRun(startCharacter, fontIndex));
                }

                builder.Append(runText);
            }

            text = builder.ToString();
            if (string.IsNullOrEmpty(text)) {
                reason = "comments without text";
                return false;
            }

            formattingRuns = collectedRuns.Count == 0
                ? new[] { new CommentFormattingRun(0, 0) }
                : collectedRuns;
            return true;
        }

        private static bool SupportsCommentRunMetadata(Run run, out string? reason) {
            reason = null;
            if (run.GetAttributes().Any()) {
                reason = "comment rich-text run metadata";
                return false;
            }

            if (run.ChildElements.Any(child => child is not RunProperties && child is not Text)) {
                reason = "comment rich-text run metadata";
                return false;
            }

            if (run.Elements<Text>().Take(2).Count() > 1) {
                reason = "comment rich-text run metadata";
                return false;
            }

            return true;
        }

        private static bool HasLegacyVmlDrawingContent(ExcelSheet sheet) {
            return sheet.WorksheetPart.Worksheet?.GetFirstChild<LegacyDrawing>() != null
                || sheet.WorksheetPart.VmlDrawingParts.Any();
        }

        private static Dictionary<string, CommentShapeInfo> ReadCommentShapes(ExcelSheet sheet, HashSet<string> commentReferences, out string? reason) {
            reason = null;
            var shapes = new Dictionary<string, CommentShapeInfo>(StringComparer.OrdinalIgnoreCase);
            LegacyDrawing? legacyDrawing = sheet.WorksheetPart.Worksheet?.GetFirstChild<LegacyDrawing>();
            string? relationshipId = legacyDrawing?.Id?.Value;
            if (string.IsNullOrWhiteSpace(relationshipId)) {
                if (sheet.WorksheetPart.VmlDrawingParts.Any()) {
                    reason = "legacy VML drawings or shapes";
                }

                return shapes;
            }

            if (sheet.WorksheetPart.GetPartById(relationshipId!) is not VmlDrawingPart vmlPart) {
                reason = "legacy VML drawings or shapes";
                return shapes;
            }

            if (sheet.WorksheetPart.VmlDrawingParts.Any(part => !ReferenceEquals(part, vmlPart))) {
                reason = "legacy VML drawings or shapes";
                return shapes;
            }

            XDocument document;
            using (Stream stream = vmlPart.GetStream(FileMode.Open, FileAccess.Read)) {
                if (stream.Length == 0) {
                    return shapes;
                }

                document = XDocument.Load(stream, LoadOptions.PreserveWhitespace);
            }

            XNamespace v = "urn:schemas-microsoft-com:vml";
            XNamespace x = "urn:schemas-microsoft-com:office:excel";
            foreach (XElement shape in document.Descendants(v + "shape")) {
                XElement? clientData = shape.Element(x + "ClientData");
                if (clientData == null) {
                    reason = "legacy VML drawings or shapes";
                    return shapes;
                }

                string? objectType = (string?)clientData.Attribute("ObjectType");
                if (!string.Equals(objectType, "Note", StringComparison.OrdinalIgnoreCase)) {
                    reason = "legacy VML drawings or shapes";
                    return shapes;
                }

                if (!SupportsCommentShapeMetadata(shape, v)) {
                    reason = "comment object shape metadata";
                    return shapes;
                }

                if (!SupportsCommentShapeFillMetadata(shape, v)) {
                    reason = "comment object fill metadata";
                    return shapes;
                }

                if (!SupportsCommentShapeLineMetadata(shape, v)) {
                    reason = "comment object line metadata";
                    return shapes;
                }

                if (!SupportsCommentShapeShadowMetadata(shape, v)) {
                    reason = "comment object shadow metadata";
                    return shapes;
                }

                if (!SupportsCommentShapeTextboxMetadata(shape, v)) {
                    reason = "comment object textbox metadata";
                    return shapes;
                }

                if (!SupportsCommentShapeClientDataMetadata(clientData, x)) {
                    reason = "comment object client metadata";
                    return shapes;
                }

                if (!TryParseInt(clientData.Element(x + "Row")?.Value, out int zeroBasedRow)
                    || !TryParseInt(clientData.Element(x + "Column")?.Value, out int zeroBasedColumn)
                    || zeroBasedRow < 0
                    || zeroBasedColumn < 0
                    || zeroBasedRow >= 65536
                    || zeroBasedColumn >= 256) {
                    reason = "legacy VML drawings or shapes";
                    return shapes;
                }

                string reference = A1.CellReference(zeroBasedRow + 1, zeroBasedColumn + 1);
                if (!commentReferences.Contains(reference)) {
                    reason = "legacy VML drawings or shapes";
                    return shapes;
                }

                bool visible = clientData.Element(x + "Visible") != null
                    || ((string?)shape.Attribute("style"))?.IndexOf("visibility:visible", StringComparison.OrdinalIgnoreCase) >= 0;
                CommentAnchor? anchor = TryParseAnchor(clientData.Element(x + "Anchor")?.Value);
                shapes[reference] = new CommentShapeInfo(visible, anchor);
            }

            return shapes;
        }

        private static bool SupportsCommentShapeMetadata(XElement shape, XNamespace v) {
            if (!IsDefaultOrEmptyStyle(
                    (string?)shape.Attribute("style"),
                    "position:absolute;margin-left:0pt;margin-top:0pt;width:108pt;height:59pt;z-index:1;visibility:hidden",
                    "position:absolute;margin-left:0pt;margin-top:0pt;width:108pt;height:59pt;z-index:1;visibility:visible")) {
                return false;
            }

            foreach (XAttribute attribute in shape.Attributes()) {
                string localName = attribute.Name.LocalName;
                if (string.Equals(localName, "id", StringComparison.OrdinalIgnoreCase)
                    || string.Equals(localName, "type", StringComparison.OrdinalIgnoreCase)
                    || string.Equals(localName, "style", StringComparison.OrdinalIgnoreCase)
                    || string.Equals(localName, "fillcolor", StringComparison.OrdinalIgnoreCase)
                    || string.Equals(localName, "strokecolor", StringComparison.OrdinalIgnoreCase)) {
                    continue;
                }

                if (string.Equals(localName, "insetmode", StringComparison.OrdinalIgnoreCase)) {
                    if (!IsDefaultOrEmptyFlag(attribute.Value, "auto")) {
                        return false;
                    }

                    continue;
                }

                if (!string.IsNullOrWhiteSpace(attribute.Value)) {
                    return false;
                }
            }

            XElement? path = shape.Element(v + "path");
            if (path == null) {
                return true;
            }

            foreach (XAttribute attribute in path.Attributes()) {
                if (string.Equals(attribute.Name.LocalName, "connecttype", StringComparison.OrdinalIgnoreCase)) {
                    if (!IsDefaultOrEmptyFlag(attribute.Value, "none")) {
                        return false;
                    }

                    continue;
                }

                if (!string.IsNullOrWhiteSpace(attribute.Value)) {
                    return false;
                }
            }

            return true;
        }

        private static bool SupportsCommentShapeFillMetadata(XElement shape, XNamespace v) {
            const string defaultFillColor = "#ffffe1";
            if (!IsDefaultOrEmptyColor((string?)shape.Attribute("fillcolor"), defaultFillColor)) {
                return false;
            }

            XElement? fill = shape.Element(v + "fill");
            if (fill != null
                && (!IsDefaultOrEmptyColor((string?)fill.Attribute("color"), defaultFillColor)
                    || !IsDefaultOrEmptyColor((string?)fill.Attribute("color2"), defaultFillColor))) {
                return false;
            }

            return true;
        }

        private static bool SupportsCommentShapeLineMetadata(XElement shape, XNamespace v) {
            if (!IsDefaultOrEmptyColor((string?)shape.Attribute("strokecolor"), "#000000", "black", "windowText")) {
                return false;
            }

            XElement? stroke = shape.Element(v + "stroke");
            if (stroke == null) {
                return true;
            }

            foreach (XAttribute attribute in stroke.Attributes()) {
                string localName = attribute.Name.LocalName;
                if (string.Equals(localName, "color", StringComparison.OrdinalIgnoreCase)
                    || string.Equals(localName, "color2", StringComparison.OrdinalIgnoreCase)) {
                    if (!IsDefaultOrEmptyColor(attribute.Value, "#000000", "black", "windowText")) {
                        return false;
                    }

                    continue;
                }

                if (!string.IsNullOrWhiteSpace(attribute.Value)) {
                    return false;
                }
            }

            return true;
        }

        private static bool SupportsCommentShapeShadowMetadata(XElement shape, XNamespace v) {
            XElement? shadow = shape.Element(v + "shadow");
            if (shadow == null) {
                return true;
            }

            foreach (XAttribute attribute in shadow.Attributes()) {
                string localName = attribute.Name.LocalName;
                if (string.Equals(localName, "on", StringComparison.OrdinalIgnoreCase)) {
                    if (!IsDefaultOrEmptyFlag(attribute.Value, "t", "true", "1")) {
                        return false;
                    }

                    continue;
                }

                if (string.Equals(localName, "color", StringComparison.OrdinalIgnoreCase)) {
                    if (!IsDefaultOrEmptyColor(attribute.Value, "black", "#000000")) {
                        return false;
                    }

                    continue;
                }

                if (string.Equals(localName, "obscured", StringComparison.OrdinalIgnoreCase)) {
                    if (!IsDefaultOrEmptyFlag(attribute.Value, "t", "true", "1")) {
                        return false;
                    }

                    continue;
                }

                if (!string.IsNullOrWhiteSpace(attribute.Value)) {
                    return false;
                }
            }

            return true;
        }

        private static bool SupportsCommentShapeTextboxMetadata(XElement shape, XNamespace v) {
            IReadOnlyList<XElement> textboxes = shape.Elements(v + "textbox").ToArray();
            if (textboxes.Count == 0) {
                return true;
            }

            if (textboxes.Count != 1) {
                return false;
            }

            XElement textbox = textboxes[0];
            foreach (XAttribute attribute in textbox.Attributes()) {
                if (string.Equals(attribute.Name.LocalName, "style", StringComparison.OrdinalIgnoreCase)) {
                    if (!IsDefaultOrEmptyStyle(attribute.Value, "mso-direction-alt:auto")) {
                        return false;
                    }

                    continue;
                }

                if (!string.IsNullOrWhiteSpace(attribute.Value)) {
                    return false;
                }
            }

            foreach (XElement child in textbox.Elements()) {
                if (!string.Equals(child.Name.LocalName, "div", StringComparison.OrdinalIgnoreCase)
                    || !string.IsNullOrEmpty(child.Name.NamespaceName)) {
                    return false;
                }

                foreach (XAttribute attribute in child.Attributes()) {
                    if (string.Equals(attribute.Name.LocalName, "style", StringComparison.OrdinalIgnoreCase)) {
                        if (!IsDefaultOrEmptyStyle(attribute.Value, "text-align:left")) {
                            return false;
                        }

                        continue;
                    }

                    if (!string.IsNullOrWhiteSpace(attribute.Value)) {
                        return false;
                    }
                }

                if (!string.IsNullOrWhiteSpace(child.Value)) {
                    return false;
                }
            }

            return true;
        }

        private static bool SupportsCommentShapeClientDataMetadata(XElement clientData, XNamespace x) {
            foreach (XAttribute attribute in clientData.Attributes()) {
                if (string.Equals(attribute.Name.LocalName, "ObjectType", StringComparison.OrdinalIgnoreCase)) {
                    continue;
                }

                if (!string.IsNullOrWhiteSpace(attribute.Value)) {
                    return false;
                }
            }

            foreach (XElement child in clientData.Elements()) {
                if (child.Name.Namespace != x) {
                    return false;
                }

                string localName = child.Name.LocalName;
                if (string.Equals(localName, "MoveWithCells", StringComparison.OrdinalIgnoreCase)
                    || string.Equals(localName, "SizeWithCells", StringComparison.OrdinalIgnoreCase)) {
                    if (!IsDefaultOrEmptyFlag(child.Value, "t", "true", "1")) {
                        return false;
                    }

                    continue;
                }

                if (string.Equals(localName, "AutoFill", StringComparison.OrdinalIgnoreCase)) {
                    if (!IsDefaultOrEmptyFlag(child.Value, "false", "f", "0")) {
                        return false;
                    }

                    continue;
                }

                if (string.Equals(localName, "Visible", StringComparison.OrdinalIgnoreCase)) {
                    if (!IsDefaultOrEmptyFlag(child.Value, "t", "true", "1")) {
                        return false;
                    }

                    continue;
                }

                if (string.Equals(localName, "Anchor", StringComparison.OrdinalIgnoreCase)
                    || string.Equals(localName, "Row", StringComparison.OrdinalIgnoreCase)
                    || string.Equals(localName, "Column", StringComparison.OrdinalIgnoreCase)) {
                    continue;
                }

                return false;
            }

            return true;
        }

        private static bool IsDefaultOrEmptyColor(string? value, string defaultValue) {
            string trimmed = value?.Trim() ?? string.Empty;
            return trimmed.Length == 0
                || string.Equals(trimmed, defaultValue, StringComparison.OrdinalIgnoreCase);
        }

        private static bool IsDefaultOrEmptyColor(string? value, params string[] defaultValues) {
            string trimmed = value?.Trim() ?? string.Empty;
            if (trimmed.Length == 0) {
                return true;
            }

            foreach (string defaultValue in defaultValues) {
                if (string.Equals(trimmed, defaultValue, StringComparison.OrdinalIgnoreCase)) {
                    return true;
                }
            }

            return false;
        }

        private static bool IsDefaultOrEmptyStyle(string? value, string defaultValue) {
            return IsDefaultOrEmptyStyle(value, new[] { defaultValue });
        }

        private static bool IsDefaultOrEmptyStyle(string? value, params string[] defaultValues) {
            string normalized = NormalizeStyle(value);
            if (normalized.Length == 0) {
                return true;
            }

            foreach (string defaultValue in defaultValues) {
                if (string.Equals(normalized, NormalizeStyle(defaultValue), StringComparison.OrdinalIgnoreCase)) {
                    return true;
                }
            }

            return false;
        }

        private static string NormalizeStyle(string? value) {
            return string.Join(
                ";",
                (value ?? string.Empty)
                    .Split(new[] { ';' }, StringSplitOptions.RemoveEmptyEntries)
                    .Select(part => part.Trim())
                    .Where(part => part.Length > 0));
        }

        private static bool IsDefaultOrEmptyFlag(string? value, params string[] defaultValues) {
            string trimmed = value?.Trim() ?? string.Empty;
            if (trimmed.Length == 0) {
                return true;
            }

            foreach (string defaultValue in defaultValues) {
                if (string.Equals(trimmed, defaultValue, StringComparison.OrdinalIgnoreCase)) {
                    return true;
                }
            }

            return false;
        }

        private static bool TryParseInt(string? text, out int value) {
            return int.TryParse(text?.Trim(), System.Globalization.NumberStyles.Integer, System.Globalization.CultureInfo.InvariantCulture, out value);
        }

        private static CommentAnchor? TryParseAnchor(string? text) {
            if (string.IsNullOrWhiteSpace(text)) {
                return null;
            }

            string[] parts = text!.Split(',').Select(part => part.Trim()).ToArray();
            if (parts.Length != 8) {
                return null;
            }

            var values = new ushort[8];
            for (int i = 0; i < parts.Length; i++) {
                if (!ushort.TryParse(parts[i], System.Globalization.NumberStyles.Integer, System.Globalization.CultureInfo.InvariantCulture, out values[i])) {
                    return null;
                }
            }

            return new CommentAnchor(values[0], values[1], values[2], values[3], values[4], values[5], values[6], values[7]);
        }

        private static CommentRecordSet BuildCommentRecordSet(
            CommentInfo comment,
            ushort objectId,
            ushort drawingId,
            uint shapeIdBase,
            int shapeCount,
            uint lastShapeId,
            bool firstShape) {
            CommentAnchor anchor = ClampAnchor(comment.Anchor ?? GetDefaultAnchor(comment.Row, comment.Column), comment.Row, comment.Column);
            return new CommentRecordSet(
                BuildDrawingPayload(comment, anchor, objectId, drawingId, shapeIdBase, shapeCount, lastShapeId, firstShape),
                BuildObjectPayload(objectId),
                BuildOfficeArtRecord(0xf00d, instance: 0, version: 0x00, Array.Empty<byte>()),
                BuildTextObjectPayload(comment.Text, comment.FormattingRuns.Count),
                BuildStringContinuePayload(comment.Text),
                BuildFormattingContinuePayload(checked((ushort)comment.Text.Length), comment.FormattingRuns),
                BuildNotePayload(comment, objectId));
        }

        private static CommentAnchor GetDefaultAnchor(ushort row, ushort column) {
            ushort endColumn = checked((ushort)Math.Min(column + 3, 255));
            ushort endRow = checked((ushort)Math.Min(row + 4, 65535));
            return new CommentAnchor(column, 15, row, 2, endColumn, 15, endRow, 16);
        }

        private static CommentAnchor ClampAnchor(CommentAnchor anchor, ushort row, ushort column) {
            ushort startColumn = Clamp(anchor.StartColumn, 0, 255);
            ushort startRow = Clamp(anchor.StartRow, 0, 65535);
            ushort endColumn = Clamp(anchor.EndColumn, 0, 255);
            ushort endRow = Clamp(anchor.EndRow, 0, 65535);

            if (startColumn > endColumn || (startColumn == endColumn && anchor.StartDx > anchor.EndDx)) {
                startColumn = column;
                endColumn = checked((ushort)Math.Min(column + 3, 255));
            }

            if (startRow > endRow || (startRow == endRow && anchor.StartDy > anchor.EndDy)) {
                startRow = row;
                endRow = checked((ushort)Math.Min(row + 4, 65535));
            }

            return new CommentAnchor(
                startColumn,
                anchor.StartDx,
                startRow,
                anchor.StartDy,
                endColumn,
                anchor.EndDx,
                endRow,
                anchor.EndDy);
        }

        private static ushort Clamp(ushort value, ushort min, ushort max) {
            if (value < min) {
                return min;
            }

            if (value > max) {
                return max;
            }

            return value;
        }

        private static byte[] BuildDrawingPayload(
            CommentInfo comment,
            CommentAnchor anchor,
            ushort objectId,
            ushort drawingId,
            uint shapeIdBase,
            int shapeCount,
            uint lastShapeId,
            bool firstShape) {
            byte[] shape = BuildOfficeArtRecord(0xf00a, instance: 0x00ca, version: 0x02, BuildShapePayload(shapeIdBase + objectId, 0x00000a00));
            byte[] shapeProperties = BuildOfficeArtRecord(0xf00b, instance: 10, version: 0x03, BuildCommentShapeProperties(comment, objectId));
            byte[] clientAnchor = BuildOfficeArtRecord(0xf010, instance: 0, version: 0x00, BuildClientAnchorPayload(anchor));
            byte[] clientData = BuildOfficeArtRecord(0xf011, instance: 0, version: 0x00, Array.Empty<byte>());
            byte[] partialShapeContainer = Combine(
                BuildOfficeArtHeader(0xf004, instance: 0, version: 0x0f, CommentShapeContainerPayloadLength),
                shape,
                shapeProperties,
                clientAnchor,
                clientData);
            if (!firstShape) {
                return partialShapeContainer;
            }

            byte[] drawingInfo = BuildOfficeArtRecord(0xf008, instance: drawingId, version: 0x00, BuildDrawingInfoPayload(shapeCount, lastShapeId));
            byte[] groupShape = BuildGroupShapeContainer(shapeIdBase);
            int shapeGroupPayloadLength = checked(groupShape.Length + (shapeCount * CommentShapeContainerLength));
            int drawingPayloadLength = checked(drawingInfo.Length + 8 + shapeGroupPayloadLength);
            return Combine(
                BuildOfficeArtHeader(0xf002, instance: 0, version: 0x0f, drawingPayloadLength),
                drawingInfo,
                BuildOfficeArtHeader(0xf003, instance: 0, version: 0x0f, shapeGroupPayloadLength),
                groupShape,
                partialShapeContainer);
        }

        private static byte[] BuildGroupShapeContainer(uint shapeIdBase) {
            byte[] groupBounds = BuildOfficeArtRecord(0xf009, instance: 0, version: 0x01, new byte[16]);
            byte[] groupShape = BuildOfficeArtRecord(0xf00a, instance: 0, version: 0x02, BuildShapePayload(shapeIdBase, 0x00000005));
            return BuildOfficeArtRecord(0xf004, instance: 0, version: 0x0f, Combine(groupBounds, groupShape));
        }

        private static byte[] BuildCommentShapeProperties(CommentInfo comment, ushort objectId) {
            using var stream = new MemoryStream();
            WriteOfficeArtProperty(stream, 0x0080, BuildCommentTextId(comment.Text, objectId));
            WriteOfficeArtProperty(stream, 0x0085, 0);
            WriteOfficeArtProperty(stream, 0x0087, 0);
            WriteOfficeArtProperty(stream, 0x0181, 0x08000050);
            WriteOfficeArtProperty(stream, 0x01bf, 0x00010000);
            WriteOfficeArtProperty(stream, 0x01c0, 0x08000040);
            WriteOfficeArtProperty(stream, 0x01cb, 0x00002535);
            WriteOfficeArtProperty(stream, 0x01ce, 0);
            WriteOfficeArtProperty(stream, 0x01ff, 0x00080008);
            WriteOfficeArtProperty(stream, 0x03bf, comment.Visible ? 0x000a0000U : 0x010a0002U);
            return stream.ToArray();
        }

        private static uint BuildCommentTextId(string text, ushort objectId) {
            const uint OffsetBasis = 2166136261;
            const uint Prime = 16777619;
            uint hash = OffsetBasis;
            unchecked {
                foreach (char character in text) {
                    hash = (hash ^ character) * Prime;
                }
                hash = (hash ^ objectId) * Prime;
            }
            return hash == 0 ? objectId : hash;
        }

        private static void WriteOfficeArtProperty(Stream stream, ushort propertyId, uint value) {
            WriteUInt16(stream, propertyId);
            WriteUInt32(stream, value);
        }

        private static byte[] BuildObjectPayload(ushort objectId) {
            using var stream = new MemoryStream();
            WriteUInt16(stream, 0x0015);
            WriteUInt16(stream, 0x0012);
            WriteUInt16(stream, 0x0019);
            WriteUInt16(stream, objectId);
            WriteUInt16(stream, 0x4011);
            WriteUInt32(stream, 0);
            WriteUInt32(stream, 0);
            WriteUInt32(stream, 0);
            WriteUInt16(stream, 0x000d);
            WriteUInt16(stream, 0x0016);
            stream.Write(new byte[16], 0, 16);
            WriteUInt16(stream, 0x0000);
            WriteUInt32(stream, 0);
            WriteUInt32(stream, 0);
            return stream.ToArray();
        }

        private static byte[] BuildTextObjectPayload(string text, int formattingRunCount) {
            using var stream = new MemoryStream();
            WriteUInt16(stream, 0x0212);
            WriteUInt16(stream, 0);
            WriteUInt16(stream, 0);
            WriteUInt32(stream, 0);
            WriteUInt16(stream, checked((ushort)text.Length));
            WriteUInt16(stream, checked((ushort)((formattingRunCount + 1) * 8)));
            WriteUInt16(stream, 0);
            WriteUInt16(stream, 0);
            return stream.ToArray();
        }

        private static byte[] BuildStringContinuePayload(string text) {
            using var stream = new MemoryStream();
            byte[] textBytes = EncodeUnicodeString(text, out byte flags);
            stream.WriteByte(flags);
            stream.Write(textBytes, 0, textBytes.Length);
            return stream.ToArray();
        }

        private static byte[] BuildFormattingContinuePayload(ushort textLength, IReadOnlyList<CommentFormattingRun> formattingRuns) {
            using var stream = new MemoryStream();
            foreach (CommentFormattingRun run in formattingRuns) {
                WriteUInt16(stream, run.StartCharacter);
                WriteUInt16(stream, run.FontIndex);
                WriteUInt32(stream, 0);
            }

            WriteUInt16(stream, textLength);
            WriteUInt16(stream, 0);
            WriteUInt32(stream, 0);
            return stream.ToArray();
        }

        private static byte[] BuildNotePayload(CommentInfo comment, ushort objectId) {
            using var stream = new MemoryStream();
            WriteUInt16(stream, comment.Row);
            WriteUInt16(stream, comment.Column);
            WriteUInt16(stream, comment.Visible ? (ushort)0x0002 : (ushort)0);
            WriteUInt16(stream, objectId);
            WriteUInt16(stream, checked((ushort)comment.Author.Length));
            byte[] authorBytes = EncodeUnicodeString(comment.Author, out byte flags);
            stream.WriteByte(flags);
            stream.Write(authorBytes, 0, authorBytes.Length);
            stream.WriteByte(0);
            return stream.ToArray();
        }

        private static long GetStringContinuePayloadLength(string text) {
            return 1L + GetEncodedUnicodeStringByteCount(text);
        }

        private static long GetNotePayloadLength(string author) {
            return 12L + GetEncodedUnicodeStringByteCount(author);
        }

        private static long GetEncodedUnicodeStringByteCount(string text) {
            return CanUseCompressedString(text) ? text.Length : 2L * text.Length;
        }

        private static byte[] BuildDrawingInfoPayload(int shapeCount, uint lastShapeId) {
            using var stream = new MemoryStream();
            WriteUInt32(stream, checked((uint)(shapeCount + 1)));
            WriteUInt32(stream, lastShapeId);
            return stream.ToArray();
        }

        private static byte[] BuildShapePayload(uint shapeId, uint flags) {
            using var stream = new MemoryStream();
            WriteUInt32(stream, shapeId);
            WriteUInt32(stream, flags);
            return stream.ToArray();
        }

        private static byte[] BuildClientAnchorPayload(CommentAnchor anchor) {
            using var stream = new MemoryStream();
            WriteUInt16(stream, 0);
            WriteUInt16(stream, anchor.StartColumn);
            WriteUInt16(stream, anchor.StartDx);
            WriteUInt16(stream, anchor.StartRow);
            WriteUInt16(stream, anchor.StartDy);
            WriteUInt16(stream, anchor.EndColumn);
            WriteUInt16(stream, anchor.EndDx);
            WriteUInt16(stream, anchor.EndRow);
            WriteUInt16(stream, anchor.EndDy);
            return stream.ToArray();
        }

        private static byte[] BuildOfficeArtRecord(ushort recordType, ushort instance, byte version, byte[] payload) {
            return Combine(BuildOfficeArtHeader(recordType, instance, version, payload.Length), payload);
        }

        private static byte[] BuildOfficeArtHeader(ushort recordType, ushort instance, byte version, int payloadLength) {
            using var stream = new MemoryStream();
            WriteUInt16(stream, checked((ushort)((instance << 4) | (version & 0x0f))));
            WriteUInt16(stream, recordType);
            WriteUInt32(stream, checked((uint)payloadLength));
            return stream.ToArray();
        }

        private static byte[] Combine(params byte[][] arrays) {
            int length = arrays.Sum(array => array.Length);
            byte[] combined = new byte[length];
            int offset = 0;
            foreach (byte[] array in arrays) {
                Buffer.BlockCopy(array, 0, combined, offset, array.Length);
                offset += array.Length;
            }

            return combined;
        }

        private static byte[] EncodeUnicodeString(string text, out byte flags) {
            if (CanUseCompressedString(text)) {
                flags = 0;
                return Encoding.ASCII.GetBytes(text);
            }

            flags = 1;
            return Encoding.Unicode.GetBytes(text);
        }

        private static bool CanUseCompressedString(string text) {
            for (int i = 0; i < text.Length; i++) {
                if (text[i] > 0x7f) {
                    return false;
                }
            }

            return true;
        }

        private static void WriteUInt16(Stream stream, ushort value) {
            stream.WriteByte((byte)(value & 0xff));
            stream.WriteByte((byte)((value >> 8) & 0xff));
        }

        private static void WriteUInt32(Stream stream, uint value) {
            stream.WriteByte((byte)(value & 0xff));
            stream.WriteByte((byte)((value >> 8) & 0xff));
            stream.WriteByte((byte)((value >> 16) & 0xff));
            stream.WriteByte((byte)((value >> 24) & 0xff));
        }

        internal readonly struct CommentRecordSet {
            internal CommentRecordSet(byte[] drawingPayload, byte[] objectPayload, byte[] textboxPayload, byte[] textObjectPayload, byte[] textPayload, byte[] formattingPayload, byte[] notePayload) {
                DrawingPayload = drawingPayload;
                ObjectPayload = objectPayload;
                TextboxPayload = textboxPayload;
                TextObjectPayload = textObjectPayload;
                TextPayload = textPayload;
                FormattingPayload = formattingPayload;
                NotePayload = notePayload;
            }

            internal byte[] DrawingPayload { get; }
            internal byte[] ObjectPayload { get; }
            internal byte[] TextboxPayload { get; }
            internal byte[] TextObjectPayload { get; }
            internal byte[] TextPayload { get; }
            internal byte[] FormattingPayload { get; }
            internal byte[] NotePayload { get; }
        }

        private readonly struct CommentInfo {
            internal CommentInfo(ushort row, ushort column, string text, string author, IReadOnlyList<CommentFormattingRun> formattingRuns, bool visible, CommentAnchor? anchor) {
                Row = row;
                Column = column;
                Text = text;
                Author = author;
                FormattingRuns = formattingRuns;
                Visible = visible;
                Anchor = anchor;
            }

            internal ushort Row { get; }
            internal ushort Column { get; }
            internal string Text { get; }
            internal string Author { get; }
            internal IReadOnlyList<CommentFormattingRun> FormattingRuns { get; }
            internal bool Visible { get; }
            internal CommentAnchor? Anchor { get; }
        }

        private readonly struct CommentFormattingRun {
            internal CommentFormattingRun(ushort startCharacter, ushort fontIndex) {
                StartCharacter = startCharacter;
                FontIndex = fontIndex;
            }

            internal ushort StartCharacter { get; }
            internal ushort FontIndex { get; }
        }

        private readonly struct DrawingGroupInfo {
            internal DrawingGroupInfo(ushort drawingId, int commentCount) {
                DrawingId = drawingId;
                CommentCount = commentCount;
            }

            internal ushort DrawingId { get; }
            internal int CommentCount { get; }
        }

        private readonly struct CommentShapeInfo {
            internal CommentShapeInfo(bool visible, CommentAnchor? anchor) {
                Visible = visible;
                Anchor = anchor;
            }

            internal bool Visible { get; }
            internal CommentAnchor? Anchor { get; }
        }

        private readonly struct CommentAnchor {
            internal CommentAnchor(ushort startColumn, ushort startDx, ushort startRow, ushort startDy, ushort endColumn, ushort endDx, ushort endRow, ushort endDy) {
                StartColumn = startColumn;
                StartDx = startDx;
                StartRow = startRow;
                StartDy = startDy;
                EndColumn = endColumn;
                EndDx = endDx;
                EndRow = endRow;
                EndDy = endDy;
            }

            internal ushort StartColumn { get; }
            internal ushort StartDx { get; }
            internal ushort StartRow { get; }
            internal ushort StartDy { get; }
            internal ushort EndColumn { get; }
            internal ushort EndDx { get; }
            internal ushort EndRow { get; }
            internal ushort EndDy { get; }
        }
    }
}

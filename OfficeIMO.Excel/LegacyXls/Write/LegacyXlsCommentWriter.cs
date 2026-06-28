using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Text;
using System.Xml.Linq;

namespace OfficeIMO.Excel.LegacyXls.Write {
    internal static class LegacyXlsCommentWriter {
        private const int BiffMaxRecordDataLength = 8224;

        internal static bool SupportsWorksheetComments(ExcelSheet sheet, LegacyXlsFontTable fontTable, out string? reason) {
            reason = null;
            if (sheet.WorksheetPart.WorksheetThreadedCommentsParts.Any()) {
                reason = "threaded comments";
                return false;
            }

            foreach (CommentInfo comment in GetWorksheetComments(sheet, fontTable, out reason)) {
                if (!SupportsComment(comment, out reason)) {
                    return false;
                }
            }

            return reason == null;
        }

        internal static IReadOnlyList<CommentRecordSet> CreateCommentRecordSets(ExcelSheet sheet, LegacyXlsFontTable fontTable) {
            string? reason;
            var records = new List<CommentRecordSet>();
            ushort objectId = 1;
            foreach (CommentInfo comment in GetWorksheetComments(sheet, fontTable, out reason)) {
                if (!SupportsComment(comment, out _)) {
                    continue;
                }

                records.Add(BuildCommentRecordSet(comment, objectId));
                objectId++;
            }

            return records;
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

        private static CommentRecordSet BuildCommentRecordSet(CommentInfo comment, ushort objectId) {
            CommentAnchor anchor = ClampAnchor(comment.Anchor ?? GetDefaultAnchor(comment.Row, comment.Column), comment.Row, comment.Column);
            return new CommentRecordSet(
                BuildDrawingPayload(anchor, objectId),
                BuildObjectPayload(objectId),
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

        private static byte[] BuildDrawingPayload(CommentAnchor anchor, ushort objectId) {
            byte[] drawingInfo = BuildOfficeArtRecord(0xf008, instance: 1, version: 0x00, BuildDrawingInfoPayload());
            byte[] shape = BuildOfficeArtRecord(0xf00a, instance: 0x00ca, version: 0x02, BuildShapePayload(0x00000400U + objectId, 0x00000a00));
            byte[] shapeProperties = BuildOfficeArtRecord(0xf00b, instance: 2, version: 0x03, Array.Empty<byte>());
            byte[] clientAnchor = BuildOfficeArtRecord(0xf010, instance: 0, version: 0x00, BuildClientAnchorPayload(anchor));
            byte[] shapeContainer = BuildOfficeArtRecord(0xf004, instance: 0, version: 0x0f, Combine(shape, shapeProperties, clientAnchor));
            return BuildOfficeArtRecord(0xf002, instance: 1, version: 0x0f, Combine(drawingInfo, shapeContainer));
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

        private static byte[] BuildDrawingInfoPayload() {
            using var stream = new MemoryStream();
            WriteUInt32(stream, 1);
            WriteUInt32(stream, 1024);
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
            using var stream = new MemoryStream();
            WriteUInt16(stream, checked((ushort)((instance << 4) | (version & 0x0f))));
            WriteUInt16(stream, recordType);
            WriteUInt32(stream, checked((uint)payload.Length));
            stream.Write(payload, 0, payload.Length);
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
            internal CommentRecordSet(byte[] drawingPayload, byte[] objectPayload, byte[] textObjectPayload, byte[] textPayload, byte[] formattingPayload, byte[] notePayload) {
                DrawingPayload = drawingPayload;
                ObjectPayload = objectPayload;
                TextObjectPayload = textObjectPayload;
                TextPayload = textPayload;
                FormattingPayload = formattingPayload;
                NotePayload = notePayload;
            }

            internal byte[] DrawingPayload { get; }
            internal byte[] ObjectPayload { get; }
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

using OfficeIMO.PowerPoint.LegacyPpt.Internal;
using static OfficeIMO.PowerPoint.LegacyPpt.Internal.LegacyPptXml;

namespace OfficeIMO.PowerPoint.LegacyPpt.Write {
    internal static partial class LegacyPptWriter {
        private const ushort RecordStyleTextPropAtomForWrite = 0x0FA1;
        private const ushort RecordTextRulerAtomForWrite = 0x0FA6;

        internal static LegacyPptWriterFontCatalog
            CreateFontCatalogForWrite() => new(Template.Value.Document);

        /// <summary>The logical text of a text shape in the presentation captured for the write.</summary>
        internal static string ReadLogicalTextForWrite(
            LegacyPptPackageView view, PowerPointTextBox textBox) {
            ArgumentNullException.ThrowIfNull(view);
            ArgumentNullException.ThrowIfNull(textBox);
            return ReadLogicalTextForWrite(ResolveShapeForWrite(view, textBox));
        }

        /// <summary>The logical text of a text shape as the document holds it now.</summary>
        internal static string ReadLogicalTextForWrite(
            PowerPointTextBox textBox) {
            ArgumentNullException.ThrowIfNull(textBox);
            return ReadLogicalTextForWrite(textBox.XmlElement);
        }

        /// <summary>The logical text of a <c>p:sp</c> element: paragraphs, line breaks and field placeholders.</summary>
        internal static string ReadLogicalTextForWrite(OfficeXmlNode shape) {
            OfficeXmlNode body = shape.Is(P, "sp") ? shape.Child(P, "txBody") : default;
            if (body.IsNone) return string.Empty;
            var result = new System.Text.StringBuilder();
            bool first = true;
            foreach (OfficeXmlNode paragraph in body.Children(A, "p")) {
                if (!first) result.Append('\n');
                first = false;
                foreach (OfficeXmlNode child in paragraph.Children()) {
                    if (child.Is(A, "r")) {
                        result.Append(RunText(child) ?? string.Empty);
                    } else if (child.Is(A, "fld")) {
                        result.Append('*');
                    } else if (child.Is(A, "br")) {
                        result.Append('\v');
                    }
                }
            }
            return result.ToString();
        }

        // The text of a run's (or field's) a:t, or null when it has none.
        private static string? RunText(OfficeXmlNode run) => run.ChildText(A, "t");

        internal static bool TryReadTextBoxForWrite(
            PowerPointTextBox textBox, LegacyPptWriterFontCatalog fonts,
            out string? reason) => TryBuildTextBoxContent(textBox, fonts,
            LegacyPptWriterPictureBulletCatalog.Empty, out _, out reason);

        internal static bool TryReadTextBoxForWrite(
            PowerPointTextBox textBox, LegacyPptWriterFontCatalog fonts,
            LegacyPptWriterPictureBulletCatalog pictureBullets,
            out string? reason) => TryBuildTextBoxContent(textBox, fonts,
            pictureBullets, out _, out reason);

        internal static bool TryReadTextBoxForWrite(LegacyPptPackageView view,
            PowerPointTextBox textBox, LegacyPptWriterFontCatalog fonts,
            LegacyPptWriterPictureBulletCatalog pictureBullets,
            out string? reason) => TryBuildTextBoxContent(view, textBox, fonts,
            pictureBullets, out _, out reason);

        internal static bool TryBuildTextBoxContent(
            PowerPointTextBox textBox, LegacyPptWriterFontCatalog fonts,
            out LegacyPptWriterTextBoxContent? content,
            out string? reason) => TryBuildTextBoxContent(textBox, fonts,
                LegacyPptWriterPictureBulletCatalog.Empty, out content,
                out reason);

        /// <summary>
        /// Builds a text shape's binary text from the presentation the picture-bullet catalog was read from, or from the
        /// document as it is now when the catalog has no presentation.
        /// </summary>
        internal static bool TryBuildTextBoxContent(
            PowerPointTextBox textBox, LegacyPptWriterFontCatalog fonts,
            LegacyPptWriterPictureBulletCatalog pictureBullets,
            out LegacyPptWriterTextBoxContent? content,
            out string? reason) {
            ArgumentNullException.ThrowIfNull(textBox);
            ArgumentNullException.ThrowIfNull(pictureBullets);
            return TryBuildTextBoxContent(ResolveShapeForWrite(pictureBullets.View, textBox),
                fonts, pictureBullets, out content, out reason);
        }

        /// <summary>Builds a text shape's binary text from the presentation captured for the write.</summary>
        internal static bool TryBuildTextBoxContent(LegacyPptPackageView view,
            PowerPointTextBox textBox, LegacyPptWriterFontCatalog fonts,
            LegacyPptWriterPictureBulletCatalog pictureBullets,
            out LegacyPptWriterTextBoxContent? content,
            out string? reason) {
            ArgumentNullException.ThrowIfNull(view);
            ArgumentNullException.ThrowIfNull(textBox);
            return TryBuildTextBoxContent(ResolveShapeForWrite(view, textBox),
                fonts, pictureBullets, out content, out reason);
        }

        /// <summary>Builds the binary text of a <c>p:sp</c> element.</summary>
        internal static bool TryBuildTextBoxContent(
            OfficeXmlNode shape, LegacyPptWriterFontCatalog fonts,
            LegacyPptWriterPictureBulletCatalog pictureBullets,
            out LegacyPptWriterTextBoxContent? content,
            out string? reason) {
            ArgumentNullException.ThrowIfNull(fonts);
            ArgumentNullException.ThrowIfNull(pictureBullets);
            content = null;
            reason = null;
            OfficeXmlNode body = shape.Is(P, "sp") ? shape.Child(P, "txBody") : default;
            if (body.IsNone) {
                reason = "The text shape has no DrawingML text body.";
                return false;
            }
            if (!TryReadTextFrameForWrite(shape, out _, out reason)) {
                return false;
            }
            return TryBuildTextBodyContent(body, fonts, pictureBullets, out content,
                out reason);
        }

        /// <summary>Builds the binary text of a table cell (<c>a:tc</c>).</summary>
        internal static bool TryBuildTableCellContent(
            OfficeXmlNode cell, LegacyPptWriterFontCatalog fonts,
            LegacyPptWriterPictureBulletCatalog pictureBullets,
            out LegacyPptWriterTextBoxContent? content,
            out string? reason) {
            ArgumentNullException.ThrowIfNull(fonts);
            ArgumentNullException.ThrowIfNull(pictureBullets);
            OfficeXmlNode body = cell.Child(A, "txBody");
            if (body.IsNone) {
                content = null;
                reason = "The table cell has no DrawingML text body.";
                return false;
            }
            return TryBuildTextBodyContent(body, fonts, pictureBullets, out content,
                out reason);
        }

        private static bool TryBuildTextBodyContent(
            OfficeXmlNode body, LegacyPptWriterFontCatalog fonts,
            LegacyPptWriterPictureBulletCatalog pictureBullets,
            out LegacyPptWriterTextBoxContent? content,
            out string? reason) {
            content = null;
            reason = null;
            OfficeXmlNode bodyProperties = body.Child(A, "bodyPr");
            OfficeXmlNode listStyle = body.Child(A, "lstStyle");
            bool unsupportedChild = false;
            foreach (OfficeXmlNode child in body.Children()) {
                if (!child.Is(A, "bodyPr") && !child.Is(A, "lstStyle") && !child.Is(A, "p")) {
                    unsupportedChild = true;
                    break;
                }
            }
            if (bodyProperties.IsNone || listStyle.IsNone
                || HasAnyAttribute(body)
                || HasAnyAttribute(listStyle) || HasChildElements(listStyle)
                || unsupportedChild) {
                reason = "The text body contains content outside the base binary text contract.";
                return false;
            }
            List<OfficeXmlNode> paragraphs = body.Children(A, "p").ToList();
            if (paragraphs.Count == 0) {
                reason = "The text body must contain at least one paragraph.";
                return false;
            }
            if (!TryBuildTextRulerRecord(paragraphs, out byte[]? rulerRecord,
                    out reason)) return false;
            if (!TryBuildStyleTextProp9Record(paragraphs,
                    pictureBullets, out byte[]? style9Record,
                    out reason)) return false;
            if (!TryBuildTextSpecialInfoRecord(paragraphs,
                    out byte[]? specialInfoRecord,
                    out reason)) return false;

            var logicalText = new System.Text.StringBuilder();
            using var paragraphRuns = new MemoryStream();
            using var characterRuns = new MemoryStream();
            var fieldRecords = new List<byte[]>();
            for (int index = 0; index < paragraphs.Count; index++) {
                OfficeXmlNode paragraph = paragraphs[index];
                byte? ppt9RunId = style9Record == null
                    ? null
                    : checked((byte)(index % 16));
                if (!TryWriteParagraphRun(paragraph, fonts, logicalText,
                        paragraphRuns, characterRuns, fieldRecords,
                        ppt9RunId,
                        out reason)) {
                    content = null;
                    return false;
                }
            }
            string styledText = logicalText.ToString();
            if (styledText.Length == 0
                || styledText[styledText.Length - 1] != '\r') {
                reason = "The encoded text body has no terminal paragraph marker.";
                content = null;
                return false;
            }
            string text = styledText.Substring(0, styledText.Length - 1);
            using var payload = new MemoryStream();
            paragraphRuns.Position = 0;
            paragraphRuns.CopyTo(payload);
            characterRuns.Position = 0;
            characterRuns.CopyTo(payload);
            byte[]? styleRecord = null;
            if (HasBinaryTextFormatting(paragraphs, rulerRecord != null)) {
                styleRecord = BuildRecord(version: 0, instance: 0,
                    RecordStyleTextPropAtomForWrite, payload.ToArray());
            }
            content = new LegacyPptWriterTextBoxContent(text, styleRecord,
                rulerRecord, style9Record, specialInfoRecord, fieldRecords);
            return true;
        }

        private static bool TryWriteParagraphRun(OfficeXmlNode paragraph,
            LegacyPptWriterFontCatalog fonts,
            System.Text.StringBuilder logicalText, Stream paragraphRuns,
            Stream characterRuns, ICollection<byte[]> fieldRecords,
            byte? ppt9RunId,
            out string? reason) {
            reason = null;
            bool unsupportedChild = false;
            foreach (OfficeXmlNode child in paragraph.Children()) {
                if (!child.Is(A, "pPr") && !child.Is(A, "r") && !child.Is(A, "br")
                    && !child.Is(A, "fld") && !child.Is(A, "endParaRPr")) {
                    unsupportedChild = true;
                    break;
                }
            }
            if (HasAnyAttribute(paragraph)
                || unsupportedChild
                || paragraph.Children(A, "pPr").Count > 1
                || paragraph.Children(A, "endParaRPr").Count > 1) {
                reason = "A paragraph contains extensions or duplicate properties that are not in the base binary text contract.";
                return false;
            }
            OfficeXmlNode sourceProperties = paragraph.Child(A, "pPr");
            if (!sourceProperties.Child(A, "defRPr").IsNone) {
                reason = "Shape-level default run properties must be materialized on individual runs before binary writing.";
                return false;
            }
            int level = sourceProperties.Int("lvl") ?? 0;
            if (level < 0 || level > 4) {
                reason = "Base binary PowerPoint text supports paragraph levels zero through four.";
                return false;
            }
            var paragraphText = new System.Text.StringBuilder();
            foreach (OfficeXmlNode child in paragraph.Children()) {
                if (child.Is(A, "r")) {
                    if (!TryWriteCharacterRun(child, fonts, paragraphText,
                            characterRuns, ppt9RunId, out reason)) {
                        return false;
                    }
                } else if (child.Is(A, "br")) {
                    if (!TryWriteLineBreak(child, fonts,
                            paragraphText, characterRuns, ppt9RunId,
                            out reason)) return false;
                } else if (child.Is(A, "fld")) {
                    if (!TryWriteField(child, fonts,
                            checked(logicalText.Length
                                + paragraphText.Length), paragraphText,
                            characterRuns, fieldRecords, ppt9RunId,
                            out reason)) return false;
                }
            }
            paragraphText.Append('\r');
            logicalText.Append(paragraphText);
            WriteUInt32(paragraphRuns, checked((uint)paragraphText.Length));
            WriteUInt16(paragraphRuns, checked((ushort)level));
            // The level, margins, default tab size and tab stops go to the shared text ruler instead.
            if (!TryWriteParagraphException(paragraphRuns, sourceProperties,
                    fonts, out reason,
                    allowAutoNumbering: true, excludeRuler: true)) return false;

            OfficeXmlNode endProperties = paragraph.Child(A, "endParaRPr");
            WriteUInt32(characterRuns, 1U);
            return TryWriteCharacterException(characterRuns,
                endProperties, fonts, out reason, ppt9RunId, excludeMetadata: true);
        }

        private static bool TryWriteCharacterRun(OfficeXmlNode run,
            LegacyPptWriterFontCatalog fonts,
            System.Text.StringBuilder paragraphText, Stream characterRuns,
            byte? ppt9RunId, out string? reason) {
            reason = null;
            bool unsupportedChild = false;
            foreach (OfficeXmlNode child in run.Children()) {
                if (!child.Is(A, "rPr") && !child.Is(A, "t")) {
                    unsupportedChild = true;
                    break;
                }
            }
            if (HasAnyAttribute(run)
                || unsupportedChild
                || run.Children(A, "rPr").Count > 1
                || run.Children(A, "t").Count != 1) {
                reason = "A text run contains unsupported or duplicate content.";
                return false;
            }
            string value = (RunText(run) ?? string.Empty)
                .Replace("\r\n", "\r").Replace("\n", "\r");
            if (value.IndexOf('\0') >= 0) {
                reason = "Run text contains a NUL character that has no binary PowerPoint text representation.";
                return false;
            }
            if (value.Length == 0) return true;
            paragraphText.Append(value);
            WriteUInt32(characterRuns, checked((uint)value.Length));
            return TryWriteCharacterException(characterRuns,
                run.Child(A, "rPr"), fonts,
                out reason, ppt9RunId, excludeMetadata: true);
        }

        private static bool TryWriteLineBreak(OfficeXmlNode lineBreak,
            LegacyPptWriterFontCatalog fonts,
            System.Text.StringBuilder paragraphText,
            Stream characterRuns, byte? ppt9RunId,
            out string? reason) {
            reason = null;
            bool unsupportedChild = false;
            foreach (OfficeXmlNode child in lineBreak.Children()) {
                if (!child.Is(A, "rPr")) {
                    unsupportedChild = true;
                    break;
                }
            }
            if (HasAnyAttribute(lineBreak)
                || unsupportedChild
                || lineBreak.Children(A, "rPr").Count > 1) {
                reason = "A line break contains unsupported or duplicate run properties.";
                return false;
            }
            paragraphText.Append('\v');
            WriteUInt32(characterRuns, 1U);
            return TryWriteCharacterException(characterRuns,
                lineBreak.Child(A, "rPr"),
                fonts, out reason, ppt9RunId, excludeMetadata: true);
        }

        private static bool TryWriteField(OfficeXmlNode field,
            LegacyPptWriterFontCatalog fonts, int position,
            System.Text.StringBuilder paragraphText,
            Stream characterRuns, ICollection<byte[]> fieldRecords,
            byte? ppt9RunId, out string? reason) {
            reason = null;
            bool unsupportedChild = false;
            foreach (OfficeXmlNode child in field.Children()) {
                if (!child.Is(A, "rPr") && !child.Is(A, "pPr") && !child.Is(A, "t")) {
                    unsupportedChild = true;
                    break;
                }
            }
            string? fieldType = field.Attribute("type");
            if (!HasOnlyAttributes(field, "id", "type")
                || !Guid.TryParse(field.Attribute("id"), out _)
                || string.IsNullOrWhiteSpace(fieldType)
                || unsupportedChild
                || field.Children(A, "rPr").Count > 1
                || field.Children(A, "pPr").Count > 1
                || field.Children(A, "t").Count != 1) {
                reason = "A DrawingML field must contain a valid id, one supported type, one text value, and at most one run or paragraph property element.";
                return false;
            }
            OfficeXmlNode fieldParagraph = field.Child(A, "pPr");
            if (!fieldParagraph.IsNone
                && (HasAnyAttribute(fieldParagraph) || HasChildElements(fieldParagraph))) {
                reason = "Field-local paragraph formatting has no native classic binary PowerPoint mapping.";
                return false;
            }
            if (!TryBuildTextFieldRecord(fieldType!, position,
                    out byte[] fieldRecord, out reason)) return false;
            paragraphText.Append('*');
            WriteUInt32(characterRuns, 1U);
            if (!TryWriteCharacterException(characterRuns,
                    field.Child(A, "rPr"),
                    fonts, out reason, ppt9RunId, excludeMetadata: true)) return false;
            fieldRecords.Add(fieldRecord);
            return true;
        }

        private static bool TryBuildTextFieldRecord(string fieldType,
            int position, out byte[] record, out string? reason) {
            reason = null;
            record = Array.Empty<byte>();
            ushort recordType;
            byte? dateTimeIndex = null;
            byte[]? rtfFormat = null;
            if (string.Equals(fieldType, "slidenum",
                    StringComparison.OrdinalIgnoreCase)) {
                recordType = 0x0FD8;
            } else if (string.Equals(fieldType, "datetime",
                           StringComparison.OrdinalIgnoreCase)
                       || string.Equals(fieldType, "datetimeFigureOut",
                           StringComparison.OrdinalIgnoreCase)) {
                recordType = 0x0FF8;
            } else if (fieldType.StartsWith("datetimeRtf:",
                           StringComparison.OrdinalIgnoreCase)) {
                recordType = 0x1015;
                try {
                    rtfFormat = Convert.FromBase64String(fieldType.Substring(
                        "datetimeRtf:".Length));
                } catch (FormatException) {
                    reason = "The legacy RTF date field contains an invalid encoded format string.";
                    return false;
                }
                if ((rtfFormat.Length & 1) != 0
                    || rtfFormat.Length > 128) {
                    reason = "The legacy RTF date field format exceeds the 64-character binary limit.";
                    return false;
                }
            } else if (string.Equals(fieldType, "header",
                           StringComparison.OrdinalIgnoreCase)) {
                recordType = 0x0FF9;
            } else if (string.Equals(fieldType, "footer",
                           StringComparison.OrdinalIgnoreCase)) {
                recordType = 0x0FFA;
            } else if (fieldType.StartsWith("datetime",
                           StringComparison.OrdinalIgnoreCase)
                       && int.TryParse(fieldType.Substring("datetime".Length),
                           System.Globalization.NumberStyles.None,
                           System.Globalization.CultureInfo.InvariantCulture,
                           out int oneBasedIndex)
                       && oneBasedIndex >= 1 && oneBasedIndex <= 13) {
                recordType = 0x0FF7;
                dateTimeIndex = checked((byte)(oneBasedIndex - 1));
            } else {
                reason = $"DrawingML field type '{fieldType}' has no native classic binary PowerPoint metacharacter mapping.";
                return false;
            }
            int payloadLength = dateTimeIndex.HasValue ? 8
                : rtfFormat != null ? 132 : 4;
            var payload = new byte[payloadLength];
            WriteUInt32(payload, 0, checked((uint)position));
            if (dateTimeIndex.HasValue) payload[4] = dateTimeIndex.Value;
            if (rtfFormat != null) {
                Buffer.BlockCopy(rtfFormat, 0, payload, 4,
                    rtfFormat.Length);
            }
            record = BuildRecord(version: 0, instance: 0,
                recordType, payload);
            return true;
        }

        internal static bool IsTextMetaCharacterRecord(ushort type) =>
            type is 0x0FD8 or 0x0FF7 or 0x0FF8 or 0x0FF9
                or 0x0FFA or 0x1015;

        private static bool HasBinaryTextFormatting(
            IEnumerable<OfficeXmlNode> paragraphs, bool hasRuler) {
            if (hasRuler) return true;
            foreach (OfficeXmlNode paragraph in paragraphs) {
                OfficeXmlNode source = paragraph.Child(A, "pPr");
                if (!source.IsNone) {
                    if ((source.Int("lvl") ?? 0) != 0) return true;
                    if (HasParagraphFormattingOutsideRuler(source)) return true;
                }
                foreach (OfficeXmlNode run in paragraph.Children(A, "r")) {
                    if (HasCharacterFormatting(run.Child(A, "rPr"))) return true;
                }
                foreach (OfficeXmlNode field in paragraph.Children(A, "fld")) {
                    if (HasCharacterFormatting(field.Child(A, "rPr"))) return true;
                }
                foreach (OfficeXmlNode lineBreak in paragraph.Children(A, "br")) {
                    if (HasCharacterFormatting(lineBreak.Child(A, "rPr"))) return true;
                }
                if (HasCharacterFormatting(paragraph.Child(A, "endParaRPr"))) return true;
            }
            return false;
        }

        private static bool TryBuildTextRulerRecord(
            IReadOnlyList<OfficeXmlNode> paragraphs, out byte[]? record,
            out string? reason) {
            record = null;
            reason = null;
            short? defaultTab = null;
            IReadOnlyList<KeyValuePair<short, ushort>>? tabStops = null;
            var leftMargins = new Dictionary<int, short>();
            var indents = new Dictionary<int, short>();
            int maximumLevel = 0;
            foreach (OfficeXmlNode paragraph in paragraphs) {
                OfficeXmlNode properties = paragraph.Child(A, "pPr");
                if (properties.IsNone) continue;
                int level = properties.Int("lvl") ?? 0;
                if (level < 0 || level > 4) {
                    reason = "Base binary PowerPoint text supports paragraph levels zero through four.";
                    return false;
                }
                maximumLevel = Math.Max(maximumLevel, level);
                if (properties.Int("marL") is int leftMargin) {
                    if (!TryToMasterInt16(leftMargin,
                            out short value)
                        || !TrySetLevelValue(leftMargins, level, value)) {
                        reason = "Paragraphs at one level use incompatible or out-of-range left margins for the shared binary text ruler.";
                        return false;
                    }
                }
                if (properties.Int("indent") is int indent) {
                    if (!TryToMasterInt16(indent,
                            out short value)
                        || !TrySetLevelValue(indents, level, value)) {
                        reason = "Paragraphs at one level use incompatible or out-of-range first-line indents for the shared binary text ruler.";
                        return false;
                    }
                }
                if (properties.Int("defTabSz") is int defaultTabSize) {
                    if (!TryToMasterInt16(defaultTabSize,
                            out short value) || value < 0
                        || defaultTab.HasValue
                        && defaultTab.Value != value) {
                        reason = "Paragraphs use incompatible or out-of-range default tab sizes for the shared binary text ruler.";
                        return false;
                    }
                    defaultTab = value;
                }
                OfficeXmlNode list = properties.Child(A, "tabLst");
                if (list.IsNone) continue;
                var values = new List<KeyValuePair<short, ushort>>();
                if (!TryReadTabStops(list, values, out bool unsupportedList)) {
                    reason = unsupportedList
                        ? "A paragraph tab-stop list contains unsupported content."
                        : "A paragraph tab stop has an unsupported position or alignment.";
                    return false;
                }
                if (tabStops != null && !tabStops.SequenceEqual(values)) {
                    reason = "Paragraphs use different explicit tab-stop lists that cannot share one binary text ruler.";
                    return false;
                }
                tabStops = values;
            }
            if (!defaultTab.HasValue && tabStops == null
                && leftMargins.Count == 0 && indents.Count == 0) {
                return true;
            }
            uint mask = 1U << 1;
            if (defaultTab.HasValue) mask |= 1U;
            if (tabStops != null) mask |= 1U << 2;
            foreach (int level in leftMargins.Keys) mask |= 1U << (3 + level);
            foreach (int level in indents.Keys) mask |= 1U << (8 + level);
            using var payload = new MemoryStream();
            WriteUInt32(payload, mask);
            WriteInt16(payload, checked((short)(maximumLevel + 1)));
            if (defaultTab.HasValue) WriteInt16(payload, defaultTab.Value);
            if (tabStops != null) {
                WriteUInt16(payload, checked((ushort)tabStops.Count));
                foreach (KeyValuePair<short, ushort> tab in tabStops) {
                    WriteInt16(payload, tab.Key);
                    WriteUInt16(payload, tab.Value);
                }
            }
            for (int level = 0; level < 5; level++) {
                if (leftMargins.TryGetValue(level, out short left)) {
                    WriteInt16(payload, left);
                }
                if (indents.TryGetValue(level, out short indent)) {
                    WriteInt16(payload, indent);
                }
            }
            record = BuildRecord(version: 0, instance: 0,
                RecordTextRulerAtomForWrite, payload.ToArray());
            return true;
        }

        private static bool TrySetLevelValue(IDictionary<int, short> values,
            int level, short value) {
            if (values.TryGetValue(level, out short existing)) {
                return existing == value;
            }
            values.Add(level, value);
            return true;
        }

        internal sealed class LegacyPptWriterTextBoxContent {
            internal LegacyPptWriterTextBoxContent(string text,
                byte[]? styleRecord, byte[]? rulerRecord,
                byte[]? style9Record, byte[]? specialInfoRecord,
                IEnumerable<byte[]> fieldRecords) {
                Text = text;
                StyleRecord = styleRecord;
                RulerRecord = rulerRecord;
                Style9Record = style9Record;
                SpecialInfoRecord = specialInfoRecord;
                FieldRecords = fieldRecords.ToArray();
            }

            internal string Text { get; }

            internal byte[]? StyleRecord { get; }

            internal byte[]? RulerRecord { get; }

            internal byte[]? Style9Record { get; }

            internal byte[]? SpecialInfoRecord { get; }

            internal IReadOnlyList<byte[]> FieldRecords { get; }
        }
    }
}

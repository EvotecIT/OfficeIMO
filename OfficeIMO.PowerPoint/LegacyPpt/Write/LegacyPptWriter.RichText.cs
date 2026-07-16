using DocumentFormat.OpenXml;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace OfficeIMO.PowerPoint.LegacyPpt.Write {
    internal static partial class LegacyPptWriter {
        private const ushort RecordStyleTextPropAtomForWrite = 0x0FA1;
        private const ushort RecordTextRulerAtomForWrite = 0x0FA6;

        internal static LegacyPptWriterFontCatalog
            CreateFontCatalogForWrite() => new(Template.Value.Document);

        internal static bool TryReadTextBoxForWrite(
            PowerPointTextBox textBox, LegacyPptWriterFontCatalog fonts,
            out string? reason) => TryBuildTextBoxContent(textBox, fonts,
            out _, out reason);

        internal static bool TryBuildTextBoxContent(
            PowerPointTextBox textBox, LegacyPptWriterFontCatalog fonts,
            out LegacyPptWriterTextBoxContent? content,
            out string? reason) {
            if (textBox == null) throw new ArgumentNullException(
                nameof(textBox));
            if (fonts == null) throw new ArgumentNullException(nameof(fonts));
            content = null;
            reason = null;
            if (textBox.Element is not P.Shape shape
                || shape.TextBody == null) {
                reason = "The text shape has no DrawingML text body.";
                return false;
            }
            P.TextBody body = shape.TextBody;
            if (!TryReadTextFrameForWrite(textBox, out _, out reason)) {
                return false;
            }
            if (body.BodyProperties == null || body.ListStyle == null
                || body.ListStyle.HasAttributes || body.ListStyle.HasChildren
                || body.ChildElements.Any(child => child is not A.BodyProperties
                    and not A.ListStyle and not A.Paragraph)) {
                reason = "The text body contains content outside the base binary text contract.";
                return false;
            }
            A.Paragraph[] paragraphs = body.Elements<A.Paragraph>().ToArray();
            if (paragraphs.Length == 0) {
                reason = "The text body must contain at least one paragraph.";
                return false;
            }
            if (!TryBuildTextRulerRecord(paragraphs, out byte[]? rulerRecord,
                    out reason)) return false;
            if (!TryBuildStyleTextProp9Record(paragraphs,
                    out byte[]? style9Record, out reason)) return false;

            var logicalText = new System.Text.StringBuilder();
            using var paragraphRuns = new MemoryStream();
            using var characterRuns = new MemoryStream();
            for (int index = 0; index < paragraphs.Length; index++) {
                A.Paragraph paragraph = paragraphs[index];
                byte? ppt9RunId = style9Record == null
                    ? null
                    : checked((byte)(index % 16));
                if (!TryWriteParagraphRun(paragraph, fonts, logicalText,
                        paragraphRuns, characterRuns, ppt9RunId,
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
                rulerRecord, style9Record);
            return true;
        }

        private static bool TryWriteParagraphRun(A.Paragraph paragraph,
            LegacyPptWriterFontCatalog fonts,
            System.Text.StringBuilder logicalText, Stream paragraphRuns,
            Stream characterRuns, byte? ppt9RunId,
            out string? reason) {
            reason = null;
            if (paragraph.HasAttributes
                || paragraph.ChildElements.Any(child => child
                    is not A.ParagraphProperties and not A.Run
                    and not A.EndParagraphRunProperties)
                || paragraph.Elements<A.ParagraphProperties>().Count() > 1
                || paragraph.Elements<A.EndParagraphRunProperties>().Count()
                    > 1) {
                reason = "A paragraph contains fields, line breaks, extensions, or duplicate properties that are not in the base binary text contract.";
                return false;
            }
            A.ParagraphProperties? sourceProperties = paragraph
                .ParagraphProperties;
            if (sourceProperties?.GetFirstChild<A.DefaultRunProperties>()
                != null) {
                reason = "Shape-level default run properties must be materialized on individual runs before binary writing.";
                return false;
            }
            int level = sourceProperties?.Level?.Value ?? 0;
            if (level < 0 || level > 4) {
                reason = "Base binary PowerPoint text supports paragraph levels zero through four.";
                return false;
            }
            var paragraphText = new System.Text.StringBuilder();
            foreach (A.Run run in paragraph.Elements<A.Run>()) {
                if (!TryWriteCharacterRun(run, fonts, paragraphText,
                        characterRuns, ppt9RunId, out reason)) return false;
            }
            paragraphText.Append('\r');
            logicalText.Append(paragraphText);
            WriteUInt32(paragraphRuns, checked((uint)paragraphText.Length));
            WriteUInt16(paragraphRuns, checked((ushort)level));
            A.ParagraphProperties? properties = sourceProperties == null
                ? null
                : (A.ParagraphProperties)sourceProperties.CloneNode(true);
            if (properties != null) {
                properties.Level = null;
                properties.LeftMargin = null;
                properties.Indent = null;
                properties.DefaultTabSize = null;
                properties.RemoveAllChildren<A.TabStopList>();
            }
            if (!TryWriteParagraphException(paragraphRuns, properties,
                    fonts, out reason,
                    allowAutoNumbering: true)) return false;

            A.EndParagraphRunProperties? endProperties = paragraph
                .GetFirstChild<A.EndParagraphRunProperties>();
            WriteUInt32(characterRuns, 1U);
            return TryWriteCharacterException(characterRuns,
                NormalizeCharacterProperties(endProperties), fonts,
                out reason, ppt9RunId);
        }

        private static bool TryWriteCharacterRun(A.Run run,
            LegacyPptWriterFontCatalog fonts,
            System.Text.StringBuilder paragraphText, Stream characterRuns,
            byte? ppt9RunId, out string? reason) {
            reason = null;
            if (run.HasAttributes
                || run.ChildElements.Any(child => child is not A.RunProperties
                    and not A.Text)
                || run.Elements<A.RunProperties>().Count() > 1
                || run.Elements<A.Text>().Count() != 1) {
                reason = "A text run contains unsupported or duplicate content.";
                return false;
            }
            string value = (run.Text?.Text ?? string.Empty)
                .Replace("\r\n", "\r").Replace("\n", "\r");
            if (value.IndexOf('\0') >= 0) {
                reason = "Run text contains a NUL character that has no binary PowerPoint text representation.";
                return false;
            }
            if (value.Length == 0) return true;
            paragraphText.Append(value);
            WriteUInt32(characterRuns, checked((uint)value.Length));
            return TryWriteCharacterException(characterRuns,
                NormalizeCharacterProperties(run.RunProperties), fonts,
                out reason, ppt9RunId);
        }

        private static A.TextCharacterPropertiesType?
            NormalizeCharacterProperties(
                A.TextCharacterPropertiesType? properties) {
            if (properties == null) return null;
            var clone = (A.TextCharacterPropertiesType)properties
                .CloneNode(true);
            foreach (OpenXmlAttribute attribute in clone.GetAttributes()
                         .Where(attribute => attribute.LocalName is "lang"
                             or "altLang" or "dirty" or "smtClean")
                         .ToArray()) {
                clone.RemoveAttribute(attribute.LocalName,
                    attribute.NamespaceUri);
            }
            clone.RemoveAllChildren<A.HyperlinkOnClick>();
            clone.RemoveAllChildren<A.HyperlinkOnMouseOver>();
            return clone;
        }

        private static bool HasBinaryTextFormatting(
            IEnumerable<A.Paragraph> paragraphs, bool hasRuler) {
            if (hasRuler) return true;
            foreach (A.Paragraph paragraph in paragraphs) {
                A.ParagraphProperties? source = paragraph
                    .ParagraphProperties;
                if (source != null) {
                    if ((source.Level?.Value ?? 0) != 0) return true;
                    var clone = (A.ParagraphProperties)source
                        .CloneNode(true);
                    clone.Level = null;
                    clone.LeftMargin = null;
                    clone.Indent = null;
                    clone.DefaultTabSize = null;
                    clone.RemoveAllChildren<A.TabStopList>();
                    if (clone.HasAttributes || clone.HasChildren) return true;
                }
                foreach (A.Run run in paragraph.Elements<A.Run>()) {
                    A.TextCharacterPropertiesType? properties =
                        NormalizeCharacterProperties(run.RunProperties);
                    if (properties is { HasAttributes: true }
                        || properties is { HasChildren: true }) return true;
                }
                A.TextCharacterPropertiesType? end =
                    NormalizeCharacterProperties(paragraph
                        .GetFirstChild<A.EndParagraphRunProperties>());
                if (end is { HasAttributes: true }
                    || end is { HasChildren: true }) return true;
            }
            return false;
        }

        private static bool TryBuildTextRulerRecord(
            IReadOnlyList<A.Paragraph> paragraphs, out byte[]? record,
            out string? reason) {
            record = null;
            reason = null;
            short? defaultTab = null;
            IReadOnlyList<KeyValuePair<short, ushort>>? tabStops = null;
            var leftMargins = new Dictionary<int, short>();
            var indents = new Dictionary<int, short>();
            int maximumLevel = 0;
            foreach (A.Paragraph paragraph in paragraphs) {
                A.ParagraphProperties? properties = paragraph
                    .ParagraphProperties;
                if (properties == null) continue;
                int level = properties.Level?.Value ?? 0;
                maximumLevel = Math.Max(maximumLevel, level);
                if (properties.LeftMargin?.HasValue == true) {
                    if (!TryToMasterInt16(properties.LeftMargin.Value,
                            out short value)
                        || !TrySetLevelValue(leftMargins, level, value)) {
                        reason = "Paragraphs at one level use incompatible or out-of-range left margins for the shared binary text ruler.";
                        return false;
                    }
                }
                if (properties.Indent?.HasValue == true) {
                    if (!TryToMasterInt16(properties.Indent.Value,
                            out short value)
                        || !TrySetLevelValue(indents, level, value)) {
                        reason = "Paragraphs at one level use incompatible or out-of-range first-line indents for the shared binary text ruler.";
                        return false;
                    }
                }
                if (properties.DefaultTabSize?.HasValue == true) {
                    if (!TryToMasterInt16(properties.DefaultTabSize.Value,
                            out short value) || value < 0
                        || defaultTab.HasValue
                        && defaultTab.Value != value) {
                        reason = "Paragraphs use incompatible or out-of-range default tab sizes for the shared binary text ruler.";
                        return false;
                    }
                    defaultTab = value;
                }
                A.TabStopList? list = properties
                    .GetFirstChild<A.TabStopList>();
                if (list == null) continue;
                if (!TryReadRulerTabStops(list,
                        out IReadOnlyList<KeyValuePair<short, ushort>> values,
                        out reason)) return false;
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

        private static bool TryReadRulerTabStops(A.TabStopList list,
            out IReadOnlyList<KeyValuePair<short, ushort>> values,
            out string? reason) {
            var result = new List<KeyValuePair<short, ushort>>();
            values = result;
            reason = null;
            if (list.HasAttributes || list.Elements<A.TabStop>().Count()
                    != list.ChildElements.Count) {
                reason = "A paragraph tab-stop list contains unsupported content.";
                return false;
            }
            foreach (A.TabStop tab in list.Elements<A.TabStop>()) {
                if (!HasOnlyAttributes(tab, "pos", "algn")
                    || tab.Position?.HasValue != true
                    || tab.Alignment?.HasValue != true
                    || !TryToMasterInt16(tab.Position.Value,
                        out short position)
                    || !TryMapTabAlignment(tab.Alignment.Value,
                        out ushort alignment)) {
                    reason = "A paragraph tab stop has an unsupported position or alignment.";
                    return false;
                }
                result.Add(new KeyValuePair<short, ushort>(position,
                    alignment));
            }
            return true;
        }

        internal sealed class LegacyPptWriterTextBoxContent {
            internal LegacyPptWriterTextBoxContent(string text,
                byte[]? styleRecord, byte[]? rulerRecord,
                byte[]? style9Record) {
                Text = text;
                StyleRecord = styleRecord;
                RulerRecord = rulerRecord;
                Style9Record = style9Record;
            }

            internal string Text { get; }

            internal byte[]? StyleRecord { get; }

            internal byte[]? RulerRecord { get; }

            internal byte[]? Style9Record { get; }
        }
    }
}

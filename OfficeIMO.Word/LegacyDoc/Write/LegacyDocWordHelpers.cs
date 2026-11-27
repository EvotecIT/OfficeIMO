using System.Collections.Concurrent;
using System.Globalization;
using System.Text;
using OfficeIMO.Core.Internal.OpenXml;

namespace OfficeIMO.Word.LegacyDoc.Write {
    /// <summary>The WordprocessingML readings the native DOC writer shares with the rest of the Word engine.</summary>
    internal static class LegacyDocWordHelpers {
        private const int NativeDefaultTableLook = 0x04A0;
        private const int ExpandedDefaultTableLook = 0x07E0;
        private const int TableLookFirstRow = 0x0020;
        private const int TableLookLastRow = 0x0040;
        private const int TableLookFirstColumn = 0x0080;
        private const int TableLookLastColumn = 0x0100;
        private const int TableLookNoHorizontalBand = 0x0200;
        private const int TableLookNoVerticalBand = 0x0400;

        private static readonly ConcurrentDictionary<WordTableStyle, Lazy<LegacyDocElement>> BuiltInTableStyles = new();

        /// <summary>
        /// Resolves abstract numbering ids to the first definition with the same authored identity (<c>w:nsid</c>).
        /// Definitions without an identity stay independent.
        /// </summary>
        internal static Dictionary<int, LegacyDocElement> GetCanonicalAbstractDefinitions(LegacyDocElement numbering) {
            var definitions = new Dictionary<int, LegacyDocElement>();
            var identities = new Dictionary<string, LegacyDocElement>(StringComparer.OrdinalIgnoreCase);
            foreach (LegacyDocElement definition in numbering.Elements("abstractNum")) {
                if (LegacyDocValues.Int32(definition.Attr("abstractNumId")) is not int id || definitions.ContainsKey(id)) continue;
                LegacyDocElement canonical = definition;
                string? identity = definition.Child("nsid").Val;
                if (!string.IsNullOrEmpty(identity)) {
                    if (identities.TryGetValue(identity!, out LegacyDocElement first)) canonical = first;
                    else identities.Add(identity!, definition);
                }
                definitions.Add(id, canonical);
            }
            return definitions;
        }

        /// <summary>
        /// A list level's alignment: left, right or center. The logical start and end values map to left and right;
        /// paragraph-only values fall back to left, as Word does.
        /// </summary>
        internal static string? ReadLevelJustification(LegacyDocElement element) {
            string? token = element.Val;
            if (string.IsNullOrWhiteSpace(token)) {
                return null;
            }

            switch (token!.Trim()) {
                case "left":
                case "start":
                    return "left";
                case "right":
                case "end":
                    return "right";
                case "center":
                    return "center";
                case "both":
                case "distribute":
                case "mediumKashida":
                case "highKashida":
                case "lowKashida":
                case "thaiDistribute":
                case "numTab":
                    return "left";
                default:
                    return null;
            }
        }

        /// <summary>Reads Word's native table look; expanded attributes together replace the packed value and default to on.</summary>
        internal static int GetTableLookMask(LegacyDocElement tableLook) {
            if (tableLook.IsNone) return NativeDefaultTableLook;
            if (!HasExpandedTableLookFlags(tableLook)) {
                string? value = tableLook.Val;
                if (string.IsNullOrWhiteSpace(value)) return NativeDefaultTableLook;
                return int.TryParse(value, NumberStyles.HexNumber, CultureInfo.InvariantCulture, out int parsed) ? parsed : 0;
            }

            int mask = ExpandedDefaultTableLook;
            ApplyExpandedTableLookFlag(tableLook.Attr("firstRow"), TableLookFirstRow, ref mask);
            ApplyExpandedTableLookFlag(tableLook.Attr("lastRow"), TableLookLastRow, ref mask);
            ApplyExpandedTableLookFlag(tableLook.Attr("firstColumn"), TableLookFirstColumn, ref mask);
            ApplyExpandedTableLookFlag(tableLook.Attr("lastColumn"), TableLookLastColumn, ref mask);
            ApplyExpandedTableLookFlag(tableLook.Attr("noHBand"), TableLookNoHorizontalBand, ref mask);
            ApplyExpandedTableLookFlag(tableLook.Attr("noVBand"), TableLookNoVerticalBand, ref mask);
            return mask;
        }

        /// <summary>Whether expanded flags, rather than the packed mask, describe the table look.</summary>
        internal static bool HasExpandedTableLookFlags(LegacyDocElement tableLook) =>
            tableLook.HasAttr("firstRow") || tableLook.HasAttr("lastRow") || tableLook.HasAttr("firstColumn") ||
            tableLook.HasAttr("lastColumn") || tableLook.HasAttr("noHBand") || tableLook.HasAttr("noVBand");

        private static void ApplyExpandedTableLookFlag(string? value, int flag, ref int mask) {
            if (value == null) {
                return;
            }

            if (LegacyDocValues.ParseOnOff(value)) {
                mask |= flag;
            } else {
                mask &= ~flag;
            }
        }

        /// <summary>The unit of a preferred width: dxa (also when unstated), pct, auto or nil, or null for another token.</summary>
        internal static string? NativeTableWidthType(LegacyDocElement width) {
            string? token = width.Attr("type");
            if (token == null || token == "dxa") return "dxa";
            if (token == "pct") return "pct";
            if (token == "auto") return "auto";
            if (token == "nil") return "nil";
            return null;
        }

        /// <summary>The hyperlink relationship with this id from <paramref name="owner"/>, or null.</summary>
        internal static LegacyDocRelationship? FindHyperlinkRelationship(LegacyDocPart owner, string? id) {
            if (string.IsNullOrEmpty(id)) return null;
            LegacyDocRelationship? relationship = owner.FindRelationship(id);
            return relationship != null && string.Equals(relationship.Type, LegacyDocSourcePackage.HyperlinkRelationshipType, StringComparison.Ordinal)
                ? relationship
                : null;
        }

        /// <summary>The <c>w:style</c> definition of a built-in table style, read once and shared read-only.</summary>
        internal static LegacyDocElement BuiltInTableStyle(WordTableStyle style) =>
            BuiltInTableStyles.GetOrAdd(style, key => new Lazy<LegacyDocElement>(() => ParseBuiltInTableStyle(key))).Value;

        private static LegacyDocElement ParseBuiltInTableStyle(WordTableStyle style) {
            string markup = WordTableStyles.GetStyleMarkup(style)
                ?? throw new InvalidOperationException("The built-in table style " + style + " has no fixed definition.");
            byte[] bytes = Encoding.UTF8.GetBytes(markup);
            OfficeXmlTree tree = OfficeXmlTree.Parse(bytes, bytes.Length, new OfficeXmlNamespaces.Table());
            return new LegacyDocElement(tree, tree.Root);
        }

        /// <summary>
        /// The section properties of each section in document order, as the document's section list holds them: those
        /// of body paragraphs, then the body's own (none when the body has no final section properties).
        /// </summary>
        internal static IReadOnlyList<LegacyDocElement> GetDocumentSectionProperties(LegacyDocSourcePackage source) {
            var sections = new List<LegacyDocElement>();
            LegacyDocElement body = source.Body;
            foreach (LegacyDocElement paragraph in body.Elements("p")) {
                LegacyDocElement sectionProperties = paragraph.Child("pPr").Child("sectPr");
                if (sectionProperties.Exists) sections.Add(sectionProperties);
            }
            sections.Add(body.Child("sectPr"));
            return sections;
        }
    }
}

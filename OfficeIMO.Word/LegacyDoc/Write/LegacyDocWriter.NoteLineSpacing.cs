using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.Word.OpenXml;

namespace OfficeIMO.Word.LegacyDoc.Write {
    internal static partial class LegacyDocWriter {
        /// <summary>
        /// The note paragraph's properties with the line spacing its style chain or the document defaults give it. When
        /// the paragraph doesn't set a line itself and one is inherited, the result is a detached copy of its
        /// <c>w:pPr</c> (or a new one) carrying that line and rule; the package's trees stay unchanged.
        /// </summary>
        private static LegacyDocElement MaterializeNoteLineSpacing(LegacyDocElement paragraph, string noteKind) {
            LegacyDocElement properties = paragraph.Child("pPr");
            if (!string.IsNullOrWhiteSpace(properties.Child("spacing").Attr("line"))) return properties;
            LegacyDocSourcePackage? source = t_writingSource;
            if (source == null || !source.Styles.Exists) return properties;
            LegacyDocStyleCatalog definitions = source.Styles;
            string? styleId = properties.Child("pStyle").Val ?? (noteKind == "footnote" ? "FootnoteText" : "EndnoteText");
            var visited = new HashSet<string>(StringComparer.Ordinal);
            LegacyDocElement inherited = LegacyDocElement.None;
            while (!string.IsNullOrWhiteSpace(styleId) && visited.Add(styleId!)) {
                if (visited.Count > MaxInheritedStyleDepth) throw new NotSupportedException("Native DOC saving does not support note paragraph style chains beyond 256 levels.");
                LegacyDocElement style = definitions.FindById(styleId);
                inherited = style.Child("pPr").Child("spacing");
                if (!string.IsNullOrWhiteSpace(inherited.Attr("line"))) break;
                styleId = style.Child("basedOn").Val;
            }
            if (string.IsNullOrWhiteSpace(inherited.Attr("line"))) inherited = definitions.DocumentDefaults.Child("pPrDefault").Child("pPr").Child("spacing");
            string? inheritedLine = inherited.Attr("line");
            if (string.IsNullOrWhiteSpace(inheritedLine)) return properties;
            // DOC uses fixed built-in note style slots. Keep the source's effective
            // spacing on the note paragraph when that style cannot carry it.
            OfficeXmlTree copy = properties.IsNone
                ? OfficeXmlTree.Create(paragraph.Tree.Namespaces, LegacyDocElement.W, "pPr")
                : OfficeXmlTree.CopyElement(properties.Tree, properties.Node);
            int spacing = copy.GetOrInsertChild(copy.Root, LegacyDocElement.W, "spacing", WPPrView.Order);
            copy.SetAttribute(spacing, LegacyDocElement.W, "line", inheritedLine!);
            copy.SetAttribute(spacing, LegacyDocElement.W, "lineRule", inherited.Attr("lineRule") ?? "auto");
            return new LegacyDocElement(copy, copy.Root);
        }
    }
}

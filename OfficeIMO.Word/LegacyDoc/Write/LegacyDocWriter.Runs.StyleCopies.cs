using OfficeIMO.Core.Internal.OpenXml;

namespace OfficeIMO.Word.LegacyDoc.Write {
    internal static partial class LegacyDocWriter {
        // CT_Style child sequence, for adding a missing w:pPr or w:rPr where the schema puts it.
        private static readonly OfficeXmlChildOrder StyleChildOrder = new(
            new[] {
                LegacyDocElement.W, LegacyDocElement.W, LegacyDocElement.W, LegacyDocElement.W, LegacyDocElement.W, LegacyDocElement.W,
                LegacyDocElement.W, LegacyDocElement.W, LegacyDocElement.W, LegacyDocElement.W, LegacyDocElement.W, LegacyDocElement.W,
                LegacyDocElement.W, LegacyDocElement.W, LegacyDocElement.W, LegacyDocElement.W, LegacyDocElement.W, LegacyDocElement.W,
                LegacyDocElement.W, LegacyDocElement.W, LegacyDocElement.W, LegacyDocElement.W
            },
            new[] {
                "name", "aliases", "basedOn", "next", "link", "autoRedefine", "hidden", "uiPriority", "semiHidden",
                "unhideWhenUsed", "qFormat", "locked", "personal", "personalCompose", "personalReply", "rsid", "pPr", "rPr",
                "tblPr", "trPr", "tcPr", "tblStylePr"
            });

        /// <summary>
        /// A detached, editable copy of a style definition. The writer materializes document defaults and theme fonts
        /// into copies; the package's own trees stay unchanged.
        /// </summary>
        private static OfficeXmlTree CopyStyleDefinition(LegacyDocElement style) => OfficeXmlTree.CopyElement(style.Tree, style.Node);

        /// <summary>The first child with this <c>w</c> name, added in schema order when the parent has none.</summary>
        private static int GetOrInsertStyleChild(OfficeXmlTree tree, int parent, string localName, OfficeXmlChildOrder order) =>
            tree.GetOrInsertChild(parent, LegacyDocElement.W, localName, order);

        private static bool HasChildWithLocalName(OfficeXmlTree tree, int parent, string localName) {
            for (int child = tree.FirstChildOf(parent); child != OfficeXmlTree.NoNode; child = tree.NextSiblingOf(child)) {
                if (tree.KindOf(child) == OfficeXmlTreeNodeKind.Element && string.Equals(tree.LocalNameOf(child), localName, StringComparison.Ordinal)) {
                    return true;
                }
            }

            return false;
        }

        private static void SetOrRemoveAttribute(OfficeXmlTree tree, int element, string localName, string? value) {
            if (value == null) {
                tree.RemoveAttribute(element, LegacyDocElement.W, localName);
            } else {
                tree.SetAttribute(element, LegacyDocElement.W, localName, value);
            }
        }

        /// <summary>
        /// The font names of a <c>w:rFonts</c> element as the native writer reads them: <c>w:ascii</c>, <c>w:hAnsi</c>,
        /// <c>w:eastAsia</c> and <c>w:cs</c>, with theme selectors already resolved when <see cref="ResolvedTheme"/> is set.
        /// </summary>
        private readonly struct LegacyDocRunFonts {
            internal LegacyDocRunFonts(string? ascii, string? highAnsi, string? eastAsia, string? complexScript, bool resolvedTheme = false) {
                Ascii = ascii;
                HighAnsi = highAnsi;
                EastAsia = eastAsia;
                ComplexScript = complexScript;
                ResolvedTheme = resolvedTheme;
            }

            internal string? Ascii { get; }

            internal string? HighAnsi { get; }

            internal string? EastAsia { get; }

            internal string? ComplexScript { get; }

            /// <summary>Whether the element carried theme selectors that were resolved to these names.</summary>
            internal bool ResolvedTheme { get; }

            internal static LegacyDocRunFonts FromElement(LegacyDocElement fonts) =>
                new(fonts.Attr("ascii"), fonts.Attr("hAnsi"), fonts.Attr("eastAsia"), fonts.Attr("cs"));

            /// <summary>Writes the resolved names onto a <c>w:rFonts</c> element of a detached copy and drops its theme selectors.</summary>
            internal void ApplyTo(OfficeXmlTree tree, int runFonts) {
                SetOrRemoveAttribute(tree, runFonts, "ascii", Ascii);
                SetOrRemoveAttribute(tree, runFonts, "hAnsi", HighAnsi);
                SetOrRemoveAttribute(tree, runFonts, "eastAsia", EastAsia);
                SetOrRemoveAttribute(tree, runFonts, "cs", ComplexScript);
                tree.RemoveAttribute(runFonts, LegacyDocElement.W, "asciiTheme");
                tree.RemoveAttribute(runFonts, LegacyDocElement.W, "hAnsiTheme");
                tree.RemoveAttribute(runFonts, LegacyDocElement.W, "eastAsiaTheme");
                tree.RemoveAttribute(runFonts, LegacyDocElement.W, "cstheme");
            }
        }
    }
}

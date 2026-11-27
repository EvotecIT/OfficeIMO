using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.Word.OpenXml;

namespace OfficeIMO.Word.LegacyDoc.Write {
    internal static partial class LegacyDocWriter {
        private static LegacyDocElement CreateDefaultParagraphStyle(LegacyDocSourcePackage source,
            IReadOnlyDictionary<string, LegacyDocElement> paragraphStyles, LegacyDocElement documentDefaults) {
            OfficeXmlTree tree;
            if (paragraphStyles.TryGetValue("Normal", out LegacyDocElement existing)) {
                tree = CopyStyleDefinition(existing);
            } else {
                tree = OfficeXmlTree.Create(source.Document.Tree.Namespaces, LegacyDocElement.W, "style");
                tree.SetAttribute(tree.Root, LegacyDocElement.W, "type", "paragraph");
                tree.SetAttribute(tree.Root, LegacyDocElement.W, "styleId", "Normal");
                tree.SetAttribute(tree.Root, LegacyDocElement.W, "default", "true");
            }

            int normal = tree.Root;
            int runProperties = GetOrInsertStyleChild(tree, normal, "rPr", StyleChildOrder);
            LegacyDocElement defaults = documentDefaults.Child("rPrDefault").Child("rPr");
            if (defaults.Exists) {
                foreach (LegacyDocElement property in defaults.Elements()) {
                    // Font selectors are inherited individually below.
                    if (!property.Is("sz") && !property.Is("szCs")) continue;
                    if (!HasChildWithLocalName(tree, runProperties, property.LocalName))
                        tree.Append(runProperties, tree.ImportElement(property.Tree, property.Node));
                }
            }
            LegacyDocElement fonts = new LegacyDocElement(tree, runProperties).Child("rFonts");
            LegacyDocElement inherited = defaults.Child("rFonts");
            LegacyDocElement scheme = GetThemeFontScheme(source);
            var effectiveFonts = new LegacyDocRunFonts(
                ResolveDefaultFontSlot(scheme, fonts.Attr("ascii"), fonts.Attr("asciiTheme"), inherited.Attr("ascii"), inherited.Attr("asciiTheme")),
                ResolveDefaultFontSlot(scheme, fonts.Attr("hAnsi"), fonts.Attr("hAnsiTheme"), inherited.Attr("hAnsi"), inherited.Attr("hAnsiTheme")),
                ResolveDefaultFontSlot(scheme, fonts.Attr("eastAsia"), fonts.Attr("eastAsiaTheme"), inherited.Attr("eastAsia"), inherited.Attr("eastAsiaTheme")),
                ResolveDefaultFontSlot(scheme, fonts.Attr("cs"), fonts.Attr("cstheme"), inherited.Attr("cs"), inherited.Attr("cstheme")));
            string? family = ReadSupportedRunFontFamily(effectiveFonts);
            // A native DOC has no theme-based default font. Materialize the resolved default.
            family = string.IsNullOrWhiteSpace(family) ? "Calibri" : family;
            if (fonts.Exists) tree.Remove(fonts.Node);
            int materializedFonts = tree.CreateElement(LegacyDocElement.W, "rFonts");
            tree.SetAttribute(materializedFonts, LegacyDocElement.W, "ascii", family!);
            tree.SetAttribute(materializedFonts, LegacyDocElement.W, "hAnsi", family!);
            tree.InsertBefore(runProperties, materializedFonts, tree.FirstChildOf(runProperties));
            return new LegacyDocElement(tree, normal);
        }

        private static string? ResolveDefaultFontSlot(LegacyDocElement scheme, string? name, string? selector,
            string? inheritedName, string? inheritedSelector) =>
            selector != null || !string.IsNullOrWhiteSpace(name)
                ? ResolveDefaultThemeFont(scheme, selector) ?? name
                : ResolveDefaultThemeFont(scheme, inheritedSelector) ?? inheritedName;

        private static void MaterializeDocumentDefaultSpacing(Dictionary<string, LegacyDocElement> paragraphStyles, LegacyDocElement documentDefaults) {
            LegacyDocElement paragraphDefaults = documentDefaults.Child("pPrDefault").Child("pPr");
            LegacyDocElement defaults = paragraphDefaults.Child("spacing");
            LegacyDocElement alignment = paragraphDefaults.Child("textAlignment");
            if (defaults.IsNone && alignment.IsNone) return;
            // DOC spacing projection only understands these WordprocessingML values.
            // Extension attributes must not be multiplied into every root style.
            LegacyDocAttribute[]? inheritedSpacingAttributes = defaults.IsNone ? null : defaults.Attributes.Where(attribute =>
                attribute.Namespace == LegacyDocElement.W &&
                attribute.LocalName is "before" or "after" or "beforeLines" or "afterLines" or
                    "beforeAutospacing" or "afterAutospacing").ToArray();
            // DOC has no docDefaults record. Put the defaults on roots of the style
            // hierarchy; derived styles continue to inherit their base's overrides.
            foreach (string styleId in paragraphStyles.Keys.ToArray()) {
                LegacyDocElement original = paragraphStyles[styleId];
                string? baseId = original.Child("basedOn").Val;
                if (!string.IsNullOrWhiteSpace(baseId) && paragraphStyles.ContainsKey(baseId!)) continue;
                OfficeXmlTree tree = CopyStyleDefinition(original);
                int properties = GetOrInsertStyleChild(tree, tree.Root, "pPr", StyleChildOrder);
                if (alignment.Exists && tree.FindChild(properties, LegacyDocElement.W, "textAlignment") == OfficeXmlTree.NoNode)
                    tree.InsertOrdered(properties, tree.ImportElement(alignment.Tree, alignment.Node), WPPrView.Order);
                if (defaults.IsNone) {
                    paragraphStyles[styleId] = new LegacyDocElement(tree, tree.Root);
                    continue;
                }
                int spacing = GetOrInsertStyleChild(tree, properties, "spacing", WPPrView.Order);
                bool authoredLine = !string.IsNullOrWhiteSpace(tree.GetAttribute(spacing, LegacyDocElement.W, "line"));
                string? defaultLine = defaults.Attr("line");
                if (!authoredLine && !string.IsNullOrWhiteSpace(defaultLine)) {
                    // Word inherits the numeric value and its interpretation
                    // together. A rule without a local value is not an override.
                    tree.SetAttribute(spacing, LegacyDocElement.W, "line", defaultLine!);
                    tree.SetAttribute(spacing, LegacyDocElement.W, "lineRule", defaults.Attr("lineRule") ?? "auto");
                }
                foreach (LegacyDocAttribute attribute in inheritedSpacingAttributes!) {
                    if (!tree.HasAttribute(spacing, attribute.Namespace, attribute.LocalName)) {
                        tree.SetAttribute(spacing, attribute.Namespace, attribute.LocalName, attribute.Value);
                    }
                }
                paragraphStyles[styleId] = new LegacyDocElement(tree, tree.Root);
            }
        }

        private static string? ResolveDefaultThemeFont(LegacyDocElement scheme, string? selector) {
            if (scheme.IsNone || selector == null) return null;
            string? family = selector switch {
                "majorAscii" or "majorHAnsi" => ThemeTypeface(scheme, "majorFont", "latin"),
                "minorAscii" or "minorHAnsi" => ThemeTypeface(scheme, "minorFont", "latin"),
                "majorEastAsia" => ThemeTypeface(scheme, "majorFont", "ea"),
                "minorEastAsia" => ThemeTypeface(scheme, "minorFont", "ea"),
                "majorBidi" => ThemeTypeface(scheme, "majorFont", "cs"),
                "minorBidi" => ThemeTypeface(scheme, "minorFont", "cs"),
                _ => throw new NotSupportedException("Native DOC saving cannot resolve the default theme font selector.")
            };
            return string.IsNullOrWhiteSpace(family) ? null : family;
        }

        private static string? ThemeTypeface(LegacyDocElement scheme, string fontCollection, string script) =>
            scheme.Child(OfficeXmlNamespaces.DrawingML, fontCollection)
                .Child(OfficeXmlNamespaces.DrawingML, script)
                .Attr(OfficeXmlNamespaces.None, "typeface");
    }
}

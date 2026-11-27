using System.Globalization;
using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.Word.OpenXml;

namespace OfficeIMO.Word.LegacyDoc.Docx {
    /// <summary>
    /// Small helpers over <see cref="OfficeXmlTree"/> for writing WordprocessingML: elements and attributes in the
    /// <c>w</c> namespace, and children inserted where the schema expects them.
    /// </summary>
    internal static class LegacyDocXml {
        internal const int W = OfficeXmlNamespaces.WordprocessingML;

        internal static readonly OfficeXmlChildOrder ParagraphPropertiesOrder = WPPrView.Order;
        internal static readonly OfficeXmlChildOrder RunPropertiesOrder = WRPrView.Order;
        internal static readonly OfficeXmlChildOrder ParagraphMarkRunPropertiesOrder = WParaRPrView.Order;
        internal static readonly OfficeXmlChildOrder SectionPropertiesOrder = WSectPrView.Order;
        internal static readonly OfficeXmlChildOrder ParagraphOrder = WPView.Order;

        internal static readonly OfficeXmlChildOrder StyleOrder = Order(
            "name", "aliases", "basedOn", "next", "link", "autoRedefine", "hidden", "uiPriority", "semiHidden", "unhideWhenUsed",
            "qFormat", "locked", "personal", "personalCompose", "personalReply", "rsid", "pPr", "rPr", "tblPr", "trPr", "tcPr", "tblStylePr");

        internal static readonly OfficeXmlChildOrder LevelOrder = Order(
            "start", "numFmt", "lvlRestart", "pStyle", "isLgl", "suff", "lvlText", "lvlPicBulletId", "legacy", "lvlJc", "pPr", "rPr");

        internal static readonly OfficeXmlChildOrder TablePropertiesOrder = Order(
            "tblStyle", "tblpPr", "tblOverlap", "bidiVisual", "tblStyleRowBandSize", "tblStyleColBandSize", "tblW", "jc", "tblCellSpacing",
            "tblInd", "tblBorders", "shd", "tblLayout", "tblCellMar", "tblLook", "tblCaption", "tblDescription", "tblPrChange");

        internal static readonly OfficeXmlChildOrder TableRowPropertiesOrder = Order(
            "cnfStyle", "divId", "gridBefore", "gridAfter", "wBefore", "wAfter", "cantSplit", "trHeight", "tblHeader", "tblCellSpacing",
            "jc", "hidden", "ins", "del", "trPrChange");

        internal static readonly OfficeXmlChildOrder TableCellPropertiesOrder = Order(
            "cnfStyle", "tcW", "gridSpan", "hMerge", "vMerge", "tcBorders", "shd", "noWrap", "tcMar", "textDirection", "tcFitText", "vAlign",
            "hideMark", "headers", "cellIns", "cellDel", "cellMerge", "tcPrChange");

        internal static readonly OfficeXmlChildOrder BordersOrder = Order(
            "top", "left", "start", "bottom", "right", "end", "insideH", "insideV", "tl2br", "tr2bl", "between", "bar");

        internal static readonly OfficeXmlChildOrder MarginsOrder = Order("top", "left", "start", "bottom", "right", "end");

        internal static readonly OfficeXmlChildOrder SettingsOrder = Order(
            "writeProtection", "view", "zoom", "removePersonalInformation", "removeDateAndTime", "doNotDisplayPageBoundaries",
            "displayBackgroundShape", "printPostScriptOverText", "printFractionalCharacterWidth", "printFormsData", "embedTrueTypeFonts",
            "embedSystemFonts", "saveSubsetFonts", "saveFormsData", "mirrorMargins", "alignBordersAndEdges", "bordersDoNotSurroundHeader",
            "bordersDoNotSurroundFooter", "gutterAtTop", "hideSpellingErrors", "hideGrammaticalErrors", "activeWritingStyle", "proofState",
            "formsDesign", "attachedTemplate", "linkStyles", "stylePaneFormatFilter", "stylePaneSortMethod", "documentType", "mailMerge",
            "revisionView", "trackRevisions", "doNotTrackMoves", "doNotTrackFormatting", "documentProtection", "autoFormatOverride",
            "styleLockTheme", "styleLockQFSet", "defaultTabStop", "autoHyphenation", "consecutiveHyphenLimit", "hyphenationZone",
            "doNotHyphenateCaps", "showEnvelope", "summaryLength", "clickAndTypeStyle", "defaultTableStyle", "evenAndOddHeaders",
            "bookFoldRevPrinting", "bookFoldPrinting", "bookFoldPrintingSheets", "drawingGridHorizontalSpacing", "drawingGridVerticalSpacing",
            "displayHorizontalDrawingGridEvery", "displayVerticalDrawingGridEvery", "doNotUseMarginsForDrawingGridOrigin",
            "drawingGridHorizontalOrigin", "drawingGridVerticalOrigin", "doNotShadeFormData", "noPunctuationKerning",
            "characterSpacingControl", "printTwoOnOne", "strictFirstAndLastChars", "noLineBreaksAfter", "noLineBreaksBefore",
            "savePreviewPicture", "doNotValidateAgainstSchema", "saveInvalidXml", "ignoreMixedContent", "alwaysShowPlaceholderText",
            "doNotDemarcateInvalidXml", "saveXmlDataOnly", "useXSLTWhenSaving", "saveThroughXslt", "showXMLTags",
            "alwaysMergeEmptyNamespace", "updateFields", "hdrShapeDefaults", "footnotePr", "endnotePr", "compat", "docVars", "rsids",
            "mathPr", "attachedSchema", "themeFontLang", "clrSchemeMapping", "doNotIncludeSubdocsInStats", "doNotAutoCompressPictures",
            "forceUpgrade", "captions", "readModeInkLockDown", "smartTagType", "schemaLibrary", "shapeDefaults", "doNotEmbedSmartTags",
            "decimalSymbol", "listSeparator");

        internal static OfficeXmlChildOrder Order(params string[] names) {
            var namespaces = new int[names.Length];
            for (int index = 0; index < names.Length; index++) namespaces[index] = W;
            return new OfficeXmlChildOrder(namespaces, names);
        }

        internal static string Text(int value) => value.ToString(CultureInfo.InvariantCulture);

        internal static string Text(long value) => value.ToString(CultureInfo.InvariantCulture);

        internal static string Text(uint value) => value.ToString(CultureInfo.InvariantCulture);

        internal static string Text(bool value) => value ? "true" : "false";

        /// <summary>Creates a <c>w</c> element.</summary>
        internal static int Element(OfficeXmlTree tree, string localName) => tree.CreateElement(W, localName);

        /// <summary>Creates a <c>w</c> element with a <c>w:val</c> attribute.</summary>
        internal static int Element(OfficeXmlTree tree, string localName, string value) {
            int element = tree.CreateElement(W, localName);
            tree.SetAttribute(element, W, "val", value);
            return element;
        }

        /// <summary>Creates an on/off property: no attribute when on, <c>w:val="false"</c> when off.</summary>
        internal static int OnOff(OfficeXmlTree tree, string localName, bool enabled) {
            int element = tree.CreateElement(W, localName);
            if (!enabled) tree.SetAttribute(element, W, "val", "false");
            return element;
        }

        /// <summary>Creates an on/off property that always states its value, as the SDK writes an explicit boolean.</summary>
        internal static int OnOffValue(OfficeXmlTree tree, string localName, bool enabled) =>
            Element(tree, localName, enabled ? "true" : "false");

        internal static void Set(OfficeXmlTree tree, int element, string localName, string? value) {
            if (value != null) tree.SetAttribute(element, W, localName, value);
        }

        internal static string? Get(OfficeXmlTree tree, int element, string localName) => tree.GetAttribute(element, W, localName);

        internal static int Child(OfficeXmlTree tree, int parent, string localName) => tree.FindChild(parent, W, localName);

        internal static bool Is(OfficeXmlTree tree, int node, string localName) => tree.IsElement(node, W, localName);

        /// <summary>Appends <paramref name="child"/> and returns it.</summary>
        internal static int Append(OfficeXmlTree tree, int parent, int child) {
            tree.Append(parent, child);
            return child;
        }

        /// <summary>Inserts <paramref name="child"/> where <paramref name="order"/> puts it, or appends it when the order doesn't list it.</summary>
        internal static int Insert(OfficeXmlTree tree, int parent, int child, OfficeXmlChildOrder order) {
            if (order.RankOf(tree.NamespaceOf(child), tree.LocalNameOf(child)) < 0) tree.Append(parent, child);
            else tree.InsertOrdered(parent, child, order);
            return child;
        }

        /// <summary>Replaces every child with the same name as <paramref name="child"/>, then inserts it in schema order.</summary>
        internal static int Replace(OfficeXmlTree tree, int parent, int child, OfficeXmlChildOrder order) {
            RemoveAll(tree, parent, tree.LocalNameOf(child));
            return Insert(tree, parent, child, order);
        }

        internal static int GetOrAdd(OfficeXmlTree tree, int parent, string localName, OfficeXmlChildOrder order) {
            int child = tree.FindChild(parent, W, localName);
            return child != OfficeXmlTree.NoNode ? child : Insert(tree, parent, tree.CreateElement(W, localName), order);
        }

        internal static void RemoveAll(OfficeXmlTree tree, int parent, string localName) {
            int child = tree.FirstChildOf(parent);
            while (child != OfficeXmlTree.NoNode) {
                int next = tree.NextSiblingOf(child);
                if (tree.IsElement(child, W, localName)) tree.Remove(child);
                child = next;
            }
        }

        internal static void RemoveChildren(OfficeXmlTree tree, int parent) {
            int child = tree.FirstChildOf(parent);
            while (child != OfficeXmlTree.NoNode) {
                int next = tree.NextSiblingOf(child);
                tree.Remove(child);
                child = next;
            }
        }

        /// <summary>The element children of <paramref name="parent"/>, as a snapshot that survives edits.</summary>
        internal static List<int> Elements(OfficeXmlTree tree, int parent) {
            var children = new List<int>();
            for (int child = tree.FirstChildOf(parent); child != OfficeXmlTree.NoNode; child = tree.NextSiblingOf(child)) {
                if (tree.KindOf(child) == OfficeXmlTreeNodeKind.Element) children.Add(child);
            }
            return children;
        }

        internal static List<int> Elements(OfficeXmlTree tree, int parent, string localName) {
            var children = new List<int>();
            for (int child = tree.FirstChildOf(parent); child != OfficeXmlTree.NoNode; child = tree.NextSiblingOf(child)) {
                if (tree.IsElement(child, W, localName)) children.Add(child);
            }
            return children;
        }


        internal static bool HasElements(OfficeXmlTree tree, int parent) {
            for (int child = tree.FirstChildOf(parent); child != OfficeXmlTree.NoNode; child = tree.NextSiblingOf(child)) {
                if (tree.KindOf(child) == OfficeXmlTreeNodeKind.Element) return true;
            }
            return false;
        }

        /// <summary>The descendants of <paramref name="root"/> in document order, as a snapshot.</summary>
        internal static List<int> Descendants(OfficeXmlTree tree, int root) {
            var result = new List<int>();
            var stack = new Stack<int>();
            for (int child = tree.LastChildOf(root); child != OfficeXmlTree.NoNode; child = tree.PreviousSiblingOf(child)) stack.Push(child);
            while (stack.Count > 0) {
                int node = stack.Pop();
                if (tree.KindOf(node) != OfficeXmlTreeNodeKind.Element) continue;
                result.Add(node);
                for (int child = tree.LastChildOf(node); child != OfficeXmlTree.NoNode; child = tree.PreviousSiblingOf(child)) stack.Push(child);
            }
            return result;
        }

        /// <summary>Creates <c>w:t</c> (or <paramref name="localName"/>) with the text and <c>xml:space="preserve"</c>.</summary>
        internal static int TextElement(OfficeXmlTree tree, string text, string localName = "t") {
            int element = tree.CreateElement(W, localName);
            tree.SetAttribute(element, OfficeXmlNamespaces.Xml, "space", "preserve");
            tree.Append(element, tree.CreateText(text));
            return element;
        }

    }
}

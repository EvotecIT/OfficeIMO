using System.Globalization;
using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.Word.OpenXml;

namespace OfficeIMO.Word {
    /// <summary>
    /// A paragraph of a document held by OfficeIMO's own Open XML layer points at its elements in the owned tree. Text
    /// and basic run formatting are read and written there, with the same results the SDK path gives. Members that need
    /// more resolve the SDK elements, which opens the SDK package for the whole document.
    /// </summary>
    public partial class WordParagraph {
        private const int W = OfficeXmlNamespaces.WordprocessingML;
        private const int NoNode = OfficeXmlTree.NoNode;

        // Set while the paragraph lives in an owned tree; cleared once its SDK elements are resolved. The child is the
        // run (a node index) or the bookmark (its complement), or NoNode. The ordinals locate the same elements by
        // position: the paragraph among the body's paragraphs, the child among the paragraph's runs or bookmarks in
        // document order. Owned edits only append, so the ordinals stay valid for a copy of the package parsed again.
        internal OfficeXmlTree? _ownedTree;
        internal int _ownedParagraphNode = NoNode;
        private int _ownedChildNode = NoNode;
        private int _ownedParagraphOrdinal;
        private int _ownedChildOrdinal = -1;

        internal int _ownedRunNode => _ownedChildNode >= 0 ? _ownedChildNode : NoNode;
        internal int _ownedBookmarkNode => _ownedChildNode < NoNode ? ~_ownedChildNode : NoNode;

        internal static WordParagraph CreateOwned(WordDocument document, OfficeXmlTree tree, int paragraph, int paragraphOrdinal,
            int run = NoNode, int bookmark = NoNode) {
            var wordParagraph = new WordParagraph(document, newParagraph: false, newRun: false) {
                _ownedTree = tree,
                _ownedParagraphNode = paragraph,
                _ownedParagraphOrdinal = paragraphOrdinal,
            };
            if (run != NoNode) wordParagraph.SetOwnedChild(run, "r");
            else if (bookmark != NoNode) wordParagraph.SetOwnedChild(bookmark, "bookmarkStart");
            if (document._ownedModel is { } model) model.ParagraphsHandedOut = true;
            return wordParagraph;
        }

        private void SetOwnedChild(int node, string localName) {
            _ownedChildNode = localName == "r" ? node : ~node;
            _ownedChildOrdinal = OwnedOrdinal(_ownedTree!, _ownedParagraphNode, node, localName);
        }

        // The number of elements with this name before node among the descendants of root, in document order.
        private static int OwnedOrdinal(OfficeXmlTree tree, int root, int node, string localName) {
            int ordinal = 0;
            return Walk(tree.FirstChildOf(root)) ? ordinal : -1;

            bool Walk(int child) {
                for (; child != NoNode; child = tree.NextSiblingOf(child)) {
                    if (tree.KindOf(child) != OfficeXmlTreeNodeKind.Element) continue;
                    if (child == node) return true;
                    if (tree.IsElement(child, W, localName)) ordinal++;
                    if (Walk(tree.FirstChildOf(child))) return true;
                }
                return false;
            }
        }

        // The element at that ordinal, the inverse of OwnedOrdinal.
        private static int OwnedNodeAt(OfficeXmlTree tree, int root, int ordinal, string localName) {
            int remaining = ordinal;
            return Walk(tree.FirstChildOf(root));

            int Walk(int child) {
                for (; child != NoNode; child = tree.NextSiblingOf(child)) {
                    if (tree.KindOf(child) != OfficeXmlTreeNodeKind.Element) continue;
                    if (tree.IsElement(child, W, localName) && remaining-- == 0) return child;
                    int found = Walk(tree.FirstChildOf(child));
                    if (found != NoNode) return found;
                }
                return NoNode;
            }
        }

        /// <summary>Points the paragraph at the same elements in <paramref name="model"/>, a copy of its package parsed again.</summary>
        private bool RebindOwned(WordPackageModel model) {
            OfficeXmlTree tree = model.Document;
            int paragraph = NoNode;
            int seen = 0;
            for (int child = tree.FirstChildOf(model.Body); child != NoNode; child = tree.NextSiblingOf(child)) {
                if (!tree.IsElement(child, W, "p")) continue;
                if (seen++ == _ownedParagraphOrdinal) {
                    paragraph = child;
                    break;
                }
            }
            if (paragraph == NoNode) return false;
            int childNode = NoNode;
            if (_ownedChildNode != NoNode) {
                bool run = _ownedChildNode >= 0;
                int found = OwnedNodeAt(tree, paragraph, _ownedChildOrdinal, run ? "r" : "bookmarkStart");
                if (found == NoNode) return false;
                childNode = run ? found : ~found;
            }
            _ownedTree = tree;
            _ownedParagraphNode = paragraph;
            _ownedChildNode = childNode;
            return true;
        }

        /// <summary>
        /// The owned model this paragraph can still be read and edited through, or null when it has to use the SDK. A
        /// paragraph of a disposed document reads a copy of the package as it was when the document was disposed.
        /// </summary>
        private WordPackageModel? OwnedModel {
            get {
                OfficeXmlTree? tree = _ownedTree;
                if (tree == null || _document == null) return null;
                WordPackageModel? model = _document._ownedModel;
                if (model != null) return model.IsDocumentParsed && ReferenceEquals(model.Document, tree) ? model : null;
                WordPackageModel? disposed = _document.DisposedReadModel;
                if (disposed == null) return null;
                if (ReferenceEquals(disposed.Document, tree) || RebindOwned(disposed)) return disposed;
                return null;
            }
        }

        /// <summary>Rebinds a paragraph of a disposed document to the copy its document keeps, for the SDK path to resolve.</summary>
        internal bool TryRebindToDisposedCopy() => _ownedTree != null && _document?.DisposedReadModel is { } disposed
            && (ReferenceEquals(disposed.Document, _ownedTree) || RebindOwned(disposed));

        // ---- Text --------------------------------------------------------------------------------------------------

        private bool TryGetOwnedText(out string text) {
            text = string.Empty;
            if (OwnedModel == null) return false;
            // A paragraph without a run (an empty paragraph or a bookmark) has no text, as on the SDK path.
            if (_ownedRunNode == NoNode) return true;
            return WordOwnedText.TryRead(_ownedTree!, _ownedRunNode, out text);
        }

        private bool TrySetOwnedText(string? value) {
            WordPackageModel? model = OwnedModel;
            if (model == null) return false;
            OfficeXmlTree tree = _ownedTree!;
            if (_ownedRunNode != NoNode && !WordOwnedText.CanWrite(tree, _ownedRunNode)) return false;
            WordOwnedEdits.PrepareBodyEdit(model);
            int run = EnsureOwnedRun(tree);
            WordOwnedText.Write(tree, run, value ?? string.Empty);
            model.MarkChanged();
            return true;
        }

        private int EnsureOwnedRun(OfficeXmlTree tree) {
            if (_ownedRunNode == NoNode) {
                int run = tree.CreateElement(W, "r");
                tree.Append(_ownedParagraphNode, run);
                SetOwnedChild(run, "r");
            }
            return _ownedRunNode;
        }

        /// <summary>Appends a run with <paramref name="text"/>, as <see cref="AddText"/> does on the SDK path.</summary>
        private WordParagraph? TryAddOwnedText(string text) {
            WordPackageModel? model = OwnedModel;
            if (model == null) return null;
            OfficeXmlTree tree = _ownedTree!;
            WordOwnedEdits.PrepareBodyEdit(model);
            if (text.IndexOf('\n') < 0) {
                int run = tree.CreateElement(W, "r");
                WordOwnedText.Write(tree, run, text);
                tree.Append(_ownedParagraphNode, run);
                model.MarkChanged();
                return CreateOwned(_document, tree, _ownedParagraphNode, _ownedParagraphOrdinal, run);
            }
            // Lines become runs separated by text-wrapping breaks, each break in a run of its own.
            string[] lines = text.Split(new[] { "\r\n", "\n" }, StringSplitOptions.None);
            WordParagraph last = this;
            for (int index = 0; index < lines.Length; index++) {
                if (index > 0) {
                    int breakRun = tree.CreateElement(W, "r");
                    tree.Append(breakRun, tree.CreateElement(W, "br"));
                    tree.Append(_ownedParagraphNode, breakRun);
                    last = CreateOwned(_document, tree, _ownedParagraphNode, _ownedParagraphOrdinal, breakRun);
                }
                if (lines[index].Length > 0) {
                    int run = tree.CreateElement(W, "r");
                    WordOwnedText.Write(tree, run, lines[index]);
                    tree.Append(_ownedParagraphNode, run);
                    last = CreateOwned(_document, tree, _ownedParagraphNode, _ownedParagraphOrdinal, run);
                }
            }
            model.MarkChanged();
            return last;
        }

        // ---- Run properties ----------------------------------------------------------------------------------------

        /// <summary>The run properties when the paragraph is owned; <paramref name="properties"/> is NoNode when the run has none.</summary>
        private bool TryGetOwnedRunProperties(out OfficeXmlTree tree, out int properties) {
            tree = _ownedTree!;
            properties = NoNode;
            if (OwnedModel == null) return false;
            if (_ownedRunNode != NoNode) properties = tree.FindChild(_ownedRunNode, W, "rPr");
            return true;
        }

        /// <summary>The run properties to edit, created as the run's first child when missing, as the SDK path does.</summary>
        private bool TryEditOwnedRunProperties(out OfficeXmlTree tree, out int properties) {
            tree = _ownedTree!;
            properties = NoNode;
            WordPackageModel? model = OwnedModel;
            if (model == null) return false;
            WordOwnedEdits.PrepareBodyEdit(model);
            int run = EnsureOwnedRun(tree);
            properties = tree.FindChild(run, W, "rPr");
            if (properties == NoNode) {
                properties = tree.CreateElement(W, "rPr");
                tree.InsertBefore(run, properties, tree.FirstChildOf(run));
            }
            model.MarkChanged();
            return true;
        }

        private bool TryGetOwnedOnOff(string name, out bool value) {
            value = false;
            if (!TryGetOwnedRunProperties(out OfficeXmlTree tree, out int properties)) return false;
            if (properties == NoNode) return true;
            return WordOwnedEdits.TryReadOnOff(tree, tree.FindChild(properties, W, name), out value);
        }

        private bool TrySetOwnedOnOffPair(string name, string complexScriptName, bool value) {
            if (!TryEditOwnedRunProperties(out OfficeXmlTree tree, out int properties)) return false;
            string token = value ? "true" : "false";
            WordOwnedEdits.ReplaceChild(tree, properties, name, WRPrView.Order).SetValue(W, "val", token);
            WordOwnedEdits.ReplaceChild(tree, properties, complexScriptName, WRPrView.Order).SetValue(W, "val", token);
            return true;
        }

        private bool TrySetOwnedToggle(string name, bool value) {
            if (!TryEditOwnedRunProperties(out OfficeXmlTree tree, out int properties)) return false;
            if (value) WordOwnedEdits.ReplaceChild(tree, properties, name, WRPrView.Order);
            else WordOwnedEdits.RemoveChild(tree, properties, name);
            return true;
        }

        private bool TryGetOwnedValue(string name, out string? value) {
            value = null;
            if (!TryGetOwnedRunProperties(out OfficeXmlTree tree, out int properties)) return false;
            if (properties == NoNode) return true;
            int child = tree.FindChild(properties, W, name);
            if (child != NoNode) value = tree.GetAttribute(child, W, "val");
            return true;
        }

        private bool TrySetOwnedValue(string name, string? value) {
            if (!TryEditOwnedRunProperties(out OfficeXmlTree tree, out int properties)) return false;
            if (value == null) WordOwnedEdits.RemoveChild(tree, properties, name);
            else WordOwnedEdits.ReplaceChild(tree, properties, name, WRPrView.Order).SetValue(W, "val", value);
            return true;
        }

        private bool TryGetOwnedUnderline(out WordUnderlineStyle? value) {
            value = null;
            if (!TryGetOwnedValue("u", out string? token)) return false;
            if (!TryGetOwnedRunProperties(out OfficeXmlTree tree, out int properties)) return false;
            if (properties == NoNode || tree.FindChild(properties, W, "u") == NoNode) return true;
            if (token == null) return true;
            if (!WordUnderlineStyleTokens.TryParse(token, out WordUnderlineStyle style)) return false;
            value = style;
            return true;
        }

        private bool TrySetOwnedUnderline(WordUnderlineStyle? value) {
            if (!TryEditOwnedRunProperties(out OfficeXmlTree tree, out int properties)) return false;
            if (value == null) {
                WordOwnedEdits.RemoveChild(tree, properties, "u");
                return true;
            }
            // An existing underline keeps its other attributes, such as its colour.
            int underline = tree.GetOrInsertChild(properties, W, "u", WRPrView.Order);
            tree.SetAttribute(underline, W, "val", WordUnderlineStyleTokens.ToToken(value.Value));
            return true;
        }

        private bool TryGetOwnedHighlight(out WordHighlightColor? value) {
            value = null;
            if (!TryGetOwnedValue("highlight", out string? token)) return false;
            if (token == null) return true;
            if (!WordOpenXmlValueTypeExtensions.TryParseOpenXmlValue(token, out WordHighlightColor color)) return false;
            value = color;
            return true;
        }

        // Clearing the highlight leaves a run without properties as it is, as on the SDK path.
        private bool TrySetOwnedHighlight(WordHighlightColor? value) {
            if (value == null) {
                if (!TryGetOwnedRunProperties(out OfficeXmlTree tree, out int properties)) return false;
                if (properties != NoNode && tree.FindChild(properties, W, "highlight") != NoNode) {
                    WordOwnedEdits.RemoveChild(tree, properties, "highlight");
                    _document._ownedModel!.MarkChanged();
                }
                return true;
            }
            return TrySetOwnedValue("highlight", value.Value.ToOpenXmlToken());
        }

        private bool TryGetOwnedFontSize(out double? points) {
            points = null;
            if (!TryGetOwnedValue("sz", out string? token)) return false;
            if (int.TryParse(token, NumberStyles.Integer, CultureInfo.InvariantCulture, out int halfPoints)) points = halfPoints / 2D;
            return true;
        }

        private bool TryGetOwnedFont(string attribute, out string? value) {
            value = null;
            if (!TryGetOwnedRunProperties(out OfficeXmlTree tree, out int properties)) return false;
            if (properties == NoNode) return true;
            int fonts = tree.FindChild(properties, W, "rFonts");
            if (fonts != NoNode) value = tree.GetAttribute(fonts, W, attribute);
            return true;
        }

        private bool TrySetOwnedFontFamily(string? value) {
            if (!TryEditOwnedRunProperties(out OfficeXmlTree tree, out int properties)) return false;
            int fonts = tree.GetOrInsertChild(properties, W, "rFonts", WRPrView.Order);
            if (string.IsNullOrEmpty(value)) {
                tree.RemoveAttribute(fonts, W, "ascii");
            } else {
                tree.SetAttribute(fonts, W, "ascii", value!);
                tree.SetAttribute(fonts, W, "hAnsi", value!);
                tree.SetAttribute(fonts, W, "eastAsia", value!);
                tree.SetAttribute(fonts, W, "cs", value!);
            }
            return true;
        }

        private bool TryGetOwnedColor(out OfficeColor? value) {
            value = null;
            if (!TryGetOwnedRunProperties(out OfficeXmlTree tree, out int properties)) return false;
            if (properties == NoNode) return true;
            int color = tree.FindChild(properties, W, "color");
            if (color == NoNode) return true;
            // Theme colours resolve against the document's theme, which the SDK path reads.
            if (tree.TryGetAttribute(color, W, "themeColor", out _)) return false;
            value = OfficeColor.TryParseHex(tree.GetAttribute(color, W, "val"), out OfficeColor parsed) ? parsed : null;
            return true;
        }

        private bool TryGetOwnedCaps(out WordCapsStyle value) {
            value = WordCapsStyle.None;
            if (!TryGetOwnedOnOff("caps", out bool caps) || !TryGetOwnedOnOff("smallCaps", out bool smallCaps)) return false;
            value = caps ? WordCapsStyle.Caps : smallCaps ? WordCapsStyle.SmallCaps : WordCapsStyle.None;
            return true;
        }

        private bool TrySetOwnedCaps(WordCapsStyle value) {
            if (!TryEditOwnedRunProperties(out OfficeXmlTree tree, out int properties)) return false;
            if (value == WordCapsStyle.Caps) WordOwnedEdits.ReplaceChild(tree, properties, "caps", WRPrView.Order);
            else WordOwnedEdits.RemoveChild(tree, properties, "caps");
            if (value == WordCapsStyle.SmallCaps) WordOwnedEdits.ReplaceChild(tree, properties, "smallCaps", WRPrView.Order);
            else WordOwnedEdits.RemoveChild(tree, properties, "smallCaps");
            return true;
        }
    }

    /// <summary>The XML tokens of <see cref="WordUnderlineStyle"/>, in declaration order.</summary>
    internal static class WordUnderlineStyleTokens {
        private static readonly string[] Tokens = {
            "single", "words", "double", "thick", "dotted", "dottedHeavy", "dash", "dashedHeavy", "dashLong", "dashLongHeavy",
            "dotDash", "dashDotHeavy", "dotDotDash", "dashDotDotHeavy", "wave", "wavyHeavy", "wavyDouble", "none"
        };

        internal static string ToToken(WordUnderlineStyle style) => Tokens[(int)style];

        internal static bool TryParse(string token, out WordUnderlineStyle style) {
            int index = Array.IndexOf(Tokens, token);
            style = index < 0 ? default : (WordUnderlineStyle)index;
            return index >= 0;
        }
    }
}

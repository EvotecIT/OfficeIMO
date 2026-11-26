using System.Globalization;
using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.OpenXml.Internal;
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

        // Set while the paragraph lives in an owned tree; cleared once its SDK elements are resolved. A paragraph wraps
        // its run, its bookmark, or neither; a bookmark paragraph keeps its bookmark when text gives it a run.
        internal OfficeXmlTree? _ownedTree;
        // CreateOwned sets the nodes along with the tree; they're read only while the tree is set. They have no
        // initializers, which would run in every constructor.
        internal int _ownedParagraphNode;
        internal int _ownedRunNode;
        internal int _ownedBookmarkNode;

        internal static WordParagraph CreateOwned(WordDocument document, OfficeXmlTree tree, int paragraph,
            int run = NoNode, int bookmark = NoNode) {
            var wordParagraph = new WordParagraph(document, newParagraph: false, newRun: false) {
                _ownedTree = tree,
                _ownedParagraphNode = paragraph,
                _ownedRunNode = run,
                _ownedBookmarkNode = bookmark,
            };
            if (document._ownedModel is { } model) model.ParagraphsHandedOut = true;
            return wordParagraph;
        }

        /// <summary>Points the paragraph at the same elements in the copy of the package a disposed document keeps.</summary>
        private bool RebindOwned(WordPackageModel copy) {
            int paragraph = _document.DisposedCopyNode(_ownedParagraphNode);
            if (paragraph == NoNode) return false;
            int run = NoNode, bookmark = NoNode;
            if (_ownedRunNode != NoNode && (run = _document.DisposedCopyNode(_ownedRunNode)) == NoNode) return false;
            if (_ownedBookmarkNode != NoNode && (bookmark = _document.DisposedCopyNode(_ownedBookmarkNode)) == NoNode) return false;
            _ownedTree = copy.Document;
            _ownedParagraphNode = paragraph;
            _ownedRunNode = run;
            _ownedBookmarkNode = bookmark;
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
                // Once the SDK package is open, it is the document for every paragraph, the kept copy included.
                if (_document._ownedNodeElements != null) return null;
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
                _ownedRunNode = run;
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
                return CreateOwned(_document, tree, _ownedParagraphNode, run);
            }
            // Lines become runs separated by text-wrapping breaks, each break in a run of its own.
            string[] lines = text.Split(new[] { "\r\n", "\n" }, StringSplitOptions.None);
            WordParagraph last = this;
            for (int index = 0; index < lines.Length; index++) {
                if (index > 0) {
                    int breakRun = tree.CreateElement(W, "r");
                    tree.Append(breakRun, tree.CreateElement(W, "br"));
                    tree.Append(_ownedParagraphNode, breakRun);
                    last = CreateOwned(_document, tree, _ownedParagraphNode, breakRun);
                }
                if (lines[index].Length > 0) {
                    int run = tree.CreateElement(W, "r");
                    WordOwnedText.Write(tree, run, lines[index]);
                    tree.Append(_ownedParagraphNode, run);
                    last = CreateOwned(_document, tree, _ownedParagraphNode, run);
                }
            }
            model.MarkChanged();
            return last;
        }

        // ---- Paragraph properties ----------------------------------------------------------------------------------

        /// <summary>The paragraph properties when the paragraph is owned; <paramref name="properties"/> is NoNode when it has none.</summary>
        private bool TryGetOwnedParagraphProperties(out OfficeXmlTree tree, out int properties) {
            properties = NoNode;
            if (OwnedModel == null) {
                tree = null!;
                return false;
            }
            tree = _ownedTree!;
            properties = tree.FindChild(_ownedParagraphNode, W, "pPr");
            return true;
        }

        /// <summary>The paragraph properties to edit, created as the paragraph's first child when missing, as the SDK path does.</summary>
        private bool TryEditOwnedParagraphProperties(out OfficeXmlTree tree, out int properties) {
            properties = NoNode;
            WordPackageModel? model = OwnedModel;
            if (model == null) {
                tree = null!;
                return false;
            }
            tree = _ownedTree!;
            WordOwnedEdits.PrepareBodyEdit(model);
            properties = tree.GetOrInsertChild(_ownedParagraphNode, W, "pPr", WPView.Order);
            model.MarkChanged();
            return true;
        }

        private bool TryGetOwnedStyleId(out string? styleId) {
            styleId = null;
            if (!TryGetOwnedParagraphProperties(out OfficeXmlTree tree, out int properties)) return false;
            if (properties == NoNode) return true;
            int style = tree.FindChild(properties, W, "pStyle");
            if (style != NoNode) styleId = tree.GetAttribute(style, W, "val");
            return true;
        }

        private bool TryGetOwnedStyle(out WordParagraphStyles? style) {
            style = null;
            if (!TryGetOwnedStyleId(out string? styleId)) return false;
            if (styleId != null) style = WordParagraphStyle.GetStyle(OwnedModel!.Styles.GetBuiltInStyleId(styleId) ?? styleId);
            return true;
        }

        /// <summary>
        /// Applies a built-in style when the document already defines it, with what it depends on; a style to add, or a
        /// heading that may have to refresh a table of contents, is left to the SDK path.
        /// </summary>
        private bool TrySetOwnedStyle(WordParagraphStyles? value) {
            WordPackageModel? model = OwnedModel;
            if (model == null) return false;
            if (value == null) return true;
            string? styleId = _document.TryUseBuiltInStyleWithoutSdk(model, value.Value.ToStringStyle());
            if (styleId == null) return false;
            bool heading = value.Value >= WordParagraphStyles.Heading1 && value.Value <= WordParagraphStyles.Heading9;
            if (heading && !_document.TryHeadingModifiedWithoutSdk(model)) return false;
            if (!TryEditOwnedParagraphProperties(out OfficeXmlTree tree, out int properties)) return false;
            int style = tree.GetOrInsertChild(properties, W, "pStyle", WPPrView.Order);
            tree.SetAttribute(style, W, "val", styleId);
            return true;
        }

        // The paragraph formatting members in WordParagraph.ParagraphProperties.cs edit through these, on the paragraph
        // properties' element handle.

        private delegate bool TokenParser<TEnum>(string token, out TEnum value);

        /// <summary>A paragraph property's <c>w:val</c> token, mapped as the SDK path maps it.</summary>
        private TEnum? ParagraphTokenValue<TEnum>(string name, TokenParser<TEnum> parse) where TEnum : struct, Enum {
            string? token = WordXmlEdits.ChildValue(ParagraphPropertiesNode, name);
            return token == null ? null : WordXmlEdits.Token(parse(token, out TEnum value), value, token);
        }

        /// <summary>Replaces a paragraph property with one holding <paramref name="token"/>; null removes every such property.</summary>
        private void SetParagraphToken(string name, string? token) {
            if (token == null) {
                if (ParagraphPropertiesNode.Child(W, name).IsNone) return;
                EditParagraphProperties().RemoveChildren(W, name);
                return;
            }
            WordXmlEdits.ReplaceChild(EditParagraphProperties(), name, WPPrView.Order, token);
        }

        /// <summary>
        /// Reads a twips attribute of a paragraph property element, such as <c>w:ind/@w:left</c>: a value that doesn't
        /// parse as an integer reads as no value.
        /// </summary>
        private OfficeLength? GetParagraphTwips(string element, string attribute) {
            string? text = ParagraphPropertiesNode.Child(W, element).Attribute(W, attribute);
            return !string.IsNullOrEmpty(text) && int.TryParse(text, out int twips) ? OfficeLength.Points(twips / 20D) : null;
        }

        /// <summary>Sets or clears a twips attribute of a paragraph property element, adding the element when missing even to clear it.</summary>
        private void SetParagraphTwips(string element, string attribute, string? text) =>
            EditParagraphProperties().GetOrAddChild(W, element, WPPrView.Order).SetAttribute(W, attribute, text);

        /// <summary>Turns a paragraph on/off property on by keeping its element without a value, or off by removing it.</summary>
        private void SetParagraphFlag(string name, bool value) {
            OfficeXmlNode properties = EditParagraphProperties();
            if (value) properties.GetOrAddChild(W, name, WPPrView.Order).RemoveAttribute(W, "val");
            else WordXmlEdits.RemoveChild(properties, name);
        }

        // Widow control is on unless it says otherwise; turning it off writes an explicit false.
        private void SetWidowControl(bool value) {
            if (value) SetParagraphFlag("widowControl", true);
            else WordXmlEdits.ReplaceChild(EditParagraphProperties(), "widowControl", WPPrView.Order, "false");
        }

        // Page break before is replaced by a bare element when set.
        private void SetPageBreakBefore(bool value) {
            OfficeXmlNode properties = EditParagraphProperties();
            if (value) WordXmlEdits.ReplaceChild(properties, "pageBreakBefore", WPPrView.Order);
            else WordXmlEdits.RemoveChild(properties, "pageBreakBefore");
        }

        // Turning right-to-left on keeps an existing element as it is.
        private void SetBiDi(bool value) {
            OfficeXmlNode properties = EditParagraphProperties();
            if (value) properties.GetOrAddChild(W, "bidi", WPPrView.Order);
            else WordXmlEdits.RemoveChild(properties, "bidi");
        }

        /// <summary>The <c>w:spacing</c> line attributes, as written.</summary>
        private void GetLineAttributes(out string? line, out string? rule) {
            OfficeXmlNode spacing = ParagraphPropertiesNode.Child(W, "spacing");
            line = spacing.Attribute(W, "line");
            rule = spacing.Attribute(W, "lineRule");
        }

        /// <summary>
        /// Sets or clears <c>w:spacing</c> line attributes. Clearing both leaves a paragraph without spacing as it is;
        /// anything else adds the paragraph properties and the spacing element when missing.
        /// </summary>
        private void SetLineAttributes(bool setLine, string? line, bool setRule, string? rule) {
            if (line == null && rule == null && (setLine || setRule)) {
                OfficeXmlNode existing = ParagraphPropertiesNode.Child(W, "spacing");
                if (!existing.IsNone) {
                    EditParagraphProperties();
                    // Cleared even when no value is written, so an SDK element drops a typed value that holds none.
                    if (setLine) existing.SetAttribute(W, "line", null);
                    if (setRule) existing.SetAttribute(W, "lineRule", null);
                    return;
                }
                if (setLine && setRule) return;
            }
            OfficeXmlNode spacing = EditParagraphProperties().GetOrAddChild(W, "spacing", WPPrView.Order);
            if (setLine) spacing.SetAttribute(W, "line", line);
            if (setRule) spacing.SetAttribute(W, "lineRule", rule);
        }

        // Indentation is written with the invariant culture, spacing with the current one, as the SDK path always did.
        private static string? IndentationText(OfficeLength? value) =>
            checked((int?)value?.ToTwips())?.ToString(CultureInfo.InvariantCulture);

        // Clearing spacing writes an empty value rather than removing the attribute.
        private static string SpacingText(OfficeLength? value) => checked((int?)value?.ToTwips())?.ToString() ?? string.Empty;
    }

    /// <summary>The XML tokens of <see cref="WordUnderlineStyle"/>, in declaration order.</summary>
    internal static class WordParagraphAlignmentTokens {
        private static readonly string[] Tokens = {
            "left", "start", "center", "right", "end", "both", "mediumKashida", "distribute", "numTab", "highKashida",
            "lowKashida", "thaiDistribute"
        };

        internal static string ToToken(WordParagraphAlignment alignment) => Tokens[(int)alignment];

        internal static bool TryParse(string token, out WordParagraphAlignment alignment) {
            int index = Array.IndexOf(Tokens, token);
            alignment = index < 0 ? default : (WordParagraphAlignment)index;
            return index >= 0;
        }
    }

    /// <summary>The <c>ST_TextDirection</c> tokens for <see cref="WordTextDirection"/>, in enum order.</summary>
    internal static class WordTextDirectionTokens {
        private static readonly string[] Tokens = {
            "lrTb", "tb", "tbRl", "rl", "btLr", "lr", "lrTbV", "tbV", "tbRlV", "rlV", "tbLrV", "lrV"
        };

        internal static string ToToken(WordTextDirection direction) => Tokens[(int)direction];

        internal static bool TryParse(string token, out WordTextDirection direction) {
            int index = Array.IndexOf(Tokens, token);
            direction = (WordTextDirection)Math.Max(index, 0);
            return index >= 0;
        }
    }

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

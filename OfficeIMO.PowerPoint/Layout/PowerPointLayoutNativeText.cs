using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.OpenXml.Internal;
using OfficeIMO.PowerPoint.OpenXml;
using Ns = OfficeIMO.Core.Internal.OpenXml.OfficeXmlNamespaces;

namespace OfficeIMO.PowerPoint {
    /// <summary>
    /// Keeps native text bodies as layout inputs. Layout supplies geometry and missing theme defaults, while
    /// the authored paragraph/run nodes remain the ones returned to the caller, including across continuation.
    /// </summary>
    internal sealed class PowerPointLayoutNativeText {
        private const int D = Ns.DrawingML;
        private const int P = Ns.PresentationML;

        private readonly PowerPointLayoutNativeText _root;
        private IReadOnlyList<PowerPointParagraph>? _paragraphs;
        private bool _used;
        internal PowerPointLayoutNativeText(PowerPointTextBox box) { Box = box; _root = this; }
        private PowerPointLayoutNativeText(PowerPointLayoutNativeText root, IReadOnlyList<PowerPointParagraph> paragraphs) {
            Box = root.Box; _root = root._root; _paragraphs = paragraphs;
        }
        internal PowerPointTextBox Box { get; }
        internal IReadOnlyList<PowerPointParagraph> Paragraphs => _paragraphs ?? Box.Paragraphs;
        internal void Freeze() => _paragraphs ??= Box.Paragraphs;
        internal PowerPointLayoutNativeText Slice(IReadOnlyList<PowerPointParagraph> paragraphs) => new(_root, paragraphs);

        /// <summary>Repeating context gets independent native nodes; the original frame stays on its first page.</summary>
        internal PowerPointLayoutNativeText CopyForRepeat() {
            PowerPointTextBox box = PowerPointSlide.CopyLayoutTextBox(Box);
            OfficeXmlNode body = box.XmlElement.Child(P, "txBody");
            body.RemoveChildren(D, "p");
            foreach (var paragraph in Paragraphs) body.Append(paragraph.Element.Clone());
            var copy = new PowerPointLayoutNativeText(box);
            copy.Freeze();
            return copy;
        }

        internal static PowerPointLayoutParagraph Attach(PowerPointLayoutParagraph paragraph, PowerPointLayoutNativeText? body, int index = 0) {
            if (body != null && index < body.Paragraphs.Count) {
                paragraph.NativeBody = body;
                paragraph.NativeParagraph = body.Paragraphs[index];
            }
            return paragraph;
        }

        /// <summary>Reuses the first native frame, then moves remaining native paragraphs into continuation frames.</summary>
        internal static PowerPointTextBox Bind(PowerPointSlide slide, PowerPointTextBox generated,
            IReadOnlyList<PowerPointLayoutParagraph> planned, PowerPointLayoutNativeText? nativeBody = null) {
            var native = planned.Where(p => p.NativeParagraph != null).ToList();
            nativeBody ??= native.FirstOrDefault()?.NativeBody;
            if (nativeBody == null) return generated;
            var root = nativeBody._root;
            var defaults = generated.Paragraphs;
            var result = root._used ? generated : root.Box;
            var nodes = new List<OfficeXmlNode>();
            for (int i = 0; i < planned.Count; i++) {
                var actual = planned[i].NativeParagraph;
                OfficeXmlNode fallback = defaults[i].Element;
                if (actual == null) { nodes.Add(fallback.Clone()); continue; }
                OfficeXmlNode paragraph = actual.Element;
                slide.RehomeLayoutRelationships(paragraph, actual.OwnerPartRef);
                MergeMissing(paragraph.GetOrAddChild(D, "pPr", ATextParagraphView.Order), fallback.Child(D, "pPr"));
                OfficeXmlNode runDefaults = fallback.Child(D, "r").Child(D, "rPr");
                foreach (OfficeXmlNode child in paragraph.Children()) {
                    if (child.Is(D, "r") || child.Is(D, "fld")) MergeMissing(child.GetOrPrependChild(D, "rPr"), runDefaults);
                }
                if (paragraph.Child(D, "endParaRPr").IsNone && fallback.Child(D, "endParaRPr") is { Exists: true } end)
                    paragraph.Append(end.Clone());
                nodes.Add(paragraph);
            }
            if (planned.Count == 0) nodes.Add(defaults[0].Element);
            // Capture all references before detaching: the source frame can contain later continuation paragraphs.
            foreach (var node in nodes) if (node.Parent.Exists) node.Remove();
            OfficeXmlNode shape = result.XmlElement;
            OfficeXmlNode template = generated.XmlElement;
            if (root._used) {
                // The continuation has its own ID and geometry, but inherits the authored frame appearance.
                OfficeXmlNode source = root.Box.XmlElement;
                OfficeXmlNode geometry = shape.Child(P, "spPr").Child(D, "xfrm").Clone();
                OfficeXmlNode properties = Replace(shape, P, "spPr", source.Child(P, "spPr").Clone(), PShapeView.Order);
                Replace(properties, D, "xfrm", geometry, AShapePropertiesView.Order);
                OfficeXmlNode body = shape.Child(P, "txBody"), sourceBody = source.Child(P, "txBody");
                Replace(body, D, "bodyPr", sourceBody.Child(D, "bodyPr").Clone(), ATextBodyView.Order);
                Replace(body, D, "lstStyle", sourceBody.Child(D, "lstStyle").Clone(), ATextBodyView.Order);
            }
            if (!ReferenceEquals(result, generated)) {
                OfficeXmlNode properties = shape.GetOrAddChild(P, "spPr", PShapeView.Order);
                OfficeXmlNode templateProperties = template.Child(P, "spPr");
                Replace(properties, D, "xfrm", templateProperties.Child(D, "xfrm").Clone(), AShapePropertiesView.Order);
                MergeMissing(properties, templateProperties);
                OfficeXmlNode nonVisual = shape.Child(P, "nvSpPr"), templateNonVisual = template.Child(P, "nvSpPr");
                MergeMissing(nonVisual.Child(P, "cNvSpPr"), templateNonVisual.Child(P, "cNvSpPr"));
                MergeMissing(nonVisual.Child(P, "nvPr"), templateNonVisual.Child(P, "nvPr"));
                OfficeXmlNode body = shape.Child(P, "txBody"), templateBody = template.Child(P, "txBody");
                MergeMissing(body.GetOrAddChild(D, "bodyPr", ATextBodyView.Order), templateBody.Child(D, "bodyPr"));
                MergeMissing(body.GetOrAddChild(D, "lstStyle", ATextBodyView.Order), templateBody.Child(D, "lstStyle"));
                slide.RemoveShape(generated);
                slide.AdoptLayoutTextBox(result);
            }
            OfficeXmlNode textBody = result.XmlElement.Child(P, "txBody");
            textBody.RemoveChildren(D, "p");
            foreach (var node in nodes) textBody.Append(node);
            root._used = true;
            return result;
        }

        /// <summary>
        /// Replaces the child of that name with <paramref name="replacement"/> where it stands, or adds it in schema order,
        /// and returns it as it is now held.
        /// </summary>
        internal static OfficeXmlNode Replace(OfficeXmlNode parent, int namespaceId, string localName, OfficeXmlNode replacement, OfficeXmlChildOrder order) {
            OfficeXmlNode existing = parent.Child(namespaceId, localName);
            if (existing.IsNone) return parent.InsertOrdered(replacement, order);
            OfficeXmlNode inserted = existing.InsertAfterSelf(replacement);
            existing.Remove();
            return inserted;
        }

        /// <summary>
        /// Copies the attributes and children of <paramref name="defaults"/> that <paramref name="target"/> lacks: a child is
        /// missing when the target has none of its name and none of the same DrawingML choice.
        /// </summary>
        internal static void MergeMissing(OfficeXmlNode target, OfficeXmlNode defaults) {
            if (defaults.IsNone) return;
            (string[] attributes, OfficeXmlChildOrder order) = MergeShape(target.LocalName);
            foreach (string attribute in attributes) {
                if (defaults.Attribute(attribute) is { } value && !target.HasAttribute(attribute)) target.SetAttribute(attribute, value);
            }
            foreach (OfficeXmlNode child in defaults.Children().ToList()) {
                bool present = false;
                int choice = Choice(child);
                foreach (OfficeXmlNode existing in target.Children()) {
                    if (existing.LocalName == child.LocalName && existing.NamespaceUri == child.NamespaceUri || choice != 0 && Choice(existing) == choice) {
                        present = true;
                        break;
                    }
                }
                if (!present) target.InsertOrdered(child.Clone(), order);
            }
        }

        // The attributes the schema gives each element defaults are merged into, and its child order.
        private static (string[] Attributes, OfficeXmlChildOrder Order) MergeShape(string localName) => localName switch {
            "pPr" => (ParagraphAttributes, ATextParagraphPropertiesView.Order),
            "rPr" or "endParaRPr" or "defRPr" => (CharacterAttributes, ATextCharacterPropertiesView.Order),
            "spPr" => (new[] { "bwMode" }, AShapePropertiesView.Order),
            "cNvSpPr" => (new[] { "txBox" }, ANonVisualDrawingShapePropsView.Order),
            "nvPr" => (new[] { "isPhoto", "userDrawn" }, PApplicationNonVisualDrawingPropsView.Order),
            "bodyPr" => (BodyAttributes, ATextBodyPropertiesView.Order),
            "lstStyle" => (Array.Empty<string>(), ATextListStyleView.Order),
            "tcPr" => (new[] { "marL", "marR", "marT", "marB", "vert", "anchor", "anchorCtr", "horzOverflow" }, ATableCellPropertiesView.Order),
            _ => throw new ArgumentException("Layout defaults don't merge into '" + localName + "'.", nameof(localName))
        };

        private static readonly string[] ParagraphAttributes = {
            "marL", "marR", "lvl", "indent", "algn", "defTabSz", "rtl", "eaLnBrk", "fontAlgn", "latinLnBrk", "hangingPunct"
        };

        private static readonly string[] CharacterAttributes = {
            "kumimoji", "lang", "altLang", "sz", "b", "i", "u", "strike", "kern", "cap", "spc", "normalizeH", "baseline",
            "noProof", "dirty", "err", "smtClean", "smtId", "bmk"
        };

        private static readonly string[] BodyAttributes = {
            "rot", "spcFirstLastPara", "vertOverflow", "horzOverflow", "vert", "wrap", "lIns", "tIns", "rIns", "bIns", "numCol",
            "spcCol", "rtlCol", "fromWordArt", "anchor", "anchorCtr", "forceAA", "upright", "compatLnSpc"
        };

        // Native edits can select another member of a DrawingML choice, such as a character bullet on a numbered list.
        private static int Choice(OfficeXmlNode element) => element.NamespaceId != D ? 0 : element.LocalName switch {
            "buNone" or "buAutoNum" or "buChar" => 1,
            "buClr" or "buClrTx" => 2,
            "buFont" or "buFontTx" => 3,
            "buSzPts" or "buSzPct" or "buSzTx" => 4,
            "noFill" or "solidFill" or "gradFill" or "blipFill" or "pattFill" or "grpFill" => 5,
            "noAutofit" or "normAutofit" or "spAutoFit" => 6,
            "prstGeom" or "custGeom" => 7,
            "effectLst" or "effectDag" => 8,
            _ => 0
        };

        /// <summary>Splits a native paragraph at the same sentence boundaries as value-based layout input.</summary>
        internal static IReadOnlyList<PowerPointParagraph> SplitParagraph(PowerPointParagraph source, IReadOnlyList<string> pieces) {
            OfficeXmlNode sourceElement = source.Element;
            var inlines = sourceElement.Children().ToList().Where(IsInline).ToList();
            string full = source.Text;
            var ranges = new List<(int Start, int End)>();
            var result = new List<PowerPointParagraph>();
            bool contiguous = string.Concat(pieces) == full;
            int offset = 0;
            for (int i = 0; i < pieces.Count; i++) {
                var node = i == 0 ? sourceElement : sourceElement.Clone();
                if (i > 0) {
                    node.RemoveChildren(D, "r");
                    node.RemoveChildren(D, "fld");
                    node.RemoveChildren(D, "br");
                }
                result.Add(i == 0 ? source : new PowerPointParagraph(node, default, source.OwnerPartRef));
                ranges.Add((offset, offset + pieces[i].Length));
                offset += pieces[i].Length;
                if (!contiguous) while (offset < full.Length && char.IsWhiteSpace(full[offset])) offset++;
            }
            foreach (var inline in inlines) inline.Remove();
            int position = 0;
            foreach (var inline in inlines) {
                var template = inline.Clone();
                string? text = TextOf(template).IsNone ? null : TextOf(template).InnerText;
                int length = text?.Length ?? Environment.NewLine.Length;
                bool used = false;
                for (int i = 0; i < ranges.Count; i++) {
                    int start = Math.Max(ranges[i].Start, position), end = Math.Min(ranges[i].End, position + length);
                    if (end <= start) continue;
                    var fragment = used ? template.Clone() : inline;
                    OfficeXmlNode fragmentText = TextOf(fragment);
                    if (fragmentText.Exists) fragmentText.SetInnerText(text!.Substring(start - position, end - start));
                    var paragraph = result[i].Element;
                    var endProperties = paragraph.Child(D, "endParaRPr");
                    if (endProperties.IsNone) paragraph.Append(fragment);
                    else paragraph.InsertBefore(fragment, endProperties);
                    used = true;
                }
                position += length;
            }
            return result;
        }

        private static bool IsInline(OfficeXmlNode node) => node.Is(D, "r") || node.Is(D, "fld") || node.Is(D, "br");

        // The text (a:t) of a run or field; a break has none.
        private static OfficeXmlNode TextOf(OfficeXmlNode inline) =>
            inline.Is(D, "r") || inline.Is(D, "fld") ? inline.Child(D, "t") : default;
    }
}

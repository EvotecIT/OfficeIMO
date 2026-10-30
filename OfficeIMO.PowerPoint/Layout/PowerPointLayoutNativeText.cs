using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.PowerPoint {
    /// <summary>
    /// Keeps native text bodies as layout inputs. Layout supplies geometry and missing theme defaults, while
    /// the authored paragraph/run nodes remain the ones returned to the caller, including across continuation.
    /// </summary>
    internal sealed class PowerPointLayoutNativeText {
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
            var nodes = new List<A.Paragraph>();
            for (int i = 0; i < planned.Count; i++) {
                var actual = planned[i].NativeParagraph;
                var fallback = defaults[i].Paragraph;
                if (actual == null) { nodes.Add((A.Paragraph)fallback.CloneNode(true)); continue; }
                A.Paragraph paragraph = actual.Paragraph;
                var oldPart = actual.OwnerPart as OpenXmlPart;
                if (!ReferenceEquals(oldPart, slide.SlidePart)) RehomeLinks(paragraph, oldPart, slide.SlidePart);
                paragraph.ParagraphProperties ??= new A.ParagraphProperties();
                MergeMissing(paragraph.ParagraphProperties, fallback.ParagraphProperties);
                var runDefaults = fallback.Elements<A.Run>().FirstOrDefault()?.RunProperties;
                foreach (var child in paragraph.ChildElements) {
                    if (child is A.Run run) MergeMissing(run.RunProperties ??= new A.RunProperties(), runDefaults);
                    else if (child is A.Field field) MergeMissing(field.RunProperties ??= new A.RunProperties(), runDefaults);
                }
                if (paragraph.GetFirstChild<A.EndParagraphRunProperties>() == null && fallback.GetFirstChild<A.EndParagraphRunProperties>() is A.EndParagraphRunProperties end)
                    paragraph.Append(end.CloneNode(true));
                nodes.Add(paragraph);
            }
            if (planned.Count == 0) nodes.Add(defaults[0].Paragraph);
            // Capture all references before detaching: the source frame can contain later continuation paragraphs.
            foreach (var node in nodes) if (node.Parent != null) node.Remove();
            var shape = (Shape)result.Element;
            var template = (Shape)generated.Element;
            if (!ReferenceEquals(result, generated)) {
                shape.ShapeProperties = (ShapeProperties)template.ShapeProperties!.CloneNode(true);
                shape.NonVisualShapeProperties!.NonVisualShapeDrawingProperties =
                    (NonVisualShapeDrawingProperties)template.NonVisualShapeProperties!.NonVisualShapeDrawingProperties!.CloneNode(true);
                shape.NonVisualShapeProperties.ApplicationNonVisualDrawingProperties =
                    (ApplicationNonVisualDrawingProperties)template.NonVisualShapeProperties.ApplicationNonVisualDrawingProperties!.CloneNode(true);
                result.TextBody!.BodyProperties = (A.BodyProperties)generated.TextBody!.BodyProperties!.CloneNode(true);
                result.TextBody.ListStyle = (A.ListStyle)generated.TextBody.ListStyle!.CloneNode(true);
                slide.RemoveShape(generated);
                slide.AdoptLayoutTextBox(result);
            }
            result.TextBody!.RemoveAllChildren<A.Paragraph>();
            foreach (var node in nodes) result.TextBody.Append(node);
            root._used = true;
            return result;
        }

        private static void MergeMissing(OpenXmlCompositeElement target, OpenXmlCompositeElement? defaults) {
            if (defaults == null) return;
            foreach (var attribute in defaults.GetAttributes()) {
                if (!target.GetAttributes().Any(a => a.LocalName == attribute.LocalName && a.NamespaceUri == attribute.NamespaceUri))
                    target.SetAttribute(attribute);
            }
            foreach (var child in defaults.ChildElements) {
                if (!target.ChildElements.Any(c => c.LocalName == child.LocalName && c.NamespaceUri == child.NamespaceUri
                    || Choice(c) != 0 && Choice(c) == Choice(child)))
                    target.AddChild(child.CloneNode(true), true);
            }
        }

        // Native edits can select another member of a DrawingML choice, such as a character bullet on a numbered list.
        private static int Choice(OpenXmlElement element) => element switch {
            A.NoBullet or A.AutoNumberedBullet or A.CharacterBullet => 1,
            A.BulletColor or A.BulletColorText => 2,
            A.BulletFont or A.BulletFontText => 3,
            A.BulletSizePoints or A.BulletSizePercentage or A.BulletSizeText => 4,
            A.NoFill or A.SolidFill or A.GradientFill or A.BlipFill or A.PatternFill or A.GroupFill => 5,
            _ => 0
        };

        private static void RehomeLinks(A.Paragraph paragraph, OpenXmlPart? source, SlidePart target) {
            if (source == null) return;
            var ids = new Dictionary<string, string>(StringComparer.Ordinal);
            // Sounds have their own audio references, even when the enclosing hyperlink has no target ID.
            foreach (var sound in paragraph.Descendants<A.HyperlinkSound>()) {
                string? id = sound.Embed?.Value;
                if (string.IsNullOrEmpty(id)) continue;
                if (!ids.TryGetValue(id!, out string? replacement)) {
                    var media = PowerPointEmbeddedSound.Find(source, id);
                    if (media == null) continue;
                    replacement = target.AddAudioReferenceRelationship(media).Id;
                    ids.Add(id!, replacement);
                }
                sound.Embed = replacement;
            }
            foreach (var link in paragraph.Descendants<A.HyperlinkType>()) {
                string? id = link.Id?.Value;
                if (string.IsNullOrEmpty(id)) continue;
                if (!ids.TryGetValue(id!, out string? replacement)) {
                    var external = source.HyperlinkRelationships.FirstOrDefault(r => r.Id == id);
                    if (external != null) replacement = target.AddHyperlinkRelationship(external.Uri, external.IsExternal).Id;
                    else {
                        var part = source.Parts.FirstOrDefault(p => p.RelationshipId == id).OpenXmlPart;
                        if (part == null) continue;
                        target.AddPart(part);
                        replacement = target.GetIdOfPart(part);
                    }
                    ids.Add(id!, replacement);
                }
                link.Id = replacement;
            }
        }

        /// <summary>Splits a native paragraph at the same sentence boundaries as value-based layout input.</summary>
        internal static IReadOnlyList<PowerPointParagraph> SplitParagraph(PowerPointParagraph source, IReadOnlyList<string> pieces) {
            var inlines = source.Paragraph.ChildElements
                .Where(n => n is A.Run or A.Field or A.Break).ToList();
            string full = source.Text;
            var ranges = new List<(int Start, int End)>();
            var result = new List<PowerPointParagraph>();
            bool contiguous = string.Concat(pieces) == full;
            int offset = 0;
            for (int i = 0; i < pieces.Count; i++) {
                var node = i == 0 ? source.Paragraph : (A.Paragraph)source.Paragraph.CloneNode(true);
                if (i > 0) {
                    node.RemoveAllChildren<A.Run>();
                    node.RemoveAllChildren<A.Field>();
                    node.RemoveAllChildren<A.Break>();
                }
                result.Add(i == 0 ? source : new PowerPointParagraph(node, ownerPart: source.OwnerPart));
                ranges.Add((offset, offset + pieces[i].Length));
                offset += pieces[i].Length;
                if (!contiguous) while (offset < full.Length && char.IsWhiteSpace(full[offset])) offset++;
            }
            foreach (var inline in inlines) inline.Remove();
            int position = 0;
            foreach (var inline in inlines) {
                var template = inline.CloneNode(true);
                A.Text? text = (template as A.Run)?.Text ?? (template as A.Field)?.Text;
                int length = text?.Text.Length ?? Environment.NewLine.Length;
                bool used = false;
                for (int i = 0; i < ranges.Count; i++) {
                    int start = Math.Max(ranges[i].Start, position), end = Math.Min(ranges[i].End, position + length);
                    if (end <= start) continue;
                    var fragment = used ? template.CloneNode(true) : inline;
                    A.Text? fragmentText = (fragment as A.Run)?.Text ?? (fragment as A.Field)?.Text;
                    if (fragmentText != null) fragmentText.Text = text!.Text.Substring(start - position, end - start);
                    var paragraph = result[i].Paragraph;
                    var endProperties = paragraph.GetFirstChild<A.EndParagraphRunProperties>();
                    if (endProperties == null) paragraph.Append(fragment);
                    else paragraph.InsertBefore(fragment, endProperties);
                    used = true;
                }
                position += length;
            }
            return result;
        }
    }
}

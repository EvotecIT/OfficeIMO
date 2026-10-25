using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word.Rtf;

public static partial class WordRtfConverterExtensions {
    private static bool AppendWordRun(WordParagraph wordRun, RtfParagraph paragraph, ref RtfRun? previousRun, RtfDocument rtfDocument, Dictionary<string, int> revisionAuthorIndexes, RtfRevisionKind revisionKind = RtfRevisionKind.None, int? revisionAuthorIndex = null) {
        bool hasContent = false;
        var textChildren = new List<OpenXmlElement>();
        foreach (OpenXmlElement child in wordRun.EnumerateEffectiveRunContent()) {
            if (child is RunProperties) continue;

            if (child is Break breakNode) {
                hasContent |= AppendWordRunTextSpan(wordRun, textChildren, paragraph, ref previousRun, rtfDocument, revisionKind, revisionAuthorIndex);
                paragraph.AddBreak(ToRtfBreakKind(breakNode.Type?.Value));
                previousRun = null;
                hasContent = true;
                continue;
            }

            if (child is FootnoteReference || child is EndnoteReference) {
                hasContent |= AppendWordRunTextSpan(wordRun, textChildren, paragraph, ref previousRun, rtfDocument, revisionKind, revisionAuthorIndex);
                RtfNote note = child is FootnoteReference reference
                    ? CopyNote(new WordFootNote(wordRun._document, wordRun._paragraph, wordRun._run!, reference).Paragraphs, RtfNoteKind.Footnote, rtfDocument, revisionAuthorIndexes)
                    : CopyNote(new WordEndNote(wordRun._document, wordRun._paragraph, wordRun._run!, (EndnoteReference)child).Paragraphs, RtfNoteKind.Endnote, rtfDocument, revisionAuthorIndexes);
                if (previousRun != null && previousRun.Note == null) {
                    previousRun.Note = note;
                } else {
                    paragraph.AddNoteReference(note);
                }
                hasContent = true;
                continue;
            }

            // Retain attached source nodes for relationship and drawing ancestry resolution.
            var childView = new WordParagraph(wordRun._document, wordRun._paragraph, wordRun._run!) {
                _visibleRunSourceChildren = new[] { child }
            };
            List<WordImage> images = childView.EnumerateImages().ToList();
            if (images.Count == 0) {
                textChildren.Add(child);
                continue;
            }

            hasContent |= AppendWordRunTextSpan(wordRun, textChildren, paragraph, ref previousRun, rtfDocument, revisionKind, revisionAuthorIndex);
            foreach (WordImage wordImage in images) {
                RtfImage? image = CreateRtfImage(wordImage, out _, out _);
                if (image == null) continue;
                CopyImage(image, paragraph.AddImage(image.Format, image.Data));
                hasContent = true;
            }
            previousRun = null;
        }

        hasContent |= AppendWordRunTextSpan(wordRun, textChildren, paragraph, ref previousRun, rtfDocument, revisionKind, revisionAuthorIndex);
        return hasContent;
    }

    private static bool AppendWordRunTextSpan(WordParagraph source, List<OpenXmlElement> children, RtfParagraph paragraph, ref RtfRun? previousRun, RtfDocument rtfDocument, RtfRevisionKind revisionKind, int? revisionAuthorIndex) {
        if (children.Count == 0) return false;

        // Projection owns cloned text; the original run retains formatting and part identity.
        var visible = new Run();
        var sourceChildren = new List<OpenXmlElement>();
        if (source._run?.RunProperties is RunProperties properties) {
            visible.Append(properties.CloneNode(true));
            sourceChildren.Add(properties);
        }
        foreach (OpenXmlElement child in children) {
            visible.Append(child.CloneNode(true));
            sourceChildren.Add(child);
        }
        var span = new WordParagraph(source._document, source._paragraph, source._run!) {
            _hyperlink = source._hyperlink,
            _visibleRun = visible,
            _visibleRunSourceChildren = sourceChildren
        };
        children.Clear();
        return AppendWordRunText(span, paragraph, ref previousRun, rtfDocument, revisionKind, revisionAuthorIndex);
    }
}

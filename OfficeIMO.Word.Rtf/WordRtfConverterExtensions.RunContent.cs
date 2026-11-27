namespace OfficeIMO.Word.Rtf;

public static partial class WordRtfConverterExtensions {
    private static bool AppendWordRun(WordParagraph wordRun, RtfParagraph paragraph, ref RtfRun? previousRun, RtfDocument rtfDocument, Dictionary<string, int> revisionAuthorIndexes, RtfRevisionKind revisionKind = RtfRevisionKind.None, int? revisionAuthorIndex = null) {
        bool hasContent = false;
        var textChildren = new List<OfficeXmlNode>();
        foreach (OfficeXmlNode child in wordRun.EffectiveRunContentNodes()) {
            if (child.IsW("rPr")) continue;

            if (child.IsW("br")) {
                hasContent |= AppendWordRunTextSpan(wordRun, textChildren, paragraph, ref previousRun, rtfDocument, revisionKind, revisionAuthorIndex);
                paragraph.AddBreak(ToRtfBreakKind(child.WAttr("type")));
                previousRun = null;
                hasContent = true;
                continue;
            }

            if (child.IsW("footnoteReference") || child.IsW("endnoteReference")) {
                hasContent |= AppendWordRunTextSpan(wordRun, textChildren, paragraph, ref previousRun, rtfDocument, revisionKind, revisionAuthorIndex);
                RtfNote note = child.IsW("footnoteReference")
                    ? CopyNote(wordRun.FootNoteView(child).Paragraphs, RtfNoteKind.Footnote, rtfDocument, revisionAuthorIndexes)
                    : CopyNote(wordRun.EndNoteView(child).Paragraphs, RtfNoteKind.Endnote, rtfDocument, revisionAuthorIndexes);
                if (previousRun != null && previousRun.Note == null) {
                    previousRun.Note = note;
                } else {
                    paragraph.AddNoteReference(note);
                }
                hasContent = true;
                continue;
            }

            // Retain attached source nodes for relationship and drawing ancestry resolution.
            WordParagraph childView = wordRun._document.RunView(wordRun.Node(), wordRun.RunNode(),
                visibleRunSourceChildren: new[] { child });
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

    private static bool AppendWordRunTextSpan(WordParagraph source, List<OfficeXmlNode> children, RtfParagraph paragraph, ref RtfRun? previousRun, RtfDocument rtfDocument, RtfRevisionKind revisionKind, int? revisionAuthorIndex) {
        if (children.Count == 0) return false;

        // Projection owns cloned text; the original run retains formatting and part identity.
        OfficeXmlNode visible = WordConverterNodes.CreateDetachedRun();
        var sourceChildren = new List<OfficeXmlNode>();
        OfficeXmlNode properties = source.RunNode().WChild("rPr");
        if (properties.Exists) {
            visible.Append(properties.Clone());
            sourceChildren.Add(properties);
        }
        foreach (OfficeXmlNode child in children) {
            visible.Append(child.Clone());
            sourceChildren.Add(child);
        }
        WordParagraph span = source._document.RunView(source.Node(), source.RunNode(), source.HyperlinkNode(), visible, sourceChildren);
        children.Clear();
        return AppendWordRunText(span, paragraph, ref previousRun, rtfDocument, revisionKind, revisionAuthorIndex);
    }
}

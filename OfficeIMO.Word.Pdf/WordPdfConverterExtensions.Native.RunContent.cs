namespace OfficeIMO.Word.Pdf {

    public static partial class WordPdfConverterExtensions {
        private static bool IsNativeHiddenImageContent(WordImage image, WordParagraph paragraph) {
            OfficeXmlNode sourceRun = image.OccurrenceNode().Ancestor(WordXml.W, "r");
            return IsNativeHiddenTextRun(sourceRun.IsNone ? paragraph :
                paragraph._document.RunView(paragraph.Node(), sourceRun), paragraph);
        }

        private static void AppendNativeVisibleRunContent(List<WordParagraph> runs, WordParagraph source) {
            if (!source.IsImage) {
                runs.Add(source);
                return;
            }

            // A Word run can contain both text and drawings. Consumers that
            // dispatch image runs must see each in its original position,
            // without discarding the text sharing the source run.
            var pendingText = new List<OfficeXmlNode>();
            foreach (OfficeXmlNode child in source.EffectiveRunContentNodes()) {
                if (child.IsW("rPr")) continue;
                WordParagraph view = CreateNativeRunContentView(source, new[] { child });
                if (!view.IsImage) {
                    pendingText.Add(child);
                    continue;
                }
                if (pendingText.Count > 0) {
                    runs.Add(CreateNativeRunContentView(source, pendingText.ToArray()));
                    pendingText.Clear();
                }
                runs.Add(view);
            }
            if (pendingText.Count > 0)
                runs.Add(CreateNativeRunContentView(source, pendingText));
        }

        private static WordParagraph CreateNativeRunContentView(WordParagraph source, IReadOnlyList<OfficeXmlNode> children) {
            OfficeXmlNode paragraph = source.Node();
            OfficeXmlNode visible = paragraph.CreateElement(WordXml.W, "r");
            // Formatting is resolved against the original run. Keep the projected
            // children aligned with their source identities for positioned media.
            foreach (OfficeXmlNode child in children)
                visible.Append(child.Clone());
            return source._document.RunView(paragraph, source.RunNode(), source.HyperlinkNode(), visible, children);
        }
    }
}

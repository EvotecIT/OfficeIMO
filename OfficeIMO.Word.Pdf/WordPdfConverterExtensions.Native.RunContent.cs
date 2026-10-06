using DocumentFormat.OpenXml;
using System.Collections.Generic;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word.Pdf {
    public static partial class WordPdfConverterExtensions {
        private static bool IsNativeHiddenImageContent(WordImage image, WordParagraph paragraph) {
            OpenXmlElement? element = (OpenXmlElement?)image._Image ?? image._vmlShape;
            W.Run? sourceRun = element?.Ancestors<W.Run>().FirstOrDefault();
            return IsNativeHiddenTextRun(sourceRun == null ? paragraph :
                new WordParagraph(paragraph._document, paragraph._paragraph!, sourceRun), paragraph);
        }

        private static void AppendNativeVisibleRunContent(List<WordParagraph> runs, WordParagraph source) {
            if (!source.IsImage) {
                runs.Add(source);
                return;
            }

            // A Word run can contain both text and drawings. Consumers that
            // dispatch image runs must see each in its original position,
            // without discarding the text sharing the source run.
            var pendingText = new List<OpenXmlElement>();
            foreach (OpenXmlElement child in source.EnumerateEffectiveRunContent()) {
                if (child is W.RunProperties) continue;
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

        private static WordParagraph CreateNativeRunContentView(WordParagraph source, IReadOnlyList<OpenXmlElement> children) {
            var visible = new W.Run();
            // Formatting is resolved against the original run. Keep the projected
            // children aligned with their source identities for positioned media.
            foreach (OpenXmlElement child in children)
                visible.Append(child.CloneNode(true));
            return new WordParagraph(source._document, source._paragraph!, source._run!) {
                _hyperlink = source._hyperlink,
                _visibleRun = visible,
                _visibleRunSourceChildren = children
            };
        }
    }
}

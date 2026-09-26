using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    /// <summary>
    /// Tracks which run content is visible while walking a paragraph's complex fields.
    /// </summary>
    internal sealed class WordComplexFieldRunVisibility {
        private readonly Stack<bool> _resultRegions = new();

        internal WordComplexFieldRunVisibility(IEnumerable<bool>? outerToInnerResults = null) {
            if (outerToInnerResults != null)
                foreach (bool isResult in outerToInnerResults) _resultRegions.Push(isResult);
        }

        internal bool IsVisible => !_resultRegions.Contains(false);
        internal bool HasOpenField => _resultRegions.Count > 0;

        internal static WordComplexFieldRunVisibility ForParagraph(Paragraph paragraph) {
            var visibility = new WordComplexFieldRunVisibility();
            OpenXmlElement? textBoxStory = paragraph.Ancestors<TextBoxContent>().FirstOrDefault();
            OpenXmlElement story = textBoxStory ?? paragraph.Ancestors().FirstOrDefault(element =>
                element is Footnote or Endnote or Header or Footer)
                ?? paragraph.Ancestors().LastOrDefault() ?? paragraph;
            foreach (Paragraph earlier in story.Descendants<Paragraph>()) {
                if (ReferenceEquals(earlier, paragraph)) break;
                if (textBoxStory == null && earlier.Ancestors<TextBoxContent>().Any()) continue;
                foreach (FieldChar marker in earlier.Descendants<FieldChar>()
                    .Where(marker => ReferenceEquals(marker.Ancestors<Paragraph>().FirstOrDefault(), earlier)))
                    visibility.Observe(marker);
            }
            return visibility;
        }

        internal Run? GetVisibleRun(Run source) {
            if (!source.Elements<FieldChar>().Any()) return IsVisible ? source : null;

            var visible = new Run();
            bool hasContent = false;
            foreach (OpenXmlElement child in source.ChildElements) {
                if (child is FieldChar marker) {
                    Observe(marker);
                } else if (child is RunProperties) {
                    visible.Append(child.CloneNode(true));
                } else if (IsVisible && child is not FieldCode) {
                    visible.Append(child.CloneNode(true));
                    hasContent = true;
                }
            }

            return hasContent ? visible : null;
        }

        internal void ObserveDescendantRuns(OpenXmlElement container) {
            foreach (Run run in container.Descendants<Run>()) GetVisibleRun(run);
        }

        private void Observe(FieldChar marker) {
            if (marker.FieldCharType?.Value == FieldCharValues.Begin) {
                _resultRegions.Push(false);
            } else if (marker.FieldCharType?.Value == FieldCharValues.Separate && _resultRegions.Count > 0) {
                _resultRegions.Pop();
                _resultRegions.Push(true);
            } else if (marker.FieldCharType?.Value == FieldCharValues.End && _resultRegions.Count > 0) {
                _resultRegions.Pop();
            }
        }
    }
}

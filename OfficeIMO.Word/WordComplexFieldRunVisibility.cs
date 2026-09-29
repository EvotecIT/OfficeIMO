using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Threading;

namespace OfficeIMO.Word {
    /// <summary>
    /// Tracks which run content is visible while walking a paragraph's complex fields.
    /// </summary>
    internal sealed class WordComplexFieldRunVisibility {
        private static readonly AsyncLocal<Dictionary<OpenXmlElement, Dictionary<Paragraph, bool[]>>?> ConversionPrefixes = new();
        private readonly Stack<bool> _resultRegions = new();

        internal WordComplexFieldRunVisibility(IEnumerable<bool>? outerToInnerResults = null) {
            if (outerToInnerResults != null)
                foreach (bool isResult in outerToInnerResults) _resultRegions.Push(isResult);
        }

        internal bool IsVisible => !_resultRegions.Contains(false);
        internal bool HasOpenField => _resultRegions.Count > 0;

        internal static IDisposable BeginConversionScope() {
            Dictionary<OpenXmlElement, Dictionary<Paragraph, bool[]>>? previous = ConversionPrefixes.Value;
            ConversionPrefixes.Value = new Dictionary<OpenXmlElement, Dictionary<Paragraph, bool[]>>();
            return new ConversionScope(previous);
        }

        internal static WordComplexFieldRunVisibility ForParagraph(Paragraph paragraph) {
            OpenXmlElement? textBoxStory = paragraph.Ancestors<TextBoxContent>().FirstOrDefault();
            OpenXmlElement story = textBoxStory ?? paragraph.Ancestors().FirstOrDefault(element =>
                element is Footnote or Endnote or Header or Footer)
                ?? paragraph.Ancestors().LastOrDefault() ?? paragraph;
            Dictionary<OpenXmlElement, Dictionary<Paragraph, bool[]>>? cache = ConversionPrefixes.Value;
            if (cache != null) {
                if (!cache.TryGetValue(story, out Dictionary<Paragraph, bool[]>? prefixes)) {
                    prefixes = BuildStoryPrefixes(story, textBoxStory != null);
                    cache.Add(story, prefixes);
                }
                return prefixes.TryGetValue(paragraph, out bool[]? prefix)
                    ? new WordComplexFieldRunVisibility(prefix) : new WordComplexFieldRunVisibility();
            }

            var visibility = new WordComplexFieldRunVisibility();
            foreach (Paragraph earlier in story.Descendants<Paragraph>()) {
                if (ReferenceEquals(earlier, paragraph)) break;
                if (textBoxStory == null && earlier.Ancestors<TextBoxContent>().Any()) continue;
                visibility.ObserveParagraphMarkers(earlier);
            }
            return visibility;
        }

        private static Dictionary<Paragraph, bool[]> BuildStoryPrefixes(OpenXmlElement story, bool isTextBoxStory) {
            var prefixes = new Dictionary<Paragraph, bool[]>();
            var visibility = new WordComplexFieldRunVisibility();
            foreach (Paragraph paragraph in story.Descendants<Paragraph>()) {
                if (!isTextBoxStory && paragraph.Ancestors<TextBoxContent>().Any()) continue;
                if (visibility.HasOpenField) prefixes[paragraph] = visibility._resultRegions.Reverse().ToArray();
                visibility.ObserveParagraphMarkers(paragraph);
            }
            return prefixes;
        }

        private void ObserveParagraphMarkers(Paragraph paragraph) {
            foreach (FieldChar marker in paragraph.Descendants<FieldChar>()
                .Where(marker => ReferenceEquals(marker.Ancestors<Paragraph>().FirstOrDefault(), paragraph)))
                Observe(marker);
        }

        private sealed class ConversionScope : IDisposable {
            private readonly Dictionary<OpenXmlElement, Dictionary<Paragraph, bool[]>>? _previous;
            internal ConversionScope(Dictionary<OpenXmlElement, Dictionary<Paragraph, bool[]>>? previous) => _previous = previous;
            public void Dispose() => ConversionPrefixes.Value = _previous;
        }

        internal Run? GetVisibleRun(Run source) => GetVisibleRun(source, out _);

        internal Run? GetVisibleRun(Run source, out IReadOnlyList<OpenXmlElement>? visibleSourceChildren) {
            visibleSourceChildren = null;
            if (!source.Elements<FieldChar>().Any()) return IsVisible ? source : null;

            var visible = new Run();
            var sources = new List<OpenXmlElement>();
            bool hasContent = false;
            foreach (OpenXmlElement child in source.ChildElements) {
                if (child is FieldChar marker) {
                    Observe(marker);
                } else if (child is RunProperties) {
                    visible.Append(child.CloneNode(true));
                    sources.Add(child);
                } else if (IsVisible && child is not FieldCode) {
                    visible.Append(child.CloneNode(true));
                    sources.Add(child);
                    hasContent = true;
                }
            }

            if (hasContent) visibleSourceChildren = sources;
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

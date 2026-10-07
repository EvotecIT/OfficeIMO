using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Threading;

namespace OfficeIMO.Word {
    /// <summary>
    /// Tracks which run content is visible while walking a paragraph's complex fields.
    /// </summary>
    internal sealed class WordComplexFieldRunVisibility {
        private static readonly AsyncLocal<Dictionary<OpenXmlElement, Dictionary<Paragraph, FieldState>>?> ConversionPrefixes = new();
        private FieldState? _state;

        internal sealed class FieldState {
            internal FieldState(FieldState? parent, bool isResult, FieldChar? beginMarker) {
                Parent = parent;
                IsResult = isResult;
                BeginMarker = beginMarker;
                HiddenCount = (parent?.HiddenCount ?? 0) + (isResult ? 0 : 1);
            }

            internal FieldState? Parent { get; }
            internal bool IsResult { get; }
            internal FieldChar? BeginMarker { get; }
            internal int HiddenCount { get; }
        }

        internal WordComplexFieldRunVisibility(FieldState? state = null) => _state = state;

        internal FieldState? CurrentState => _state;
        internal bool IsVisible => (_state?.HiddenCount ?? 0) == 0;
        internal bool HasOpenField => _state != null;

        internal static IDisposable BeginConversionScope() {
            Dictionary<OpenXmlElement, Dictionary<Paragraph, FieldState>>? previous = ConversionPrefixes.Value;
            ConversionPrefixes.Value = new Dictionary<OpenXmlElement, Dictionary<Paragraph, FieldState>>();
            return new ConversionScope(previous);
        }

        internal static WordComplexFieldRunVisibility ForParagraph(Paragraph paragraph) {
            if (paragraph.Annotation<FragmentPrefix>() is FragmentPrefix fragmentPrefix)
                return new WordComplexFieldRunVisibility(fragmentPrefix.State);
            OpenXmlElement? textBoxStory = paragraph.Ancestors<TextBoxContent>().FirstOrDefault();
            OpenXmlElement story = textBoxStory ?? paragraph.Ancestors().FirstOrDefault(element =>
                element is Footnote or Endnote or Header or Footer)
                ?? paragraph.Ancestors().LastOrDefault() ?? paragraph;
            Dictionary<OpenXmlElement, Dictionary<Paragraph, FieldState>>? cache = ConversionPrefixes.Value;
            if (cache != null) {
                if (!cache.TryGetValue(story, out Dictionary<Paragraph, FieldState>? prefixes)) {
                    prefixes = BuildStoryPrefixes(story, textBoxStory != null);
                    cache.Add(story, prefixes);
                }
                return prefixes.TryGetValue(paragraph, out FieldState? prefix)
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

        /// <summary>Retains the source story's field state on detached pagination fragments.</summary>
        internal static void PreserveFragmentContext(Paragraph source, Paragraph before, Paragraph after, OpenXmlElement boundary) {
            WordComplexFieldRunVisibility visibility = ForParagraph(source);
            before.AddAnnotation(new FragmentPrefix(visibility.CurrentState));
            foreach (OpenXmlElement element in source.Descendants()) {
                if (ReferenceEquals(element, boundary)) break;
                if (element is FieldChar marker && IsObservedFieldMarker(marker, source))
                    visibility.Observe(marker);
            }
            after.AddAnnotation(new FragmentPrefix(visibility.CurrentState));
        }

        private sealed class FragmentPrefix {
            internal FragmentPrefix(FieldState? state) => State = state;
            internal FieldState? State { get; }
        }

        private static Dictionary<Paragraph, FieldState> BuildStoryPrefixes(OpenXmlElement story, bool isTextBoxStory) {
            var prefixes = new Dictionary<Paragraph, FieldState>();
            var visibility = new WordComplexFieldRunVisibility();
            foreach (Paragraph paragraph in story.Descendants<Paragraph>()) {
                if (!isTextBoxStory && paragraph.Ancestors<TextBoxContent>().Any()) continue;
                if (visibility.CurrentState is FieldState state) prefixes[paragraph] = state;
                visibility.ObserveParagraphMarkers(paragraph);
            }
            return prefixes;
        }

        internal void ObserveParagraphMarkers(Paragraph paragraph) {
            foreach (FieldChar marker in paragraph.Descendants<FieldChar>()
                .Where(marker => IsObservedFieldMarker(marker, paragraph)))
                Observe(marker);
        }

        private static bool IsObservedFieldMarker(FieldChar marker, Paragraph paragraph) =>
            ReferenceEquals(marker.Ancestors<Paragraph>().FirstOrDefault(), paragraph) &&
            !IsInIgnoredRevisionContainer(marker);

        // Visible-run readers do not traverse revision run containers. Their field
        // markers must not seed either story prefixes or detached fragment state.
        private static bool IsInIgnoredRevisionContainer(OpenXmlElement element) =>
            element.Ancestors().Any(ancestor => ancestor is DeletedRun or MoveFromRun or InsertedRun or MoveToRun);

        private sealed class ConversionScope : IDisposable {
            private readonly Dictionary<OpenXmlElement, Dictionary<Paragraph, FieldState>>? _previous;
            internal ConversionScope(Dictionary<OpenXmlElement, Dictionary<Paragraph, FieldState>>? previous) => _previous = previous;
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
            foreach (Run run in container.Descendants<Run>().Where(run => !IsInIgnoredRevisionContainer(run))) GetVisibleRun(run);
        }

        private void Observe(FieldChar marker) {
            if (marker.FieldCharType?.Value == FieldCharValues.Begin) {
                _state = new FieldState(_state, false, marker);
            } else if (marker.FieldCharType?.Value == FieldCharValues.Separate && _state != null) {
                _state = new FieldState(_state.Parent, true, _state.BeginMarker);
            } else if (marker.FieldCharType?.Value == FieldCharValues.End && _state != null) {
                _state = _state.Parent;
            }
        }
    }
}

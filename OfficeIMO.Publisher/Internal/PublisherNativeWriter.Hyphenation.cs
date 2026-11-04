namespace OfficeIMO.Publisher.Internal;

internal sealed partial class PublisherNativeWriter {
    private readonly Dictionary<uint, bool> _storyHyphenation = new();

    private void CompareStoryHyphenation() {
        foreach (var pair in _baseline.Text.Stories) {
            _token.ThrowIfCancellationRequested();
            if (!_current.Text.Stories.TryGetValue(pair.Key, out PublisherTextStory? current)
                || pair.Value.AutomaticHyphenation == current.AutomaticHyphenation) continue;
            if (_baseline.Text.IsTableStory(pair.Key) || !pair.Value.AutomaticHyphenation.HasValue || !current.AutomaticHyphenation.HasValue) {
                Reject("Contents/story/" + pair.Key, "Automatic hyphenation requires an ordinary story with a qualified source catalog flag.");
                continue;
            }
            _storyHyphenation.Add(pair.Key, current.AutomaticHyphenation.Value);
        }
    }
}

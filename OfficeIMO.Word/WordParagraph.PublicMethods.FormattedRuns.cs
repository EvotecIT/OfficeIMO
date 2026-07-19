namespace OfficeIMO.Word;

public partial class WordParagraph {
    /// <summary>
    /// Enumerates this paragraph's text and image runs with reader-facing formatting metadata.
    /// </summary>
    /// <returns>Formatted runs in document order.</returns>
    public IEnumerable<WordFormattedRun> GetFormattedRuns() {
        return FormattingHelper.GetFormattedRuns(this);
    }
}

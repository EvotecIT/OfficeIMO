using OfficeIMO.Word;

namespace PSWriteOffice.Services.Word;

/// <summary>Records native text operations until a reusable table cell is attached to its document.</summary>
internal sealed class WordCellContent {
    internal List<Action<WordParagraph>> Text { get; } = new();
    internal void Apply(WordParagraph paragraph) {
        foreach (var append in Text) append(paragraph);
    }
}

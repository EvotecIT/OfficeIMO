namespace OfficeIMO.Word;

public partial class WordList {
    /// <summary>Adds a task item containing an editable checkbox and text.</summary>
    public WordParagraph AddTaskItem(string text, bool done = false, int level = 0) {
        if (text == null) throw new ArgumentNullException(nameof(text));
        var paragraph = AddItem(string.Empty, level);
        paragraph.AddCheckBox(done);
        if (text.Length > 0) paragraph.AddText(" " + text);
        return paragraph;
    }

    /// <summary>Adds an item whose text links to the supplied destination.</summary>
    public WordParagraph AddLinkItem(string text, Uri destination, bool addStyle = false, int level = 0) {
        if (text == null) throw new ArgumentNullException(nameof(text));
        if (destination == null) throw new ArgumentNullException(nameof(destination));
        var paragraph = AddItem(string.Empty, level);
        paragraph.AddHyperLink(text, destination, addStyle);
        return paragraph;
    }
}

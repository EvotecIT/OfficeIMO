using System.Xml;

namespace OfficeIMO.Word;

/// <summary>A plain-text root comment anchored to a cell in the document body.</summary>
public sealed class WordCellComment {
    /// <summary>Creates a cell comment. A null date leaves the creation timestamp unspecified.</summary>
    public WordCellComment(WordTableCell cell, string author, string initials, string text, DateTime? dateTime = null) {
        Cell = cell ?? throw new ArgumentNullException(nameof(cell));
        Author = author ?? throw new ArgumentNullException(nameof(author));
        Initials = initials ?? throw new ArgumentNullException(nameof(initials));
        Text = text ?? throw new ArgumentNullException(nameof(text));
        DateTime = dateTime;
    }

    /// <summary>Gets the destination cell. Its direct paragraphs form the comment range.</summary>
    public WordTableCell Cell { get; }
    /// <summary>Gets the author's exact display name.</summary>
    public string Author { get; }
    /// <summary>Gets the author's initials, which may be empty.</summary>
    public string Initials { get; }
    /// <summary>Gets plain text, including line feeds, tabs and whitespace.</summary>
    public string Text { get; }
    /// <summary>Gets the creation timestamp, or null when unspecified.</summary>
    public DateTime? DateTime { get; }

    /// <summary>Determines whether Word's plain-text paragraph owner preserves the text without normalization.</summary>
    /// <remarks>Carriage returns and U+2028 have structural meanings in the paragraph API. Invalid XML characters are rejected.</remarks>
    public static bool CanPreserveText(string text) {
        if (text == null) throw new ArgumentNullException(nameof(text));
        if (text.IndexOf('\r') >= 0 || text.IndexOf('\u2028') >= 0) return false;
        try { XmlConvert.VerifyXmlChars(text); return true; }
        catch (XmlException) { return false; }
    }
}

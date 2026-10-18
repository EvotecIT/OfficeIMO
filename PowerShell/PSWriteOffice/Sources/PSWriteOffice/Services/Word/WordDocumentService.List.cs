using System;
using System.Collections;
using OfficeIMO.Word;

namespace PSWriteOffice.Services.Word;

public static partial class WordDocumentService
{
    /// <summary>Creates a new list using the supplied style.</summary>
    public static WordList AddList(WordDocument document, WordListStyle style)
    {
        return document.AddList(style);
    }

    /// <summary>Adds a list item with optional level.</summary>
    public static object AddListItem(WordList list, int level, string[] text)
    {
        // Use the correct overload of AddItem which requires a WordParagraph
        var combinedText = string.Join(" ", text);
        // The AddItem method signature is: AddItem(String text, Int32 level, WordParagraph wordParagraph)
        // We need to pass null for wordParagraph to add at the end
        return list.AddItem(combinedText, level, null);
    }
}

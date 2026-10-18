using OfficeIMO.Word;

namespace PSWriteOffice.Services.Word;

public static partial class WordDocumentService
{
    /// <summary>Removes all footers from the document.</summary>
    public static void RemoveFooters(WordDocument document)
    {
        WordFooter.RemoveFooters(document);
    }

    /// <summary>Removes all headers from the document.</summary>
    public static void RemoveHeaders(WordDocument document)
    {
        WordHeader.RemoveHeaders(document);
    }
}

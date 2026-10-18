using System;
using OfficeIMO.Word;

namespace PSWriteOffice.Services.Word;

internal static class WordHostExtensions
{
    public static WordParagraph AddParagraph(this object host, string? text = null)
    {
        switch (host)
        {
            case WordSection section:
                return string.IsNullOrEmpty(text) ? section.AddParagraph() : section.AddParagraph(text ?? string.Empty);
            case WordHeader header:
                return string.IsNullOrEmpty(text) ? header.AddParagraph() : header.AddParagraph(text ?? string.Empty);
            case WordFooter footer:
                return string.IsNullOrEmpty(text) ? footer.AddParagraph() : footer.AddParagraph(text ?? string.Empty);
            case WordTableCell cell:
                return string.IsNullOrEmpty(text) ? cell.AddParagraph() : cell.AddParagraph(text ?? string.Empty);
            default:
                throw new InvalidOperationException("Paragraphs can only be added inside sections, headers, footers, or table cells.");
        }
    }

    public static WordList AddList(this WordDocument document, WordListStyle style)
    {
        return document.AddList(style);
    }
}

using System;
using OfficeIMO.Word;

namespace PSWriteOffice.Services.Word;

public static partial class WordDocumentService
{
    /// <summary>Adds text runs to the specified paragraph.</summary>
    public static WordParagraph AddText(WordDocument? document, WordParagraph? paragraph, string[] text, bool?[]? bold,
        bool?[]? italic, WordUnderlineStyle?[]? underline, string[]? color, WordParagraphAlignment? alignment,
        WordParagraphStyles? style)
    {
        if (bold != null && bold.Length != text.Length)
        {
            throw new ArgumentException("bold length must match text length", nameof(bold));
        }
        if (italic != null && italic.Length != text.Length)
        {
            throw new ArgumentException("italic length must match text length", nameof(italic));
        }
        if (underline != null && underline.Length != text.Length)
        {
            throw new ArgumentException("underline length must match text length", nameof(underline));
        }
        if (color != null && color.Length != text.Length)
        {
            throw new ArgumentException("color length must match text length", nameof(color));
        }

        WordParagraph para;
        if (paragraph != null)
        {
            para = paragraph;
        }
        else
        {
            para = document?.AddParagraph() ?? throw new ArgumentNullException(nameof(document));
        }

        var boldArray = bold;
        var italicArray = italic;
        var underlineArray = underline;

        if (alignment.HasValue)
        {
            para.ParagraphAlignment = alignment.Value;
        }

        for (var t = 0; t < text.Length; t++)
        {
            para = para.AddText(text[t]);

            if (boldArray != null && t < boldArray.Length && boldArray[t].HasValue)
            {
                para.Bold = boldArray[t]!.Value;
            }
            if (italicArray != null && t < italicArray.Length && italicArray[t].HasValue)
            {
                para.Italic = italicArray[t]!.Value;
            }
            if (underlineArray != null && t < underlineArray.Length && underlineArray[t].HasValue)
            {
                para.Underline = underlineArray[t]!.Value;
            }
        }

        return para;
    }
}

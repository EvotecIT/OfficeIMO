using System;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Html;
using OfficeIMO.Word;
using OfficeIMO.Word.Html;

namespace PSWriteOffice.Services.Word;

public static partial class WordDocumentService
{
    /// <summary>Adds HTML content to the document.</summary>
    public static void AddHtml(WordDocument document, string html, HtmlImportMode mode = HtmlImportMode.Parse)
    {
        if (mode == HtmlImportMode.AsIs)
        {
            // For AsIs mode, embed the HTML directly without parsing
            document.AddEmbeddedFragment(html, WordAlternativeFormatImportPartType.Html);
            return;
        }

        document.AddHtmlToBody(HtmlConversionDocument.Parse(html));
    }
}

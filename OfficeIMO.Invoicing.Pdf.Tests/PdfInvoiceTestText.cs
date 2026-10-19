using System.Text.RegularExpressions;
using OfficeIMO.Pdf;

namespace OfficeIMO.Invoicing.Pdf.Tests;

internal static class PdfInvoiceTestText {
    private static readonly Regex HyphenLineWrap = new(@"-\r?\n", RegexOptions.CultureInvariant);

    /// <summary>Extracts page text with line wraps after hyphens rejoined, so identifiers read whole.</summary>
    internal static string ExtractUnwrappedText(this PdfReadDocument document) =>
        HyphenLineWrap.Replace(document.ExtractText(), "-");

    internal static string ExtractUnwrappedText(this PdfReadPage page) =>
        HyphenLineWrap.Replace(page.ExtractText(), "-");
}

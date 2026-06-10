namespace OfficeIMO.Pdf;

/// <summary>
/// Supplies preferred UTF-16 break positions for a single unspaced text token during generated PDF text wrapping.
/// </summary>
/// <param name="token">The unspaced token that did not fit in the current text box.</param>
/// <returns>Zero-based UTF-16 indexes after which the token may be split. Invalid or surrogate-splitting indexes are ignored.</returns>
public delegate System.Collections.Generic.IReadOnlyList<int> PdfTextHyphenationCallback(string token);

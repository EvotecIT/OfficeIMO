namespace OfficeIMO.Pdf;

internal static partial class ResourceResolver {
    /// <summary>Describes the decoder branch after a complete font text run decoded successfully.</summary>
    internal static PdfTextDecodingEvidence GetDecodingEvidence(PdfFontResource font, byte[] bytes,
        Action? cancellationCheck, bool preferToUnicode = true, bool useWinAnsiFallback = true) {
        cancellationCheck?.Invoke();
        PdfTextDecodingSource source;
        int unmapped = 0;
        if (preferToUnicode && font.HasToUnicode && font.CMap is not null) {
            source = PdfTextDecodingSource.ToUnicode;
            unmapped = font.CMap.CountUnmappedCodeBytes(bytes, cancellationCheck);
        } else if (UsesNamedCompositeEncoding(font) && (!preferToUnicode || !font.HasToUnicode) && font.PredefinedCMap is not null) {
            source = PdfTextDecodingSource.PredefinedCMap;
        } else if (UsesNamedCompositeEncoding(font)) {
            source = PdfTextDecodingSource.Unmappable;
        } else if (font.Encoding is "StandardEncoding" or "MacRomanEncoding" ||
            font.HasDeclaredEncoding && font.Encoding == "WinAnsiEncoding") {
            source = PdfTextDecodingSource.FontEncoding;
        } else {
            source = PdfTextDecodingSource.WinAnsiFallback;
        }
        if (source == PdfTextDecodingSource.WinAnsiFallback && font.Differences is { Count: > 0 } differences) {
            bool allSuppliedByDifferences = true;
            foreach (byte code in bytes) {
                cancellationCheck?.Invoke();
                if (!differences.ContainsKey(code)) { allSuppliedByDifferences = false; break; }
            }
            if (allSuppliedByDifferences) source = PdfTextDecodingSource.FontEncoding;
        }
        if (!useWinAnsiFallback && (source == PdfTextDecodingSource.WinAnsiFallback || unmapped > 0)) {
            source = PdfTextDecodingSource.Unmappable;
        }
        return new PdfTextDecodingEvidence(source, font.Encoding, font.HasToUnicode,
            (source is PdfTextDecodingSource.FontEncoding or PdfTextDecodingSource.WinAnsiFallback) && font.Differences is { Count: > 0 }, unmapped);
    }
}

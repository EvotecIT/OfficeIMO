using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

public sealed partial class PdfEmbeddedFontFamily {
    internal const int MaxTrueTypeCollectionFontsToInspect = OfficeTrueTypeCollection.MaxTrueTypeCollectionFontsToInspect;
    internal const int MaxExtractedTrueTypeCollectionFontBytes = OfficeTrueTypeCollection.MaxExtractedTrueTypeCollectionFontBytes;
    internal const int MaxExtractedTrueTypeCollectionBytes = OfficeTrueTypeCollection.MaxExtractedTrueTypeCollectionBytes;

    private static System.Collections.Generic.List<byte[]> ExtractTrueTypeFontPrograms(byte[] data) =>
        OfficeTrueTypeCollection.ExtractPrograms(data);

    private static bool IsTrueTypeCollection(byte[] data) => OfficeTrueTypeCollection.IsTrueTypeCollection(data);
}

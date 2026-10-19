namespace OfficeIMO.Pdf;

internal static partial class PdfIncrementalUpdater {
    private static PdfReference CreateSignatureAppearanceFont(PdfEmbeddedFont font, string text,
        Dictionary<int, PdfIndirectObject> objects, ref int nextObjectNumber,
        HashSet<int> changedObjects, out string encodedText) {
        PdfTrueTypeFontProgram program = PdfFontProgramCache.GetTrueType(font.DataSnapshot, font.FontName);
        encodedText = program.EncodeTextAsGlyphHex(text);
        int first = nextObjectNumber;
        int fontFile = nextObjectNumber++;
        int descriptor = nextObjectNumber++;
        int descendant = nextObjectNumber++;
        int toUnicode = nextObjectNumber++;
        int type0 = nextObjectNumber++;
        PdfFormFiller.MaterializeEmbeddedTextAppearanceFont(objects, program, fontFile, descriptor,
            descendant, toUnicode, type0);
        for (int number = first; number < nextObjectNumber; number++) changedObjects.Add(number);
        return new PdfReference(type0, 0);
    }
}

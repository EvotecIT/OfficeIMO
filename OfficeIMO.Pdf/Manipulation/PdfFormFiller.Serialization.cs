namespace OfficeIMO.Pdf;

internal static partial class PdfFormFiller {
    private static byte[] RewriteAllObjects(Dictionary<int, PdfIndirectObject> objects, int catalogObjectNumber, PdfMetadata metadata, byte[] sourcePdf) {
        var sourceIds = objects.Keys.OrderBy(id => id).ToArray();
        var numberMap = new Dictionary<int, int>(sourceIds.Length);
        for (int i = 0; i < sourceIds.Length; i++) {
            numberMap[sourceIds[i]] = i + 1;
        }

        var context = new PdfPageExtractor.SerializationContext(numberMap, pagesObjectId: 0, new Dictionary<int, Dictionary<string, PdfObject>>(), objects);
        var rewritten = new List<byte[]>(sourceIds.Length + 1);
        foreach (int sourceId in sourceIds) {
            byte[] body = PdfPageExtractor.SerializeObject(objects[sourceId].Value, context);
            rewritten.Add(PdfPageExtractor.WrapObject(numberMap[sourceId], body));
        }

        int infoId = rewritten.Count + 1;
        rewritten.Add(PdfPageExtractor.WrapObject(infoId, PdfEncoding.Latin1GetBytes(PdfPageExtractor.BuildInfoDictionary(metadata))));

        PdfFileVersion fileVersion = PdfFileAssembler.ParseHeaderVersionOrDefault(PdfSyntax.GetHeaderVersion(sourcePdf));
        if (ContainsOpenTypeFontFileStream(objects)) {
            fileVersion = PdfFileAssembler.RequireAtLeast(fileVersion, PdfFileVersion.Pdf16);
        }

        return PdfPageExtractor.Assemble(rewritten, numberMap[catalogObjectNumber], infoId, fileVersion);
    }

    private static bool ContainsOpenTypeFontFileStream(Dictionary<int, PdfIndirectObject> objects) {
        foreach (PdfIndirectObject indirect in objects.Values) {
            if (indirect.Value is PdfStream stream &&
                stream.Dictionary.Get<PdfName>("Subtype")?.Name == "OpenType") {
                return true;
            }
        }

        return false;
    }
}

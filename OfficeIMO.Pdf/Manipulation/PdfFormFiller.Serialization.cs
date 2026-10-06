using System.Threading;

namespace OfficeIMO.Pdf;

internal static partial class PdfFormFiller {
    private static byte[] RewriteAllObjects(Dictionary<int, PdfIndirectObject> objects, int catalogObjectNumber, PdfReadDocument source, byte[] sourcePdf,
        CancellationToken cancellationToken = default) {
        PdfSourceEncryptionContext? sourceEncryption = PdfSourceEncryptionContext.Create(source, cancellationToken);
        var sourceIds = objects.Keys.OrderBy(id => id).ToArray();
        var numberMap = new Dictionary<int, int>(sourceIds.Length);
        for (int i = 0; i < sourceIds.Length; i++) {
            cancellationToken.ThrowIfCancellationRequested();
            numberMap[sourceIds[i]] = i + 1;
        }

        var context = new PdfPageExtractor.SerializationContext(
            numberMap,
            pagesObjectId: 0,
            new Dictionary<int, Dictionary<string, PdfObject>>(),
            objects,
            preserveRawStringBytes: true);
        var rewritten = new List<byte[]>(sourceIds.Length + 1);
        foreach (int sourceId in sourceIds) {
            cancellationToken.ThrowIfCancellationRequested();
            rewritten.Add(PdfPageExtractor.SerializeIndirectObject(numberMap[sourceId], objects[sourceId].Value, context));
        }

        int infoId = rewritten.Count + 1;
        rewritten.Add(PdfPageExtractor.WrapObject(infoId, PdfEncoding.Latin1GetBytes(PdfPageExtractor.BuildInfoDictionary(source.UncheckedMetadata))));

        PdfFileVersion fileVersion = PdfFileAssembler.ParseHeaderVersionOrDefault(PdfSyntax.GetHeaderVersion(sourcePdf));
        if (ContainsOpenTypeFontFileStream(objects)) {
            fileVersion = PdfFileAssembler.RequireAtLeast(fileVersion, PdfFileVersion.Pdf16);
        }

        cancellationToken.ThrowIfCancellationRequested();
        byte[] result = PdfPageExtractor.Assemble(rewritten, numberMap[catalogObjectNumber], infoId, fileVersion, cancellationToken);
        return sourceEncryption?.Protect(result,
            generatedReadOptions: PdfLoadOptions.ForGeneratedOutput(source.ReadOptions, sourcePdf, result),
            cancellationToken: cancellationToken) ?? result;
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

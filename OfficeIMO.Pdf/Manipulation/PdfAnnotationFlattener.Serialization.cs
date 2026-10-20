using System.Threading;

namespace OfficeIMO.Pdf;

internal static partial class PdfAnnotationFlattener {
    private static byte[] RewriteAllObjects(Dictionary<int, PdfIndirectObject> objects, int catalogObjectNumber, PdfMetadata metadata,
        out IReadOnlyDictionary<int, int> objectNumberMap, PdfSourceEncryptionContext? sourceEncryption,
        PdfFileVersion fileVersion, byte[] sourcePdf, PdfLoadOptions? sourceReadOptions,
        PdfGeneratedOutputGrowth generatedGrowth, CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        var sourceIds = objects.Keys.OrderBy(id => id).ToArray();
        var numberMap = new Dictionary<int, int>(sourceIds.Length);
        for (int i = 0; i < sourceIds.Length; i++) {
            cancellationToken.ThrowIfCancellationRequested();
            numberMap[sourceIds[i]] = i + 1;
        }

        var context = new PdfPageExtractor.SerializationContext(numberMap, pagesObjectId: 0, new Dictionary<int, Dictionary<string, PdfObject>>(), objects, cancellationToken: cancellationToken);
        var rewritten = new List<byte[]>(sourceIds.Length + 1);
        foreach (int sourceId in sourceIds) {
            cancellationToken.ThrowIfCancellationRequested();
            rewritten.Add(PdfPageExtractor.SerializeIndirectObject(numberMap[sourceId], objects[sourceId].Value, context));
        }

        int infoId = rewritten.Count + 1;
        rewritten.Add(PdfPageExtractor.WrapObject(infoId, PdfPageExtractor.BuildInfoDictionaryBytesCancellable(metadata, cancellationToken)));

        objectNumberMap = numberMap;
        byte[] result = PdfPageExtractor.Assemble(rewritten, numberMap[catalogObjectNumber], infoId, fileVersion, cancellationToken);
        return sourceEncryption?.Protect(result,
            generatedReadOptions: PdfLoadOptions.ForGeneratedOutput(sourceReadOptions, sourcePdf, result, generatedGrowth), cancellationToken: cancellationToken) ?? result;
    }
}

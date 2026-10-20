using System.Threading;

namespace OfficeIMO.Pdf;

internal static partial class PdfAnnotationEditor {
    private static byte[] RewriteAllObjects(Dictionary<int, PdfIndirectObject> objects, int catalogObjectNumber, PdfMetadata metadata, byte[] sourcePdf) =>
        RewriteAllObjects(objects, catalogObjectNumber, metadata, sourcePdf, out _);

    private static byte[] RewriteAllObjects(
        Dictionary<int, PdfIndirectObject> objects,
        int catalogObjectNumber,
        PdfMetadata metadata,
        byte[] sourcePdf,
        out IReadOnlyDictionary<int, int> rewrittenObjectNumbers, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        int[] sourceIds = objects.Keys.OrderBy(static id => id).ToArray();
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

        PdfFileVersion fileVersion = PdfFileAssembler.ParseHeaderVersionOrDefault(PdfSyntax.GetHeaderVersion(sourcePdf));
        rewrittenObjectNumbers = numberMap;
        return PdfPageExtractor.Assemble(rewritten, numberMap[catalogObjectNumber], infoId, fileVersion, cancellationToken);
    }

}

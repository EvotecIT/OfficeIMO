using System.Threading;

namespace OfficeIMO.Pdf;

internal static partial class PdfFormFiller {
    /// <summary>Reserves work represented by appearance and font objects created by this fill operation.</summary>
    internal static PdfGeneratedOutputGrowth GetFilledAppearanceGrowth(Dictionary<int, PdfIndirectObject> objects,
        int firstGeneratedObject, CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        int[] generated = objects.Keys.Where(number => number >= firstGeneratedObject).ToArray();
        PdfGeneratedOutputGrowth serialized = PdfGeneratedOutputGrowth.FromSerializedObjects(objects, generated);
        cancellationToken.ThrowIfCancellationRequested();
        int contentBytes = (int)Math.Min(int.MaxValue, serialized.AdditionalTotalDecodedStreamBytes);
        // The bound comes from the actual generated payload, including any newly embedded fonts.
        // Existing source objects are admitted with the original caller policy before this allowance is used.
        return new PdfGeneratedOutputGrowth(
            minimumRawStreamBytes: serialized.MinimumRawStreamBytes,
            minimumDecodedStreamBytes: serialized.MinimumDecodedStreamBytes,
            additionalTotalDecodedStreamBytes: serialized.AdditionalTotalDecodedStreamBytes,
            additionalPageContentBytes: contentBytes,
            additionalRetainedContentBytes: serialized.AdditionalTotalDecodedStreamBytes,
            additionalDecodedTextCharacters: contentBytes,
            minimumObjectCharacters: serialized.MinimumObjectCharacters,
            minimumTokensPerObject: serialized.MinimumTokensPerObject,
            minimumObjectNestingDepth: serialized.MinimumObjectNestingDepth,
            additionalContentOperations: contentBytes,
            additionalContentOperands: contentBytes,
            additionalContentNestingDepth: generated.Length == 0 ? 0 : 1,
            additionalPrintProductionContexts: generated.Length,
            additionalPrintProductionOperations: contentBytes > int.MaxValue / 2 ? int.MaxValue : contentBytes * 2);
    }
}

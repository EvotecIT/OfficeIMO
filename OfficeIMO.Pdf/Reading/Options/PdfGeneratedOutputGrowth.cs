namespace OfficeIMO.Pdf;

internal readonly struct PdfGeneratedOutputGrowth {
    internal PdfGeneratedOutputGrowth(
        int additionalRevisions = 0,
        int additionalAnnotationsPerPage = 0,
        int minimumRawStreamBytes = 0,
        int minimumDecodedStreamBytes = 0,
        long additionalTotalDecodedStreamBytes = 0L,
        int additionalPageContentBytes = 0,
        long additionalRetainedContentBytes = 0L,
        int additionalDecodedTextCharacters = 0,
        int minimumObjectCharacters = 0,
        int minimumTokensPerObject = 0,
        int minimumObjectNestingDepth = 0,
        int additionalContentOperations = 0,
        int additionalContentOperands = 0,
        int additionalContentNestingDepth = 0) {
        AdditionalRevisions = RequireNonNegative(additionalRevisions, nameof(additionalRevisions));
        AdditionalAnnotationsPerPage = RequireNonNegative(additionalAnnotationsPerPage, nameof(additionalAnnotationsPerPage));
        MinimumRawStreamBytes = RequireNonNegative(minimumRawStreamBytes, nameof(minimumRawStreamBytes));
        MinimumDecodedStreamBytes = RequireNonNegative(minimumDecodedStreamBytes, nameof(minimumDecodedStreamBytes));
        AdditionalTotalDecodedStreamBytes = RequireNonNegative(additionalTotalDecodedStreamBytes, nameof(additionalTotalDecodedStreamBytes));
        AdditionalPageContentBytes = RequireNonNegative(additionalPageContentBytes, nameof(additionalPageContentBytes));
        AdditionalRetainedContentBytes = RequireNonNegative(additionalRetainedContentBytes, nameof(additionalRetainedContentBytes));
        AdditionalDecodedTextCharacters = RequireNonNegative(additionalDecodedTextCharacters, nameof(additionalDecodedTextCharacters));
        MinimumObjectCharacters = RequireNonNegative(minimumObjectCharacters, nameof(minimumObjectCharacters));
        MinimumTokensPerObject = RequireNonNegative(minimumTokensPerObject, nameof(minimumTokensPerObject));
        MinimumObjectNestingDepth = RequireNonNegative(minimumObjectNestingDepth, nameof(minimumObjectNestingDepth));
        AdditionalContentOperations = RequireNonNegative(additionalContentOperations, nameof(additionalContentOperations));
        AdditionalContentOperands = RequireNonNegative(additionalContentOperands, nameof(additionalContentOperands));
        AdditionalContentNestingDepth = RequireNonNegative(additionalContentNestingDepth, nameof(additionalContentNestingDepth));
    }

    internal int AdditionalRevisions { get; }
    internal int AdditionalAnnotationsPerPage { get; }
    internal int MinimumRawStreamBytes { get; }
    internal int MinimumDecodedStreamBytes { get; }
    internal long AdditionalTotalDecodedStreamBytes { get; }
    internal int AdditionalPageContentBytes { get; }
    internal long AdditionalRetainedContentBytes { get; }
    internal int AdditionalDecodedTextCharacters { get; }
    internal int MinimumObjectCharacters { get; }
    internal int MinimumTokensPerObject { get; }
    internal int MinimumObjectNestingDepth { get; }
    internal int AdditionalContentOperations { get; }
    internal int AdditionalContentOperands { get; }
    internal int AdditionalContentNestingDepth { get; }

    internal static PdfGeneratedOutputGrowth FromSerializedObjects(
        Dictionary<int, PdfIndirectObject> objects,
        IEnumerable<int> objectNumbers,
        int additionalAnnotationsPerPage = 0,
        int additionalRevisions = 0) {
        Guard.NotNull(objects, nameof(objects));
        Guard.NotNull(objectNumbers, nameof(objectNumbers));
        int maximumStreamBytes = 0;
        int maximumSerializedObjectBytes = 0;
        long totalStreamBytes = 0L;
        int[] selected = objectNumbers.Distinct().ToArray();
        var identityMap = objects.Keys.ToDictionary(static objectNumber => objectNumber, static objectNumber => objectNumber);
        var serializationContext = new PdfPageExtractor.SerializationContext(
            identityMap,
            pagesObjectId: 0,
            new Dictionary<int, Dictionary<string, PdfObject>>(),
            objects);
        foreach (int objectNumber in selected) {
            if (!objects.TryGetValue(objectNumber, out PdfIndirectObject? indirect)) continue;
            maximumSerializedObjectBytes = Math.Max(
                maximumSerializedObjectBytes,
                PdfPageExtractor.SerializeObject(indirect.Value, serializationContext).Length);
            if (indirect.Value is not PdfStream stream) continue;
            maximumStreamBytes = Math.Max(maximumStreamBytes, stream.Data.Length);
            totalStreamBytes = totalStreamBytes > long.MaxValue - stream.Data.LongLength
                ? long.MaxValue
                : totalStreamBytes + stream.Data.LongLength;
        }

        return new PdfGeneratedOutputGrowth(
            additionalRevisions: additionalRevisions,
            additionalAnnotationsPerPage: additionalAnnotationsPerPage,
            minimumRawStreamBytes: maximumStreamBytes,
            minimumDecodedStreamBytes: maximumStreamBytes,
            additionalTotalDecodedStreamBytes: totalStreamBytes,
            minimumObjectCharacters: maximumSerializedObjectBytes,
            minimumTokensPerObject: maximumSerializedObjectBytes,
            minimumObjectNestingDepth: selected.Length == 0 ? 0 : 16);
    }

    private static int RequireNonNegative(int value, string parameterName) =>
        value >= 0 ? value : throw new ArgumentOutOfRangeException(parameterName, value, "Generated-output growth values cannot be negative.");

    private static long RequireNonNegative(long value, string parameterName) =>
        value >= 0L ? value : throw new ArgumentOutOfRangeException(parameterName, value, "Generated-output growth values cannot be negative.");
}

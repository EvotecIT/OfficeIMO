namespace OfficeIMO.Publisher.Internal;

/// <summary>Resolves the qualified modern text-engine override and publication fallback.</summary>
internal static class PublisherDefaultTabInterval {
    internal static double? ReadDocument(IReadOnlyList<PublisherBlock> fields, PublisherParseContext context) =>
        ReadValue(fields, 0x15, 0x20, context, "Contents/document");

    internal static double ReadText(PublisherBinaryData data, IReadOnlyList<PublisherQuillChunk> chunks,
        double? documentInterval, PublisherParseContext context) {
        var sections = chunks.Where(chunk => chunk.Tag == "SGP ").ToArray();
        if (sections.Length == 1) {
            IReadOnlyList<PublisherBlock> fields;
            try { fields = new PublisherBlockReader(data, context).Chunk(sections[0].Offset, sections[0].End); }
            catch (NotSupportedException) {
                Unassessed(context, "Quill/SGP");
                return documentInterval ?? 36D;
            }
            double? textInterval = ReadValue(fields, 0, 0x22, context, "Quill/SGP");
            if (textInterval.HasValue) return textInterval.Value;
        } else if (sections.Length > 1) Unassessed(context, "Quill/SGP");
        if (documentInterval.HasValue) return documentInterval.Value;
        Unassessed(context, "Contents/document");
        return 36D;
    }

    private static double? ReadValue(IReadOnlyList<PublisherBlock> fields, byte id, byte type,
        PublisherParseContext context, string location) {
        var values = fields.Where(field => field.Id == id).ToArray();
        if (values.Length == 0) return null;
        if (values.Length == 1 && values[0].Type == type && values[0].Value is > 0 and <= int.MaxValue)
            return values[0].Value / 12700D;
        Unassessed(context, location);
        return null;
    }

    private static void Unassessed(PublisherParseContext context, string location) => context.Add(
        "PUB_DEFAULT_TAB_UNASSESSED", "A native default-tab carrier is missing or outside the qualified positive scalar profile. Available publication metadata supplies the fallback; otherwise layout uses the generic 36-point grid.",
        OfficeConversionLossKind.Unassessed, location);
}

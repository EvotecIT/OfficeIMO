namespace OfficeIMO.Publisher.Internal;

internal static partial class PublisherModelImporter {
    /// <summary>Reserves the imported publication identity space once, including dangling live references. All authored object kinds use the same current allocator.</summary>
    private static uint NextModelObjectId(PublisherContents source, PublisherQuillText text, IReadOnlyDictionary<uint, uint?> owners,
        IReadOnlyList<PublisherPageState> pages, PublisherTextState currentText) {
        uint maximum = source.Chunks.Keys.Concat(source.Shapes.Keys).Concat(pages.Select(page => page.Id)).Concat(owners.Keys)
            .Concat(owners.Values.Where(id => id.HasValue).Select(id => id!.Value)).Concat(text.Stories.Keys)
            .Concat(currentText.Frames.Values.SelectMany(frame => frame.WrapObjectIds)).DefaultIfEmpty(0U).Max();
        foreach (PublisherSourceShape shape in source.Shapes.Values)
            foreach (byte field in new byte[] { 0x27, 0x36, 0x37 }) maximum = Math.Max(maximum, shape.Value(field) ?? 0);
        foreach (PublisherSourcePage page in source.Pages) maximum = Math.Max(maximum, page.Master ?? 0);
        return maximum == uint.MaxValue ? 0 : maximum + 1;
    }
}

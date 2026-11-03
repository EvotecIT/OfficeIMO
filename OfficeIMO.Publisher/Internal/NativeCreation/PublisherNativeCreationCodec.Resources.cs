using OfficeIMO.Core.Internal;

namespace OfficeIMO.Publisher.Internal;

internal static partial class PublisherNativeCreationCodec {
    /// <summary>Admits the plan's semantic resources before constructing native streams; source-inspection limits do not count generated records.</summary>
    private static void ValidateResources(PublisherNativeCreationPlan plan, PublisherParseContext context, long maxOutputBytes) {
        context.Token.ThrowIfCancellationRequested();
        PublisherReadOptions options = context.Options;
        int limit = options.Limits.MaxItems;
        if ((long)plan.PhysicalPageCount + plan.MasterDefinitions.Count > options.MaximumPages)
            throw new InvalidDataException("Publisher page limit exceeded.");
        if (plan.Objects.Count > limit) throw new InvalidDataException("Publisher current object limit exceeded.");
        if ((long)plan.Objects.Count + 1 > PublisherNativeCreationArtwork.MaximumOrdinaryShapes)
            throw new NotSupportedException("The ordinary object drawing specimen is limited to one SPID cluster; table utility drawings can span clusters.");

        // Names and stories use their existing separate model admission families.
        long names = 0;
        if (plan.Masters != null) foreach (PublisherNativeCreationMaster master in plan.Masters) {
            context.Token.ThrowIfCancellationRequested();
            if (master.Name.Length > options.Limits.MaxTextCharacters - names)
                throw new InvalidDataException("Publisher master names exceed the configured character limit.");
            names += master.Name.Length;
        }

        var stories = new HashSet<PublisherNativeCreationStory>();
        var tables = new HashSet<PublisherTable>();
        var assets = new Dictionary<int, byte[]>();
        long characters = 0, nativeCharacters = 0, tabs = 0;
        foreach (PublisherNativeCreationObject item in plan.Objects) {
            context.Token.ThrowIfCancellationRequested();
            if (item.Story is { } story && stories.Add(story)) {
                if (stories.Count > limit) throw new InvalidDataException("Publisher story limit exceeded.");
                if (story.CharacterCount > options.Limits.MaxTextCharacters - characters)
                    throw new InvalidDataException("Publisher publication text exceeds the configured character limit.");
                characters += story.CharacterCount;
                if (story.Paragraphs == null) nativeCharacters += story.Text.Length;
                else foreach (var paragraph in story.Paragraphs) {
                    context.Token.ThrowIfCancellationRequested();
                    nativeCharacters++;
                    foreach (var run in paragraph.Runs) {
                        context.Token.ThrowIfCancellationRequested();
                        nativeCharacters += run.Text.Length;
                    }
                    int stops = paragraph.TabStops?.Stops.Count ?? 0;
                    if (stops > limit - tabs) throw new InvalidDataException("Publisher authored tab item limit exceeded.");
                    tabs += stops;
                }
            }
            if (item.Table is { } table && tables.Add(table)) {
                context.AccountTableModel(checked(1 + table.ColumnWidths.Count + table.RowHeights.Count + table.Cells.Count));
            }
            if (item.Image is { } image) {
                if (image.Length > options.MaximumImageBytes)
                    throw new InvalidDataException("Publisher encoded image byte limit exceeded.");
                if (item.ImageId is { } imageId && assets.TryGetValue(imageId, out byte[]? existing)) {
                    if (!PublisherNativeCreationArtwork.ImagePayloadsEqual(existing, image, context.Token))
                        throw new InvalidDataException("One model image identifier has conflicting payloads.");
                    continue;
                }
                context.AccountImageStoreEntry();
                context.AccountImage(image.Length);
                if (item.ImageId is { } id) assets.Add(id, image);
            }
        }
        // These payloads must be stored verbatim even before native and CFB overhead.
        if (context.ImageBytes > maxOutputBytes || nativeCharacters > (maxOutputBytes - context.ImageBytes) / 2)
            throw OfficeOutputLimit.Create("Publisher native payloads exceed the configured output byte limit.");
        context.Token.ThrowIfCancellationRequested();
    }
}

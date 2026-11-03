using System.Text;
using OfficeIMO.Publisher.Internal;

namespace OfficeIMO.Publisher.Internal;

/// <summary>
/// Synthesizes the document, explicit page roles and small Quill document graph.
/// Native acceptance is limited to the independently qualified specimens.
/// </summary>
internal static class PublisherNativeCreationGraph {
    internal const ushort SpecialPageBoundary = 5;
    internal const uint DrawingId = 1;
    internal const uint RootShapeId = DrawingId << 10;

    // Native enumeration admits M..N-S inclusive: physical = N-S-M+1.
    internal static ushort GetMasterPageCount(PublisherNativeCreationPlan specification) => checked((ushort)specification.MasterDefinitions.Count);
    internal static ushort GetOrderedPageCount(PublisherNativeCreationPlan specification) =>
        checked((ushort)(specification.PhysicalPageCount + GetMasterPageCount(specification) + SpecialPageBoundary - 1));

    internal static Graph Create(PublisherNativeCreationPlan specification, uint documentId, PublisherParseContext context) {
        context.Token.ThrowIfCancellationRequested();
        ushort orderedPageCount = GetOrderedPageCount(specification);
        ushort masterCount = GetMasterPageCount(specification);
        if (masterCount == 0 || specification.PageMasters != null && specification.PageMasters.Count != specification.PhysicalPageCount)
            throw new ArgumentException("The specimen requires master definitions and one assignment per physical page.");
        uint width = specification.WidthEmu, height = specification.HeightEmu;
        uint firstPageId = documentId + 1;
        uint firstObjectId = firstPageId + orderedPageCount;
        var objects = new List<ObjectBinding>();
        var stories = new List<StoryBinding>();
        var storyBindings = new Dictionary<PublisherNativeCreationStory, StoryBinding>();
        var namedPictures = new List<ObjectBinding>();
        var tables = new List<ObjectBinding>();
        var textFrames = new List<ObjectBinding>();
        for (int index = 0; index < specification.Objects.Count; index++) {
            context.Token.ThrowIfCancellationRequested();
            PublisherNativeCreationObject source = specification.Objects[index];
            StoryBinding? story = null;
            if (source.Story != null) {
                if (!storyBindings.TryGetValue(source.Story, out story)) {
                    uint key = (uint)stories.Count + 1;
                    story = new StoryBinding(source.Story, key, key);
                    stories.Add(story);
                    storyBindings.Add(source.Story, story);
                }
            }
            if (source.PageIndex < 0 || source.PageIndex >= (source.OnMaster ? masterCount : specification.PhysicalPageCount))
                throw new ArgumentException("An object owner is outside the authored page definitions.");
            uint ownerPageId = firstPageId + checked((uint)(source.OnMaster ? source.PageIndex : masterCount + source.PageIndex));
            double originX = source.OnMaster ? specification.MasterDefinitions[source.PageIndex].Side switch {
                PublisherMasterPageSide.Left => 0, PublisherMasterPageSide.Right => width, _ => width / 2D
            } : width / 2D;
            var binding = new ObjectBinding(source, firstObjectId + (uint)index, RootShapeId + (uint)index + 1, story, ownerPageId, originX);
            objects.Add(binding);
            if (source.ImageName != null) namedPictures.Add(binding);
            if (source.Table != null) tables.Add(binding);
            else if (story != null) textFrames.Add(binding);
        }
        uint quillDocumentId = firstObjectId + (uint)objects.Count;
        uint printBlockId = quillDocumentId + 3;
        uint masterGeometryId = printBlockId + 1;
        uint colorSchemeId = masterGeometryId + masterCount;
        var pictureResources = new Dictionary<uint, uint>();
        for (int index = 0; index < namedPictures.Count; index++) {
            context.Token.ThrowIfCancellationRequested();
            pictureResources.Add(namedPictures[index].Id, colorSchemeId + 1 + (uint)index);
        }
        var tableCells = new Dictionary<uint, uint>();
        for (int index = 0; index < tables.Count; index++) {
            context.Token.ThrowIfCancellationRequested();
            tableCells.Add(tables[index].Id, colorSchemeId + 1 + (uint)namedPictures.Count + (uint)index);
        }
        uint colorSeparationId = colorSchemeId + 1 + (uint)namedPictures.Count + (uint)tables.Count;
        var pageReferences = new byte[orderedPageCount][];
        for (int index = 0; index < orderedPageCount; index++) {
            context.Token.ThrowIfCancellationRequested();
            pageReferences[index] = PublisherOplWriter.UInt32(0, firstPageId + (uint)index, 0x70);
        }
        byte[] references = PublisherOplWriter.Nested(2, 0xA0, pageReferences);
        byte[] document = PublisherOplWriter.Chunk(PublisherOplWriter.UInt32(1, orderedPageCount, 0x20), references, PublisherOplWriter.UInt32(3, printBlockId, 0x70),
            PublisherOplWriter.Nested(0x12, 0x88, PublisherOplWriter.UInt32(1, width, 0x20), PublisherOplWriter.UInt32(2, height, 0x20)),
            PublisherOplWriter.UInt32(0x20, quillDocumentId, 0x70),
            PublisherOplWriter.UInt32(0x22, colorSchemeId, 0x70),
            PublisherOplWriter.UInt16(0x2C, SpecialPageBoundary, 0x18), PublisherOplWriter.UInt16(0x2D, masterCount, 0x18));
        var chunks = new List<NativeRecord> { new(0x44, null, document) };
        for (int index = 0; index < orderedPageCount; index++) {
            context.Token.ThrowIfCancellationRequested();
            bool isMaster = index < masterCount;
            PublisherNativeCreationMaster? master = isMaster ? specification.MasterDefinitions[index] : null;
            var pageObjects = new List<byte[]>();
            foreach (ObjectBinding item in objects) {
                context.Token.ThrowIfCancellationRequested();
                if (item.OwnerPageId == firstPageId + (uint)index)
                    pageObjects.Add(PublisherOplWriter.UInt32(0, item.Id, 0x70));
            }
            var fields = new List<byte[]> {
                PublisherOplWriter.UInt32(1, (uint)pageObjects.Count, 0x20),
                PublisherOplWriter.Nested(2, 0xA0, pageObjects.ToArray())
            };
            if (isMaster) fields.Add(PublisherOplWriter.UInt32(3, masterGeometryId + (uint)index, 0x70));
            if (!isMaster) {
                int physicalIndex = index - masterCount;
                int? assigned = physicalIndex < specification.PhysicalPageCount && specification.PageMasters != null
                    ? specification.PageMasters[physicalIndex] : 0;
                if (assigned.HasValue) {
                    if (assigned.Value < 0 || assigned.Value >= masterCount) throw new ArgumentException("An applied master is outside the definitions.");
                    fields.Add(PublisherOplWriter.UInt32(0x0D, firstPageId + (uint)assigned.Value, 0x68));
                }
            }
            fields.Add(PublisherOplWriter.Nested(0x0E, 0xC0, Encoding.Unicode.GetBytes((master?.Code ?? string.Empty) + "\0")));
            fields.Add(PublisherOplWriter.Nested(0x0F, 0xC0, Encoding.Unicode.GetBytes((master?.Name ?? string.Empty) + "\0")));
            // Qualified native master registrations: left 0, right 1, centered 3.
            // Physical/utility descriptors retain the existing observed code 5.
            fields.Add(PublisherOplWriter.UInt32(0x10, master?.Side switch {
                PublisherMasterPageSide.Left => 0U, PublisherMasterPageSide.Right => 1U,
                PublisherMasterPageSide.Single => 3U, _ => 5U
            }, 0x20));
            chunks.Add(new NativeRecord(0x43, documentId, PublisherOplWriter.Chunk(fields.ToArray())));
        }
        foreach (ObjectBinding item in objects) {
            context.Token.ThrowIfCancellationRequested();
            var fields = new List<byte[]>();
            if (item.Story != null) {
                fields.Add(PublisherOplWriter.UInt32(0x27, item.Story.Key, 0x20));
                if (item.Source.Table == null) fields.Add(PublisherOplWriter.UInt32(0x28, item.Source.FrameOrder, 0x20));
            }
            if (item.Source.Table != null) fields.AddRange(PublisherNativeCreationTableRecords.Fields(item.Source.Table, tableCells[item.Id], context));
            if (item.Source.FrameLayout is { VerticalAlignment: not OfficeIMO.Drawing.OfficeTextVerticalAlignment.Top } layout)
                fields.Add(PublisherOplWriter.UInt32(0x35, (uint)layout.VerticalAlignment, 0x20));
            if (pictureResources.TryGetValue(item.Id, out uint resource)) fields.Add(PublisherOplWriter.UInt32(0x3A, resource, 0x70));
            chunks.Add(new NativeRecord(item.Source.NativeKind, item.OwnerPageId, PublisherOplWriter.Chunk(fields.ToArray())));
        }
        // The ordinary native loader obtains its Quill document through field
        // 0x20 and kind 0x5B. Its frame/story catalogs use generated rows for the
        // ordered text-bearing objects; publications without stories retain empty catalogs.
        chunks.Add(new NativeRecord(0x5B, documentId, PublisherOplWriter.Chunk(
            PublisherOplWriter.UInt32(1, quillDocumentId + 1, 0x70), PublisherOplWriter.UInt32(2, quillDocumentId + 2, 0x70),
            PublisherOplWriter.Nested(10, 0x98, PublisherOplWriter.UInt32(1, 0x08000000, 0x20)))));
        ObjectBinding[] frames = textFrames.OrderBy(item => item.Story!.Key).ThenBy(item => item.Source.FrameOrder).ToArray();
        var frameCatalog = new byte[frames.Length][];
        var frameCounts = new Dictionary<StoryBinding, int>();
        for (int index = 0; index < frames.Length; index++) {
            context.Token.ThrowIfCancellationRequested();
            ObjectBinding frame = frames[index];
            frameCatalog[index] = FrameCatalogEntry(frame);
            frameCounts.TryGetValue(frame.Story!, out int count);
            frameCounts[frame.Story!] = count + 1;
        }
        chunks.Add(new NativeRecord(0x61, quillDocumentId, frames.Length != 0
            ? PublisherOplWriter.Chunk(PublisherOplWriter.UInt32(1, (uint)frames.Length, 0x20), PublisherOplWriter.Nested(2, 0xA0, frameCatalog))
            : PublisherOplWriter.Chunk()));
        // Fresh creation assigns stable progressive layout identifiers. IdLayout
        // is separate from the frame reference count and from retained native keys.
        var storyCatalog = new byte[stories.Count][];
        for (int index = 0; index < stories.Count; index++) {
            context.Token.ThrowIfCancellationRequested();
            StoryBinding story = stories[index];
            frameCounts.TryGetValue(story, out int count);
            storyCatalog[index] = StoryCatalogEntry(story, story.Source.CellEnds == null ? count : 1);
        }
        chunks.Add(new NativeRecord(0x65, quillDocumentId, stories.Count != 0
            ? PublisherOplWriter.Chunk(PublisherOplWriter.UInt32(1, (uint)stories.Count, 0x20), PublisherOplWriter.Nested(2, 0xA0, storyCatalog))
            : PublisherOplWriter.Chunk()));
        // The normal native loader requires OhPrintBlock (document field 3).
        // Normal opening accesses row zero of Rgdnb. One empty typed OplDnb
        // supplies that row; its fields use the native decoder's defaults.
        chunks.Add(new NativeRecord(0x4B, documentId, PublisherOplWriter.Chunk(PublisherOplWriter.UInt32(1, 1, 0x20),
            PublisherOplWriter.Nested(2, 0x90, PublisherOplWriter.Nested(0, 0x88)), PublisherOplWriter.UInt32(3, colorSeparationId, 0x70))));
        for (int index = 0; index < masterCount; index++) {
            context.Token.ThrowIfCancellationRequested();
            chunks.Add(new NativeRecord(0x4C, firstPageId + (uint)index, PublisherNativeCreationMasterGeometry.Create(width, height)));
        }
        context.Token.ThrowIfCancellationRequested();
        chunks.Add(new NativeRecord(0x5C, documentId, PublisherNativeCreationColorScheme.CreateDefault()));
        foreach (ObjectBinding picture in namedPictures) {
            context.Token.ThrowIfCancellationRequested();
            chunks.Add(new NativeRecord(0x66, picture.Id, PublisherOplWriter.Chunk(PublisherOplWriter.Nested(3, 0xC0,
                Encoding.Unicode.GetBytes(picture.Source.ImageName + "\0")))));
        }
        foreach (ObjectBinding table in tables) {
            context.Token.ThrowIfCancellationRequested();
            chunks.Add(new NativeRecord(0x63, table.Id, PublisherNativeCreationTableRecords.Cells(table.Source.Table!, context)));
        }
        // Print preview dereferences OhColorSepBlock even for composite output.
        // One automatic-color plate uses native defaults without printer data.
        // DzlTrap/DzlIndTrap must be explicit: absent values decode to zero in
        // validation, below 127. The native constructor initializes both to 3175.
        chunks.Add(new NativeRecord(0x4F, printBlockId, PublisherOplWriter.Chunk(
            PublisherOplWriter.Nested(2, 0xA0, PublisherOplWriter.Nested(0, 0x88,
                PublisherOplWriter.Nested(2, 0x98, PublisherOplWriter.UInt32(1, uint.MaxValue, 0x20)))),
            PublisherOplWriter.UInt32(0x0C, 3175, 0x20), PublisherOplWriter.UInt32(0x0D, 3175, 0x20))));
        context.Token.ThrowIfCancellationRequested();
        return new Graph(specification, chunks, objects, stories);
    }

    internal sealed record Graph(PublisherNativeCreationPlan Source, IReadOnlyList<NativeRecord> Chunks,
        IReadOnlyList<ObjectBinding> Objects, IReadOnlyList<StoryBinding> Stories);
    internal sealed record ObjectBinding(PublisherNativeCreationObject Source, uint Id, uint Spid, StoryBinding? Story, uint OwnerPageId, double HorizontalOriginEmu);
    internal sealed record StoryBinding(PublisherNativeCreationStory Source, uint Key, uint LayoutId);
    internal sealed record NativeRecord(uint Kind, uint? Parent, byte[] Bytes);

    private static byte[] FrameCatalogEntry(ObjectBinding frame) {
        var fields = new List<byte[]> { PublisherOplWriter.UInt32(1, frame.Story!.Key, 0x20) };
        if (frame.Source.FrameOrder != 0) fields.Add(PublisherOplWriter.UInt32(2, frame.Source.FrameOrder, 0x20));
        fields.Add(PublisherOplWriter.UInt32(3, frame.Id, 0x68));
        return PublisherOplWriter.Nested(0, 0x88, fields.ToArray());
    }

    private static byte[] StoryCatalogEntry(StoryBinding story, int frameCount) {
        var fields = new List<byte[]> { PublisherOplWriter.UInt32(1, story.Key, 0x20) };
        // The native OplQsb schema calls field 2 CRef; omission defaults to one.
        // IdLayout remains an independent identifier in field 7.
        if (frameCount != 1) fields.Add(PublisherOplWriter.UInt16(2, checked((ushort)frameCount), 0x18));
        if (story.Source.CellEnds != null) fields.Add(PublisherOplWriter.UInt16(3, 0, 0x10)); // Table story role (Tdet).
        if (story.Source.AutomaticHyphenation is { } enabled) fields.Add(PublisherStoryHyphenation.EncodeFlag(enabled));
        fields.Add(PublisherOplWriter.UInt32(7, story.LayoutId, 0x20));
        return PublisherOplWriter.Nested(0, 0x88, fields.ToArray());
    }
}

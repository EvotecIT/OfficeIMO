using System.Text;
using static OfficeIMO.Publisher.WriterProbe.NativeEncoding;

namespace OfficeIMO.Publisher.WriterProbe;

/// <summary>
/// Synthesizes the document, explicit page roles and small Quill document graph.
/// The role boundaries follow recovered native rules; native opening remains unqualified.
/// </summary>
internal static class NativePublicationGraph {
    internal const ushort MasterPageCount = 1;
    internal const ushort SpecialPageBoundary = 5;
    internal const uint DrawingId = 1;
    internal const uint RootShapeId = DrawingId << 10;

    // Native enumeration admits M..N-S inclusive: physical = N-S-M+1.
    internal static ushort GetOrderedPageCount(CandidatePublicationSpec specification) =>
        checked((ushort)(specification.PhysicalPageCount + MasterPageCount + SpecialPageBoundary - 1));

    internal static Graph Create(CandidatePublicationSpec specification, uint documentId) {
        ushort orderedPageCount = GetOrderedPageCount(specification);
        uint width = specification.WidthEmu, height = specification.HeightEmu;
        uint firstPageId = documentId + 1;
        uint masterId = firstPageId;
        uint firstObjectId = firstPageId + orderedPageCount;
        var objects = new List<ObjectBinding>();
        var stories = new List<StoryBinding>();
        var storyBindings = new Dictionary<CandidateStorySpec, StoryBinding>();
        for (int index = 0; index < specification.Objects.Count; index++) {
            CandidateObjectSpec source = specification.Objects[index];
            StoryBinding? story = null;
            if (source.Story != null) {
                if (!storyBindings.TryGetValue(source.Story, out story)) {
                    uint key = (uint)stories.Count + 1;
                    story = new StoryBinding(source.Story, key, key);
                    stories.Add(story);
                    storyBindings.Add(source.Story, story);
                }
            }
            uint ownerPageId = firstPageId + (source.OnMaster ? 0 : checked((uint)(MasterPageCount + source.PageIndex)));
            objects.Add(new ObjectBinding(source, firstObjectId + (uint)index, RootShapeId + (uint)index + 1, story, ownerPageId));
        }
        uint quillDocumentId = firstObjectId + (uint)objects.Count;
        uint printBlockId = quillDocumentId + 3;
        uint masterGeometryId = printBlockId + 1;
        uint colorSchemeId = masterGeometryId + 1;
        ObjectBinding[] namedPictures = objects.Where(item => item.Source.ImageName != null).ToArray();
        var pictureResources = namedPictures.Select((item, index) => (item.Id, ResourceId: colorSchemeId + 1 + (uint)index))
            .ToDictionary(pair => pair.Id, pair => pair.ResourceId);
        byte[] references = Variable(2, 0xA0, Enumerable.Range(0, orderedPageCount)
            .Select(index => Scalar(0, 0x70, firstPageId + (uint)index)).ToArray());
        byte[] document = Chunk(Scalar(1, 0x20, orderedPageCount), references, Scalar(3, 0x70, printBlockId),
            Variable(0x12, 0x88, Scalar(1, 0x20, width), Scalar(2, 0x20, height)),
            Scalar(0x20, 0x70, quillDocumentId),
            Scalar(0x22, 0x70, colorSchemeId),
            ShortScalar(0x2C, 0x18, SpecialPageBoundary), ShortScalar(0x2D, 0x18, MasterPageCount));
        var chunks = new List<NativeRecord> { new(0x44, null, document) };
        for (int index = 0; index < orderedPageCount; index++) {
            bool isMaster = index < MasterPageCount;
            ObjectBinding[] pageObjects = objects.Where(item => item.OwnerPageId == firstPageId + (uint)index).ToArray();
            var fields = new List<byte[]> {
                Scalar(1, 0x20, (uint)pageObjects.Length),
                Variable(2, 0xA0, pageObjects.Select(item => Scalar(0, 0x70, item.Id)).ToArray())
            };
            if (isMaster) fields.Add(Scalar(3, 0x70, masterGeometryId));
            if (!isMaster) fields.Add(Scalar(0x0D, 0x68, masterId));
            fields.Add(Variable(0x0E, 0xC0, Encoding.Unicode.GetBytes(isMaster ? "A\0" : "\0")));
            fields.Add(Variable(0x0F, 0xC0, Encoding.Unicode.GetBytes(isMaster ? "Master Page A\0" : "\0")));
            // Observed page-kind trait: 3 on the named master, 5 on the other
            // native blank descriptors. No broader semantic meaning is assumed.
            fields.Add(Scalar(0x10, 0x20, isMaster ? 3U : 5U));
            chunks.Add(new NativeRecord(0x43, documentId, Chunk(fields.ToArray())));
        }
        foreach (ObjectBinding item in objects) {
            var fields = new List<byte[]>();
            if (item.Story != null) {
                fields.Add(Scalar(0x27, 0x20, item.Story.Key)); fields.Add(Scalar(0x28, 0x20, item.Source.FrameOrder));
            }
            if (pictureResources.TryGetValue(item.Id, out uint resource)) fields.Add(Scalar(0x3A, 0x70, resource));
            chunks.Add(new NativeRecord(item.Source.NativeKind, item.OwnerPageId, Chunk(fields.ToArray())));
        }
        // The ordinary native loader obtains its Quill document through field
        // 0x20 and kind 0x5B. Its frame/story catalogs use generated rows for the
        // ordered text-bearing objects; publications without stories retain empty catalogs.
        chunks.Add(new NativeRecord(0x5B, documentId, Chunk(
            Scalar(1, 0x70, quillDocumentId + 1), Scalar(2, 0x70, quillDocumentId + 2),
            Variable(10, 0x98, Scalar(1, 0x20, 0x08000000)))));
        ObjectBinding[] frames = objects.Where(item => item.Story != null)
            .OrderBy(item => item.Story!.Key).ThenBy(item => item.Source.FrameOrder).ToArray();
        chunks.Add(new NativeRecord(0x61, quillDocumentId, frames.Length != 0
            ? Chunk(Scalar(1, 0x20, (uint)frames.Length), Variable(2, 0xA0, frames
                .Select(item => FrameCatalogEntry(item)).ToArray()))
            : Chunk()));
        // Fresh creation assigns stable progressive layout identifiers. IdLayout
        // is separate from the frame reference count and from retained native keys.
        chunks.Add(new NativeRecord(0x65, quillDocumentId, stories.Count != 0
            ? Chunk(Scalar(1, 0x20, (uint)stories.Count), Variable(2, 0xA0, stories.Select(story =>
                StoryCatalogEntry(story, frames.Count(frame => frame.Story == story))).ToArray()))
            : Chunk()));
        // The normal native loader requires OhPrintBlock (document field 3).
        // Normal opening accesses row zero of Rgdnb. One empty typed OplDnb
        // supplies that row; its fields use the native decoder's defaults.
        chunks.Add(new NativeRecord(0x4B, documentId, Chunk(Scalar(1, 0x20, 1),
            Variable(2, 0x90, Variable(0, 0x88)))));
        chunks.Add(new NativeRecord(0x4C, masterId, CandidateMasterGeometry.Create(width, height)));
        chunks.Add(new NativeRecord(0x5C, documentId, CandidateColorScheme.CreateDefault()));
        foreach (ObjectBinding picture in namedPictures)
            chunks.Add(new NativeRecord(0x66, picture.Id, Chunk(Variable(3, 0xC0,
                Encoding.Unicode.GetBytes(picture.Source.ImageName + "\0")))));
        return new Graph(specification, chunks, objects, stories);
    }

    internal sealed record Graph(CandidatePublicationSpec Source, IReadOnlyList<NativeRecord> Chunks,
        IReadOnlyList<ObjectBinding> Objects, IReadOnlyList<StoryBinding> Stories);
    internal sealed record ObjectBinding(CandidateObjectSpec Source, uint Id, uint Spid, StoryBinding? Story, uint OwnerPageId);
    internal sealed record StoryBinding(CandidateStorySpec Source, uint Key, uint LayoutId);
    internal sealed record NativeRecord(uint Kind, uint? Parent, byte[] Bytes);

    private static byte[] FrameCatalogEntry(ObjectBinding frame) {
        var fields = new List<byte[]> { Scalar(1, 0x20, frame.Story!.Key) };
        if (frame.Source.FrameOrder != 0) fields.Add(Scalar(2, 0x20, frame.Source.FrameOrder));
        fields.Add(Scalar(3, 0x68, frame.Id));
        return Variable(0, 0x88, fields.ToArray());
    }

    private static byte[] StoryCatalogEntry(StoryBinding story, int frameCount) {
        var fields = new List<byte[]> { Scalar(1, 0x20, story.Key) };
        // The native OplQsb schema calls field 2 CRef; omission defaults to one.
        // IdLayout remains an independent identifier in field 7.
        if (frameCount != 1) fields.Add(ShortScalar(2, 0x18, checked((ushort)frameCount)));
        fields.Add(Scalar(7, 0x20, story.LayoutId));
        return Variable(0, 0x88, fields.ToArray());
    }
}

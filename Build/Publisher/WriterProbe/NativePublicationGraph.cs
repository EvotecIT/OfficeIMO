using System.Text;
using static OfficeIMO.Publisher.WriterProbe.NativeEncoding;

namespace OfficeIMO.Publisher.WriterProbe;

/// <summary>
/// Synthesizes the document, explicit page roles and small Quill document graph.
/// The role boundaries follow recovered native rules; native opening remains unqualified.
/// </summary>
internal static class NativePublicationGraph {
    internal const ushort OrderedPageCount = 6;
    internal const ushort MasterPageCount = 1;
    internal const ushort SpecialPageBoundary = 5;
    private const int PhysicalPageIndex = 1;

    internal static Graph Create(string kind, uint documentId, uint width, uint height) {
        uint firstPageId = documentId + 1;
        uint masterId = firstPageId;
        uint physicalId = firstPageId + PhysicalPageIndex;
        uint shapeId = firstPageId + OrderedPageCount;
        bool hasShape = kind != "blank";
        uint quillDocumentId = shapeId + (hasShape ? 1U : 0U);
        uint printBlockId = quillDocumentId + 3;
        uint masterGeometryId = printBlockId + 1;
        byte[] references = Variable(2, 0xA0, Enumerable.Range(0, OrderedPageCount)
            .Select(index => Scalar(0, 0x70, firstPageId + (uint)index)).ToArray());
        byte[] document = Chunk(Scalar(1, 0x20, OrderedPageCount), references, Scalar(3, 0x70, printBlockId),
            Variable(0x12, 0x88, Scalar(1, 0x20, width), Scalar(2, 0x20, height)),
            Scalar(0x20, 0x70, quillDocumentId),
            ShortScalar(0x2C, 0x18, SpecialPageBoundary), ShortScalar(0x2D, 0x18, MasterPageCount));
        var chunks = new List<NativeRecord> { new(0x44, null, document) };
        for (int index = 0; index < OrderedPageCount; index++) {
            bool isMaster = index < MasterPageCount;
            bool containsShape = index == PhysicalPageIndex && hasShape;
            var fields = new List<byte[]> {
                Scalar(1, 0x20, containsShape ? 1U : 0U),
                Variable(2, 0xA0, containsShape ? Scalar(0, 0x70, shapeId) : Array.Empty<byte>())
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
        if (hasShape) chunks.Add(new NativeRecord(kind == "image" ? 0x20U : 1U, physicalId,
            kind == "text" ? Chunk(Scalar(0x27, 0x20, 1), Scalar(0x28, 0x20, 0)) : Chunk()));
        // The ordinary native loader obtains its Quill document through field
        // 0x20 and kind 0x5B. The two empty children and the option value below
        // are decoded fields from an independently authored blank publication.
        chunks.Add(new NativeRecord(0x5B, documentId, Chunk(
            Scalar(1, 0x70, quillDocumentId + 1), Scalar(2, 0x70, quillDocumentId + 2),
            Variable(10, 0x98, Scalar(1, 0x20, 0x08000000)))));
        chunks.Add(new NativeRecord(0x61, quillDocumentId, Chunk()));
        chunks.Add(new NativeRecord(0x65, quillDocumentId, Chunk()));
        // The normal native loader requires OhPrintBlock (document field 3).
        // Normal opening accesses row zero of Rgdnb. One empty typed OplDnb
        // supplies that row; its fields use the native decoder's defaults.
        chunks.Add(new NativeRecord(0x4B, documentId, Chunk(Scalar(1, 0x20, 1),
            Variable(2, 0x90, Variable(0, 0x88)))));
        chunks.Add(new NativeRecord(0x4C, masterId, CandidateMasterGeometry.Create(width, height)));
        return new Graph(chunks, shapeId);
    }

    internal sealed record Graph(IReadOnlyList<NativeRecord> Chunks, uint ShapeId);
    internal sealed record NativeRecord(uint Kind, uint? Parent, byte[] Bytes);
}

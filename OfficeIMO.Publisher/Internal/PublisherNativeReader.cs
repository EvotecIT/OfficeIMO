using OfficeIMO.Core.Internal;

namespace OfficeIMO.Publisher.Internal;

internal static class PublisherNativeReader {
    internal static PublisherDocument Read(OfficeCompoundFile compound, PublisherReadOptions options, CancellationToken token) {
        var context = new PublisherParseContext(options, token);
        PublisherBinaryData contents = Required(compound, "Contents");
        contents.Range(0, 4);
        if (contents.U8(0) != 0xE8 || contents.U8(1) != 0xAC || contents.U8(3) != 0)
            throw new InvalidDataException("The compound document is not a recognized Publisher publication.");
        if (contents.U8(2) == 0x22)
            throw new NotSupportedException("The earlier Publisher 97/98 or 2000 Contents generation is not decoded. Its shared signature does not establish the exact producer release.");
        if (contents.U8(2) != 0x2C) throw new NotSupportedException("Unsupported Publisher native version marker.");
        PublisherContents publication = new PublisherContentsReader(contents, context).Read();
        PublisherQuillText text = new PublisherQuillTextReader(Required(compound, "Quill/QuillSub/CONTENTS"), context).Read(publication.Palette);
        PublisherEscherData escher = new PublisherEscherReader(Required(compound, "Escher/EscherStm"), Optional(compound, "Escher/EscherDelayStm"), context).Read();
        context.Add("PUB_NATIVE_FIELDS_UNASSESSED", "This bounded native recovery does not establish field-level fidelity for every publication record, auxiliary stream, or embedded object.",
            OfficeConversionLossKind.Unassessed, "Contents");
        foreach (string stream in compound.Streams.Keys) {
            if (stream.StartsWith("VBA/", StringComparison.OrdinalIgnoreCase) || stream.StartsWith("Objects/", StringComparison.OrdinalIgnoreCase))
                context.Add("PUB_ACTIVE_CONTENT_OMITTED", "Macros and embedded application objects remain inert and are not projected onto publication pages.", OfficeConversionLossKind.Omission, stream);
        }
        return new PublisherPageProjector(publication, text, escher, context).Project();
    }
    private static PublisherBinaryData Required(OfficeCompoundFile compound, string name) =>
        Optional(compound, name) ?? throw new InvalidDataException("Required Publisher stream is missing: " + name);
    private static PublisherBinaryData? Optional(OfficeCompoundFile compound, string name) =>
        compound.Streams.TryGetValue(name, out byte[]? bytes) ? new PublisherBinaryData(bytes, name) : null;
}

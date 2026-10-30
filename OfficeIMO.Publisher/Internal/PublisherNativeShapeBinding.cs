namespace OfficeIMO.Publisher.Internal;

/// <summary>Exact source slots for fixed-size edits. All offsets address stream bodies, never CFB sectors or record headers.</summary>
internal sealed class PublisherNativeShapeBinding {
    internal PublisherNativeShapeBinding(uint id, uint spid, int bodyOffset, int bodyLength, int spidOffset,
        PublisherNativeValueSlot publicationId, PublisherNativeValueSlot[] anchor, IReadOnlyList<PublisherNativePropertySlot> properties) {
        Id = id; Spid = spid; BodyOffset = bodyOffset; BodyLength = bodyLength; SpidOffset = spidOffset;
        PublicationId = publicationId; Anchor = Array.AsReadOnly(anchor); Properties = Array.AsReadOnly(properties.ToArray());
    }
    internal uint Id { get; }
    internal uint Spid { get; }
    internal int BodyOffset { get; }
    internal int BodyLength { get; }
    internal int SpidOffset { get; }
    internal PublisherNativeValueSlot PublicationId { get; }
    internal IReadOnlyList<PublisherNativeValueSlot> Anchor { get; }
    internal IReadOnlyList<PublisherNativePropertySlot> Properties { get; }

    internal bool Matches(PublisherBinaryData source) => source.U16(BodyOffset - 6) == 0xF004
        && source.U32(BodyOffset - 4) == BodyLength && source.U16(SpidOffset - 6) == 0xF00A && source.U32(SpidOffset) == Spid
        && source.U16(PublicationId.Offset - 2) == 0x6801 && source.U32(PublicationId.Offset) == Id
        && Anchor.Select((slot, index) => source.U16(slot.Offset - 2) == 0x2001 + index && source.U32(slot.Offset) == slot.Original).All(value => value)
        && Properties.All(slot => source.U16(slot.Value.Offset - 2) == slot.RawOperationId && source.U32(slot.Value.Offset) == slot.Value.Original);
}

internal readonly record struct PublisherNativeValueSlot(int Offset, uint Original);
internal readonly record struct PublisherNativePropertySlot(ushort RawOperationId, PublisherNativeValueSlot Value) {
    internal ushort Id => (ushort)(RawOperationId & 0x3FFF);
}

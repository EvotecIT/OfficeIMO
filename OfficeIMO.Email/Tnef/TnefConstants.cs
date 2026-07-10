namespace OfficeIMO.Email;

internal static class TnefConstants {
    internal const uint Signature = 0x223E9F78;
    internal const uint Version = 0x00010000;
    internal const uint AttachRendData = 0x00069002;
    internal const uint AttachTransportFilename = 0x00069001;
    internal const uint AttachData = 0x0006800F;
    internal const uint AttachTitle = 0x00018010;
    internal const uint MessageClass = 0x00078008;
    internal const uint Subject = 0x00018004;
    internal const uint DateSent = 0x00038005;
    internal const uint DateReceived = 0x00038006;
    internal const uint MessageId = 0x00018009;
    internal const uint Body = 0x0002800C;
    internal const uint MessageStatus = 0x00068007;
    internal const uint MessageProperties = 0x00069003;
    internal const uint RecipientTable = 0x00069004;
    internal const uint AttachmentProperties = 0x00069005;
    internal const uint TnefVersion = 0x00089006;
    internal const uint OemCodePage = 0x00069007;
}

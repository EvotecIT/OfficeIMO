namespace OfficeIMO.Pdf;

internal static class PdfDirectStreamIdentity {
    private const int MaximumDepth = 64;
    private const int MaximumNodes = 16_384;

    internal static int Compute(PdfStream stream) {
        using var fingerprint = new PdfObjectGraphFingerprint(
            new Dictionary<int, PdfIndirectObject>(),
            MaximumDepth,
            MaximumNodes,
            preserveUnresolvedReferenceIdentity: true);
        fingerprint.AppendRoot(stream);
        byte[] digest = fingerprint.Complete();
        int identity = digest[0] | digest[1] << 8 | digest[2] << 16 | digest[3] << 24;
        return identity == 0 ? 1 : identity;
    }
}

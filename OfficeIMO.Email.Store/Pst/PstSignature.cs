namespace OfficeIMO.Email.Store;

internal static class PstSignature {
    internal static ushort Compute(long offset, ulong bid) {
        uint mixed = unchecked((uint)offset) ^ unchecked((uint)bid);
        return unchecked((ushort)((mixed >> 16) ^ (mixed & 0xFFFF)));
    }
}

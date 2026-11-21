using System.Runtime.CompilerServices;

namespace OfficeIMO.Pdf;

public sealed partial class PdfEmbeddedFontFamily {
    // Repeated family construction from one live caller array should retain the
    // same private snapshot, allowing the weak font-program cache to reuse parsing.
    // Retain only the latest requested family name so name variants cannot keep
    // earlier parsing contexts alive through a shared snapshot's cache bucket.
    // Values never reference their caller key; releasing it also releases this reuse.
    private static readonly ConditionalWeakTable<byte[], FaceSnapshot> FaceSnapshots = new();

    private static byte[] SnapshotFace(byte[] caller, string requestedFamilyName) {
        FaceSnapshot snapshot = FaceSnapshots.GetValue(caller, static _ => new FaceSnapshot());
        lock (snapshot) {
            byte[]? previous = snapshot.Data;
            if (previous == null || !string.Equals(snapshot.RequestedFamilyName, requestedFamilyName, StringComparison.Ordinal)
                || !MatchesCaller(previous, caller)) {
                // The caller can edit or reuse its buffer between constructions.
                // Never expose that buffer or change an earlier family's snapshot.
                snapshot.Data = (byte[])caller.Clone();
                snapshot.RequestedFamilyName = requestedFamilyName;
            }
            return snapshot.Data!;
        }
    }

    private static bool MatchesCaller(byte[] snapshot, byte[] caller) {
#if NET8_0_OR_GREATER
        return snapshot.AsSpan().SequenceEqual(caller);
#else
        // Older supported targets do not require a span compatibility package. Compare eight bytes at a time:
        // a font is hundreds of kilobytes and this runs for every family constructed from the same array.
        if (snapshot.Length != caller.Length) return false;
        unsafe {
            fixed (byte* left = snapshot, right = caller) {
                int index = 0;
                for (int words = snapshot.Length / 8; words > 0; words--, index += 8) {
                    if (*(ulong*)(left + index) != *(ulong*)(right + index)) return false;
                }
                for (; index < snapshot.Length; index++) {
                    if (left[index] != right[index]) return false;
                }
            }
        }
        return true;
#endif
    }

    private sealed class FaceSnapshot {
        internal byte[]? Data { get; set; }
        internal string? RequestedFamilyName { get; set; }
    }
}

using System.Runtime.CompilerServices;

namespace OfficeIMO.Pdf;

public sealed partial class PdfEmbeddedFontFamily {
    // Repeated family construction from one live caller array should retain the
    // same private snapshot, allowing the weak font-program cache to reuse parsing.
    // Values never reference their caller key; releasing it also releases this reuse.
    private static readonly ConditionalWeakTable<byte[], FaceSnapshot> FaceSnapshots = new();

    private static byte[] SnapshotFace(byte[] caller) {
        FaceSnapshot snapshot = FaceSnapshots.GetValue(caller, static _ => new FaceSnapshot());
        lock (snapshot) {
            byte[]? previous = snapshot.Data;
            if (previous == null || !previous.AsSpan().SequenceEqual(caller)) {
                // The caller can edit or reuse its buffer between constructions.
                // Never expose that buffer or change an earlier family's snapshot.
                snapshot.Data = (byte[])caller.Clone();
            }
            return snapshot.Data!;
        }
    }

    private sealed class FaceSnapshot {
        internal byte[]? Data { get; set; }
    }
}

using System;

namespace OfficeIMO.Drawing;

/// <summary>Leaf state preceding AV1 intra prediction syntax. Filter deltas are a snapshot, not shared tile arrays.</summary>
internal readonly struct OfficeAv1BlockPrelude {
    internal OfficeAv1BlockPrelude(bool skip, int segment, bool lossless, int cdef, int q, int[] filters) {
        Skip = skip; SegmentId = segment; Lossless = lossless; CdefIndex = cdef; CurrentQIndex = q;
        Filter0 = filters[0]; Filter1 = filters[1]; Filter2 = filters[2]; Filter3 = filters[3];
    }
    internal bool Skip { get; }
    internal int SegmentId { get; }
    internal bool Lossless { get; }
    internal int CdefIndex { get; }
    internal int CurrentQIndex { get; }
    internal int Filter0 { get; }
    internal int Filter1 { get; }
    internal int Filter2 { get; }
    internal int Filter3 { get; }
}

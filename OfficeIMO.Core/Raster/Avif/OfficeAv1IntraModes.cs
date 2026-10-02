namespace OfficeIMO.Drawing;

/// <summary>AV1 intra mode numbers. ChromaFromLuma is a chroma-only mode.</summary>
internal enum OfficeAv1IntraMode {
    Dc, Vertical, Horizontal, Diagonal45, Diagonal135, Diagonal113, Diagonal157, Diagonal203, Diagonal67,
    Smooth, SmoothVertical, SmoothHorizontal, Paeth, ChromaFromLuma
}

/// <summary>Intra mode syntax up to the palette boundary. Intra-block-copy motion and later leaf syntax are separate.</summary>
internal readonly struct OfficeAv1IntraModes {
    internal OfficeAv1IntraModes(bool copy, bool chroma, bool cflAllowed, OfficeAv1IntraMode y, OfficeAv1IntraMode uv,
        int angleY, int angleUv, int alphaU, int alphaV) {
        UseIntraBlockCopy=copy; HasChroma=chroma; CflAllowed=cflAllowed; YMode=y; UvMode=uv;
        AngleDeltaY=angleY; AngleDeltaUv=angleUv; CflAlphaU=alphaU; CflAlphaV=alphaV;
    }
    internal bool UseIntraBlockCopy { get; }
    internal bool HasChroma { get; }
    internal bool CflAllowed { get; }
    internal OfficeAv1IntraMode YMode { get; }
    internal OfficeAv1IntraMode UvMode { get; }
    internal int AngleDeltaY { get; }
    internal int AngleDeltaUv { get; }
    internal int CflAlphaU { get; }
    internal int CflAlphaV { get; }
}

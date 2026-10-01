using System;

namespace OfficeIMO.Drawing;

/// <summary>One immutable AV1 restoration unit, before the post-reconstruction filter is applied.</summary>
/// <remarks>Type follows frame-header order: none=0, Wiener=2, self-guided=3. Wiener taps
/// are the three symmetric outer coefficients per pass; the center is derived during filtering.</remarks>
internal readonly struct OfficeAv1RestorationUnit {
    internal OfficeAv1RestorationUnit(int plane, int row, int col, int type, int set,
        int v0, int v1, int v2, int h0, int h1, int h2, int x0, int x1) {
        Plane=plane; Row=row; Col=col; Type=type; SgrSet=set;
        _v0=v0; _v1=v1; _v2=v2; _h0=h0; _h1=h1; _h2=h2; X0=x0; X1=x1;
    }
    private readonly int _v0, _v1, _v2, _h0, _h1, _h2;
    internal int Plane { get; }
    internal int Row { get; }
    internal int Col { get; }
    internal int Type { get; }
    internal int SgrSet { get; }
    internal int X0 { get; }
    internal int X1 { get; }
    internal int WienerTap(int pass, int tap) => pass==0
        ? tap==0?_v0:tap==1?_v1:tap==2?_v2:throw new ArgumentOutOfRangeException(nameof(tap))
        : pass==1 ? tap==0?_h0:tap==1?_h1:tap==2?_h2:throw new ArgumentOutOfRangeException(nameof(tap))
        : throw new ArgumentOutOfRangeException(nameof(pass));
}

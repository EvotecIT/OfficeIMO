namespace OfficeIMO.Drawing;

/// <summary>Validated item locations and properties borrowed from one bounded AVIF payload.</summary>
internal sealed class OfficeAvifImageItem {
    internal OfficeAvifImageItem(uint id, int width, int height, int offset, int length,
        byte[] configuration, bool isAlpha, OfficeAvifColorDescription? colorDescription) {
        Id = id;
        Width = width;
        Height = height;
        Offset = offset;
        Length = length;
        Configuration = configuration;
        IsAlpha = isAlpha;
        ColorDescription = colorDescription;
    }

    internal uint Id { get; }
    internal int Width { get; }
    internal int Height { get; }
    internal int Offset { get; }
    internal int Length { get; }
    internal byte[] Configuration { get; }
    internal bool Monochrome => Configuration.Length >= 3 && (Configuration[2] & 16) != 0;
    /// <summary>Main-profile sample depth retained independently of auxiliary-alpha role.</summary>
    internal int BitDepth => Configuration.Length >= 3 && (Configuration[2] & 64) != 0 ? 10 : 8;
    /// <summary>Auxiliary-alpha role is established by item references and auxC, independently of plane count.</summary>
    internal bool IsAlpha { get; }
    internal OfficeAvifColorDescription? ColorDescription { get; }
}

/// <summary>CICP values from nclx; absence leaves interpretation to the AV1 sequence header.</summary>
internal sealed class OfficeAvifColorDescription {
    internal OfficeAvifColorDescription(int primaries, int transfer, int matrix, bool fullRange) {
        Primaries = primaries;
        Transfer = transfer;
        Matrix = matrix;
        FullRange = fullRange;
    }
    internal int Primaries { get; }
    internal int Transfer { get; }
    internal int Matrix { get; }
    internal bool FullRange { get; }
}

/// <summary>Color and optional auxiliary-alpha items; this is container validation, not decoded pixels.</summary>
internal sealed class OfficeAvifContainer {
    internal OfficeAvifContainer(OfficeAvifImageItem color, OfficeAvifImageItem? alpha) {
        Color = color;
        Alpha = alpha;
    }

    internal OfficeAvifImageItem Color { get; }
    internal OfficeAvifImageItem? Alpha { get; }
}

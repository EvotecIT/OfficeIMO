namespace OfficeIMO.Drawing;

/// <summary>The embedding licence level that OS/2 fsType bits 0–3 resolve to.</summary>
internal enum OfficeFontEmbeddingLevel {
    /// <summary>No restriction (fsType bits 0–3 clear).</summary>
    Installable,
    /// <summary>Editable embedding (bit 3).</summary>
    Editable,
    /// <summary>Preview &amp; Print embedding (bit 2): the receiving document must be opened read-only.</summary>
    PreviewAndPrint,
    /// <summary>Restricted License embedding (bit 1): the font must not be embedded.</summary>
    Restricted
}

/// <summary>
/// The embedding permissions a font declares in its OS/2 fsType field, shared by every OfficeIMO component that embeds
/// a font program in an output document.
/// </summary>
/// <remarks>
/// When several of bits 1–3 are set, the least restrictive permission applies, as the OpenType OS/2 specification
/// directs for fonts that set more than one. A font without an OS/2 table, or with one too short to hold fsType, declares
/// no restriction; OS/2 is optional in Apple TrueType fonts and this matches how font tools read such fonts.
/// </remarks>
internal readonly struct OfficeFontEmbeddingPermissions {
    internal const ushort RestrictedLicenseFlag = 0x0002;
    internal const ushort PreviewAndPrintFlag = 0x0004;
    internal const ushort EditableFlag = 0x0008;
    internal const ushort NoSubsettingFlag = 0x0100;
    internal const ushort BitmapOnlyFlag = 0x0200;
    private const int FsTypeOffset = 8;

    private OfficeFontEmbeddingPermissions(ushort fsType, bool declared) {
        FsType = fsType;
        IsDeclared = declared;
        Level = (fsType & EditableFlag) != 0 ? OfficeFontEmbeddingLevel.Editable
            : (fsType & PreviewAndPrintFlag) != 0 ? OfficeFontEmbeddingLevel.PreviewAndPrint
            : (fsType & RestrictedLicenseFlag) != 0 ? OfficeFontEmbeddingLevel.Restricted
            : OfficeFontEmbeddingLevel.Installable;
    }

    /// <summary>Permissions of a font that declares no fsType.</summary>
    internal static OfficeFontEmbeddingPermissions Undeclared => new(0, declared: false);

    /// <summary>The raw fsType value, or 0 when the font has no usable OS/2 table.</summary>
    internal ushort FsType { get; }

    /// <summary>True when the font has an OS/2 table long enough to hold fsType.</summary>
    internal bool IsDeclared { get; }

    internal OfficeFontEmbeddingLevel Level { get; }

    /// <summary>Bit 8: the font must be embedded whole, never as a subset.</summary>
    internal bool NoSubsetting => (FsType & NoSubsettingFlag) != 0;

    /// <summary>Bit 9: only the font's bitmaps may be embedded, never its outlines.</summary>
    internal bool BitmapOnly => (FsType & BitmapOnlyFlag) != 0;

    /// <summary>True when the font's outlines may be embedded.</summary>
    internal bool AllowsOutlineEmbedding => Level != OfficeFontEmbeddingLevel.Restricted && !BitmapOnly;

    /// <summary>Why the outlines must not be embedded, or null when they may be.</summary>
    internal string? OutlineEmbeddingRefusal =>
        Level == OfficeFontEmbeddingLevel.Restricted
            ? "the font's OS/2 fsType sets Restricted License embedding without Preview & Print or Editable, so its licence forbids embedding it"
            : BitmapOnly
                ? "the font's OS/2 fsType permits embedding only its bitmaps, and OfficeIMO embeds outlines"
                : null;

    /// <summary>Resolves an fsType value.</summary>
    internal static OfficeFontEmbeddingPermissions FromFsType(ushort fsType) => new(fsType, declared: true);

    /// <summary>
    /// Reads fsType from an OS/2 table at <paramref name="os2Offset"/> whose declared length is <paramref name="os2Length"/>.
    /// OS/2 versions 0 and 1 define only bits 0–3, so their other bits are ignored, as the specification directs.
    /// </summary>
    internal static OfficeFontEmbeddingPermissions FromOs2Table(byte[] data, int os2Offset, int os2Length) {
        ArgumentNullException.ThrowIfNull(data);
        if (os2Length < FsTypeOffset + 2 || os2Offset < 0 || os2Offset > data.Length - (FsTypeOffset + 2)) return Undeclared;
        int version = (data[os2Offset] << 8) | data[os2Offset + 1];
        ushort fsType = unchecked((ushort)((data[os2Offset + FsTypeOffset] << 8) | data[os2Offset + FsTypeOffset + 1]));
        return FromFsType(version < 2 ? (ushort)(fsType & 0x000F) : fsType);
    }

    /// <summary>Reads one face of a font or font collection, or returns null when it is not a readable OpenType face.</summary>
    internal static OfficeFontEmbeddingPermissions? TryRead(byte[] data, int faceIndex = 0) {
        OfficeOpenTypeReader? reader = OfficeOpenTypeReader.TryCreateFace(data, faceIndex);
        if (reader == null) return null;
        return reader.TryGetTable("OS/2", out int offset, out int length) ? FromOs2Table(data, offset, length) : Undeclared;
    }

    /// <summary>
    /// Reads fsType through the sfnt table directory alone, for a single-face program that a font engine has already
    /// accepted but that lacks tables <see cref="OfficeOpenTypeReader"/> requires. Returns null when the directory is invalid.
    /// </summary>
    internal static OfficeFontEmbeddingPermissions? TryReadTableDirectory(byte[] data) {
        ArgumentNullException.ThrowIfNull(data);
        if (data.Length < 12) return null;
        uint flavor = ((uint)data[0] << 24) | ((uint)data[1] << 16) | ((uint)data[2] << 8) | data[3];
        if (flavor != 0x00010000 && flavor != 0x74727565 && flavor != 0x4F54544F) return null;
        int count = (data[4] << 8) | data[5];
        if (count == 0 || data.Length < 12 + count * 16) return null;
        for (int index = 0; index < count; index++) {
            int record = 12 + index * 16;
            if (data[record] != 'O' || data[record + 1] != 'S' || data[record + 2] != '/' || data[record + 3] != '2') continue;
            long offset = ((long)data[record + 8] << 24) | ((long)data[record + 9] << 16) | ((long)data[record + 10] << 8) | data[record + 11];
            long length = ((long)data[record + 12] << 24) | ((long)data[record + 13] << 16) | ((long)data[record + 14] << 8) | data[record + 15];
            if (offset > data.Length || length > data.Length - offset) return null;
            return FromOs2Table(data, (int)offset, (int)length);
        }
        return Undeclared;
    }

    /// <summary>
    /// Reads every face of a font file that is embedded whole. A collection takes the most restrictive level and every
    /// subsetting or bitmap restriction of its readable faces. Returns null when no face is a readable OpenType face.
    /// </summary>
    internal static OfficeFontEmbeddingPermissions? TryReadFile(byte[] data) {
        ArgumentNullException.ThrowIfNull(data);
        int count = OfficeOpenTypeReader.GetCollectionFaceCount(data);
        OfficeFontEmbeddingPermissions? combined = null;
        for (int face = 0; face < count; face++) {
            if (TryRead(data, face) is not OfficeFontEmbeddingPermissions permissions) continue;
            combined = combined is OfficeFontEmbeddingPermissions current ? MostRestrictive(current, permissions) : permissions;
        }
        return combined;
    }

    /// <summary>
    /// Throws <see cref="ArgumentException"/> when a whole font file must not be embedded. Data that is not a readable
    /// OpenType font or collection carries no licence to check and is accepted.
    /// </summary>
    internal static void EnsureFileEmbeddingAllowed(byte[] data, string paramName) {
        if (TryReadFile(data)?.OutlineEmbeddingRefusal is string refusal) {
            throw new ArgumentException("The font cannot be embedded: " + refusal + ".", paramName);
        }
    }

    private static OfficeFontEmbeddingPermissions MostRestrictive(OfficeFontEmbeddingPermissions left, OfficeFontEmbeddingPermissions right) {
        OfficeFontEmbeddingPermissions level = left.Level >= right.Level ? left : right;
        ushort extra = (ushort)((left.FsType | right.FsType) & (NoSubsettingFlag | BitmapOnlyFlag));
        ushort levelBits = (ushort)(level.FsType & (RestrictedLicenseFlag | PreviewAndPrintFlag | EditableFlag));
        return new OfficeFontEmbeddingPermissions((ushort)(levelBits | extra), left.IsDeclared || right.IsDeclared);
    }
}

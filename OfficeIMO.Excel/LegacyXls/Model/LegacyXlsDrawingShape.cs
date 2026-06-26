namespace OfficeIMO.Excel.LegacyXls.Model {
    /// <summary>
    /// Describes shallow metadata decoded from an OfficeArtFSP shape record.
    /// </summary>
    public sealed class LegacyXlsDrawingShape {
        private const uint KnownFlagMask = 0x00000FFF;

        /// <summary>
        /// Creates preserve-only metadata for an OfficeArt shape instance.
        /// </summary>
        public LegacyXlsDrawingShape(ushort shapeType, uint shapeId, uint flags) {
            ShapeType = shapeType;
            ShapeTypeName = GetShapeTypeName(shapeType);
            ShapeId = shapeId;
            Flags = flags;
            FlagNames = GetFlagNames(flags);
        }

        /// <summary>Gets the MSOSPT shape type stored in the OfficeArt record instance field.</summary>
        public ushort ShapeType { get; }

        /// <summary>Gets a stable display name for the shape type.</summary>
        public string ShapeTypeName { get; }

        /// <summary>Gets the OfficeArt shape identifier.</summary>
        public uint ShapeId { get; }

        /// <summary>Gets the raw OfficeArtFSP shape flags.</summary>
        public uint Flags { get; }

        /// <summary>Gets stable names for the defined shape flags set on this record.</summary>
        public IReadOnlyList<string> FlagNames { get; }

        /// <summary>Gets OfficeArtFSP shape flag bits that are not decoded as named flags.</summary>
        public uint ReservedFlags => Flags & ~KnownFlagMask;

        /// <summary>Gets whether the shape flag field contains bits outside the decoded flag set.</summary>
        public bool HasReservedFlags => ReservedFlags != 0;

        /// <summary>Gets a compact state key describing whether the shape flag field has reserved bits.</summary>
        public string ReservedState => HasReservedFlags ? $"Reserved:0x{ReservedFlags:X8}" : "ReservedClear";

        private static string GetShapeTypeName(ushort shapeType) {
            return shapeType switch {
                0x0000 => "NotPrimitive",
                0x0001 => "Rectangle",
                0x0002 => "RoundRectangle",
                0x0003 => "Ellipse",
                0x0014 => "Line",
                0x004B => "PictureFrame",
                0x00C9 => "HostControl",
                0x00CA => "TextBox",
                _ => $"ShapeType:0x{shapeType:X4}"
            };
        }

        private static IReadOnlyList<string> GetFlagNames(uint flags) {
            var names = new List<string>();
            if ((flags & 0x00000001) != 0) names.Add("Group");
            if ((flags & 0x00000002) != 0) names.Add("Child");
            if ((flags & 0x00000004) != 0) names.Add("Patriarch");
            if ((flags & 0x00000008) != 0) names.Add("Deleted");
            if ((flags & 0x00000010) != 0) names.Add("OleShape");
            if ((flags & 0x00000020) != 0) names.Add("HaveMaster");
            if ((flags & 0x00000040) != 0) names.Add("FlipH");
            if ((flags & 0x00000080) != 0) names.Add("FlipV");
            if ((flags & 0x00000100) != 0) names.Add("Connector");
            if ((flags & 0x00000200) != 0) names.Add("HaveAnchor");
            if ((flags & 0x00000400) != 0) names.Add("Background");
            if ((flags & 0x00000800) != 0) names.Add("HaveShapeType");
            return names;
        }
    }
}

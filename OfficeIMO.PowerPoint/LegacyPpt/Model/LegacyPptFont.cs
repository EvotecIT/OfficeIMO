namespace OfficeIMO.PowerPoint.LegacyPpt.Model {
    /// <summary>Represents one font entry from a binary PowerPoint document font collection.</summary>
    public sealed class LegacyPptFont {
        internal LegacyPptFont(ushort index, string typeface, byte characterSet,
            bool isEmbeddedSubset, bool isRaster, bool isDevice, bool isTrueType,
            bool disableSubstitution, byte pitchAndFamily, bool hasEmbeddedData) {
            Index = index;
            Typeface = typeface ?? throw new ArgumentNullException(nameof(typeface));
            CharacterSet = characterSet;
            IsEmbeddedSubset = isEmbeddedSubset;
            IsRaster = isRaster;
            IsDevice = isDevice;
            IsTrueType = isTrueType;
            DisableSubstitution = disableSubstitution;
            PitchAndFamily = pitchAndFamily;
            HasEmbeddedData = hasEmbeddedData;
        }

        /// <summary>Gets the zero-based font index referenced by text formatting records.</summary>
        public ushort Index { get; }

        /// <summary>Gets the UTF-16 typeface name.</summary>
        public string Typeface { get; }

        /// <summary>Gets the Windows LOGFONT character-set value.</summary>
        public byte CharacterSet { get; }

        /// <summary>Gets whether an embedded font is subsetted.</summary>
        public bool IsEmbeddedSubset { get; }

        /// <summary>Gets whether the font entry identifies a raster font.</summary>
        public bool IsRaster { get; }

        /// <summary>Gets whether the font entry identifies a device font.</summary>
        public bool IsDevice { get; }

        /// <summary>Gets whether the font entry identifies a TrueType font.</summary>
        public bool IsTrueType { get; }

        /// <summary>Gets whether PowerPoint font substitution is disabled for this entry.</summary>
        public bool DisableSubstitution { get; }

        /// <summary>Gets the Windows LOGFONT pitch-and-family byte.</summary>
        public byte PitchAndFamily { get; }

        /// <summary>Gets whether one or more embedded font data records follow this entry.</summary>
        public bool HasEmbeddedData { get; }
    }
}

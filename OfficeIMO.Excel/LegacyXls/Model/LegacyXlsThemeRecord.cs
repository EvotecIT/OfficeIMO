namespace OfficeIMO.Excel.LegacyXls.Model {
    /// <summary>
    /// Describes a BIFF Theme record discovered during legacy XLS import.
    /// </summary>
    public sealed class LegacyXlsThemeRecord {
        private readonly byte[] _themeBytes;

        /// <summary>
        /// Creates theme preservation metadata.
        /// </summary>
        public LegacyXlsThemeRecord(
            int recordOffset,
            ushort recordType,
            int payloadLength,
            uint themeVersion,
            string themeVersionName,
            byte[]? themeBytes) {
            if (payloadLength < 0) {
                throw new ArgumentOutOfRangeException(nameof(payloadLength));
            }

            RecordOffset = recordOffset;
            RecordType = recordType;
            PayloadLength = payloadLength;
            ThemeVersion = themeVersion;
            ThemeVersionName = string.IsNullOrWhiteSpace(themeVersionName)
                ? "Unknown"
                : themeVersionName;
            _themeBytes = themeBytes == null || themeBytes.Length == 0
                ? Array.Empty<byte>()
                : (byte[])themeBytes.Clone();
        }

        /// <summary>Gets the byte offset of the BIFF Theme record.</summary>
        public int RecordOffset { get; }

        /// <summary>Gets the BIFF record type identifier.</summary>
        public ushort RecordType { get; }

        /// <summary>Gets the BIFF record payload length in bytes.</summary>
        public int PayloadLength { get; }

        /// <summary>Gets the raw BIFF theme version value.</summary>
        public uint ThemeVersion { get; }

        /// <summary>Gets the decoded theme version name.</summary>
        public string ThemeVersionName { get; }

        /// <summary>Gets whether the Theme record carries embedded theme content bytes.</summary>
        public bool HasThemeBytes => _themeBytes.Length > 0;

        /// <summary>Gets whether the record declares the built-in default theme without embedded theme package bytes.</summary>
        public bool IsDefaultThemeMarker => !HasThemeBytes && ThemeVersion == 124226U;

        /// <summary>Gets the number of embedded theme content bytes preserved from the Theme record.</summary>
        public int ThemeByteCount => _themeBytes.Length;

        /// <summary>Gets a copy of the embedded theme content bytes preserved from the Theme record.</summary>
        public byte[] ThemeBytes => (byte[])_themeBytes.Clone();
    }
}

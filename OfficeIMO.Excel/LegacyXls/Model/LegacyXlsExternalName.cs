namespace OfficeIMO.Excel.LegacyXls.Model {
    /// <summary>
    /// Describes a name declared by an external supporting link in a legacy XLS workbook.
    /// </summary>
    public sealed class LegacyXlsExternalName {
        internal LegacyXlsExternalName(
            string name,
            int? localSheetIndex,
            bool builtIn,
            bool wantsAdvise,
            bool wantsPicture,
            bool ole,
            bool oleLink,
            int cachedClipboardFormat,
            bool icon,
            LegacyXlsExternalNameBodyKind bodyKind) {
            Name = name;
            LocalSheetIndex = localSheetIndex;
            BuiltIn = builtIn;
            WantsAdvise = wantsAdvise;
            WantsPicture = wantsPicture;
            Ole = ole;
            OleLink = oleLink;
            CachedClipboardFormat = cachedClipboardFormat;
            Icon = icon;
            BodyKind = bodyKind;
        }

        /// <summary>
        /// Gets the external defined-name text.
        /// </summary>
        public string Name { get; }

        /// <summary>
        /// Gets the zero-based external sheet scope when the external name is sheet-local.
        /// </summary>
        public int? LocalSheetIndex { get; }

        /// <summary>
        /// Gets a value indicating whether this is a built-in external name.
        /// </summary>
        public bool BuiltIn { get; }

        /// <summary>
        /// Gets a value indicating whether the link requests automatic DDE or OLE advise updates.
        /// </summary>
        public bool WantsAdvise { get; }

        /// <summary>
        /// Gets a value indicating whether linked DDE or OLE data uses a picture format.
        /// </summary>
        public bool WantsPicture { get; }

        /// <summary>
        /// Gets the raw ExternName fOle flag.
        /// </summary>
        public bool Ole { get; }

        /// <summary>
        /// Gets the raw ExternName fOleLink flag.
        /// </summary>
        public bool OleLink { get; }

        /// <summary>
        /// Gets the signed ExternName cached clipboard format code.
        /// </summary>
        public int CachedClipboardFormat { get; }

        /// <summary>
        /// Gets a friendly cached clipboard format name.
        /// </summary>
        public string CachedClipboardFormatName => GetCachedClipboardFormatName(CachedClipboardFormat);

        /// <summary>
        /// Gets a value indicating whether linked data is displayed as an icon.
        /// </summary>
        public bool Icon { get; }

        /// <summary>
        /// Gets the decoded ExternName body kind based on its supporting link and flags.
        /// </summary>
        public LegacyXlsExternalNameBodyKind BodyKind { get; }

        private static string GetCachedClipboardFormatName(int format) {
            return format switch {
                -1 => "None",
                0 => "TextOrExternalName",
                2 => "EnhancedMetafile",
                5 => "Csv",
                6 => "Sylk",
                7 => "Rtf",
                8 => "Biff8",
                9 => "Bitmap",
                16 => "ApplicationTable",
                20 => "Biff3",
                30 => "Biff4",
                36 => "MetafilePicture",
                44 => "UnicodeText",
                63 => "Biff12",
                _ => $"Unknown:{format}"
            };
        }
    }
}

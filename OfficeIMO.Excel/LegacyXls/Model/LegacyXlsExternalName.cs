namespace OfficeIMO.Excel.LegacyXls.Model {
    /// <summary>
    /// Describes a name declared by an external supporting link in a legacy XLS workbook.
    /// </summary>
    public sealed class LegacyXlsExternalName {
        internal LegacyXlsExternalName(string name, int? localSheetIndex, bool builtIn) {
            Name = name;
            LocalSheetIndex = localSheetIndex;
            BuiltIn = builtIn;
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
    }
}

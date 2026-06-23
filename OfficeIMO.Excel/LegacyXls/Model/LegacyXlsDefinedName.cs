namespace OfficeIMO.Excel.LegacyXls.Model {
    /// <summary>
    /// Represents a BIFF defined name whose formula can be projected to an A1 reference.
    /// </summary>
    public sealed class LegacyXlsDefinedName {
        internal LegacyXlsDefinedName(string name, string reference, int? localSheetIndex, bool hidden, bool builtIn) {
            Name = name;
            Reference = reference;
            LocalSheetIndex = localSheetIndex;
            Hidden = hidden;
            BuiltIn = builtIn;
        }

        /// <summary>Gets the Open XML defined-name identifier.</summary>
        public string Name { get; }

        /// <summary>Gets the sheet-qualified A1 reference represented by the BIFF name formula.</summary>
        public string Reference { get; }

        /// <summary>Gets the zero-based local sheet index when the name is sheet-scoped.</summary>
        public int? LocalSheetIndex { get; }

        /// <summary>Gets whether the defined name is hidden.</summary>
        public bool Hidden { get; }

        /// <summary>Gets whether the defined name came from a BIFF built-in name.</summary>
        public bool BuiltIn { get; }
    }
}

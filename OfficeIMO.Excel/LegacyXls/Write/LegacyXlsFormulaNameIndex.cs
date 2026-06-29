namespace OfficeIMO.Excel.LegacyXls.Write {
    internal sealed class LegacyXlsFormulaNameIndex {
        private readonly Dictionary<string, uint> _globalNames;
        private readonly Dictionary<string, uint> _localNames;
        private readonly LegacyXlsExternSheetTable _externSheetTable;

        internal LegacyXlsFormulaNameIndex(
            Dictionary<string, uint> globalNames,
            Dictionary<string, uint> localNames)
            : this(globalNames, localNames, LegacyXlsExternSheetTable.Empty) {
        }

        internal LegacyXlsFormulaNameIndex(
            Dictionary<string, uint> globalNames,
            Dictionary<string, uint> localNames,
            LegacyXlsExternSheetTable externSheetTable) {
            _globalNames = globalNames ?? throw new ArgumentNullException(nameof(globalNames));
            _localNames = localNames ?? throw new ArgumentNullException(nameof(localNames));
            _externSheetTable = externSheetTable ?? throw new ArgumentNullException(nameof(externSheetTable));
        }

        internal static LegacyXlsFormulaNameIndex Empty { get; } = new(
            new Dictionary<string, uint>(StringComparer.OrdinalIgnoreCase),
            new Dictionary<string, uint>(StringComparer.OrdinalIgnoreCase),
            LegacyXlsExternSheetTable.Empty);

        internal bool TryGetNameIndex(string name, int sheetIndex, out uint oneBasedNameIndex) {
            oneBasedNameIndex = 0;
            if (string.IsNullOrWhiteSpace(name)) {
                return false;
            }

            if (sheetIndex >= 0 && _localNames.TryGetValue(CreateLocalKey(sheetIndex, name), out oneBasedNameIndex)) {
                return true;
            }

            return _globalNames.TryGetValue(name, out oneBasedNameIndex);
        }

        internal bool TryGetExternSheetIndex(string sheetName, out ushort externSheetIndex) {
            return _externSheetTable.TryGetSheetIndex(sheetName, out externSheetIndex);
        }

        internal bool TryGetExternSheetRangeIndex(string firstSheetName, string lastSheetName, out ushort externSheetIndex) {
            return _externSheetTable.TryGetSheetRangeIndex(firstSheetName, lastSheetName, out externSheetIndex);
        }

        internal bool TryGetExternalNameIndex(string target, string name, out ushort externSheetIndex, out uint oneBasedNameIndex) {
            return _externSheetTable.TryGetExternalNameIndex(target, sheetName: null, name, out externSheetIndex, out oneBasedNameIndex);
        }

        internal bool TryGetExternalNameIndex(string target, string? sheetName, string name, out ushort externSheetIndex, out uint oneBasedNameIndex) {
            return _externSheetTable.TryGetExternalNameIndex(target, sheetName, name, out externSheetIndex, out oneBasedNameIndex);
        }

        internal static string CreateLocalKey(int sheetIndex, string name) {
            return sheetIndex.ToString(System.Globalization.CultureInfo.InvariantCulture) + "\0" + name;
        }
    }
}

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel {
    public partial class ExcelDocument {
        private DefinedNameIndex? _definedNameIndex;

        /// <summary>
        /// Finds the defined names a new definition replaces: same name, compared case-insensitively, and same scope.
        /// Names set one after another look them up in an index instead of walking every definition.
        /// </summary>
        private List<DefinedName> FindReplacedDefinedNames(DefinedNames definedNames, string name, uint? sheetIndex) {
            if (_openXmlDocumentExposed) {
                return definedNames.Elements<DefinedName>().Where(definition => IsSameDefinedName(definition, name, sheetIndex)).ToList();
            }

            DefinedNameIndex? index = _definedNameIndex;
            if (index == null || !index.IsCurrent(definedNames)) {
                index = new DefinedNameIndex(definedNames);
                _definedNameIndex = index;
            }
            return index.Find(name, sheetIndex);
        }

        /// <summary>Records a definition just appended after <paramref name="previousLastChild"/>, so the index stays current.</summary>
        private void RecordAppendedDefinedName(DefinedNames definedNames, OpenXmlElement? previousLastChild, DefinedName definition) {
            DefinedNameIndex? index = _definedNameIndex;
            if (index != null && index.List == definedNames && index.LastChild == previousLastChild) {
                index.Appended(definition);
            }
        }

        /// <summary>Keeps the index current after definitions it returned were removed; it drops them on the next lookup.</summary>
        private void RecordRemovedDefinedNames(DefinedNames definedNames) {
            DefinedNameIndex? index = _definedNameIndex;
            if (index != null && index.List == definedNames) index.Removed();
        }

        private static bool IsSameDefinedName(DefinedName definition, string name, uint? sheetIndex)
            => string.Equals(definition.Name?.Value, name, StringComparison.OrdinalIgnoreCase)
                && definition.LocalSheetId?.Value == sheetIndex;

        /// <summary>
        /// The definitions of a definedNames list by name. It serves while the list and its last child are the ones it
        /// was built or last appended with; workbook code adds definitions only by appending, and every hit is checked
        /// against the live element, so definitions removed, renamed or rescoped since drop out.
        /// </summary>
        private sealed class DefinedNameIndex {
            private readonly Dictionary<string, List<DefinedName>> _byName = new(StringComparer.OrdinalIgnoreCase);

            internal DefinedNameIndex(DefinedNames list) {
                List = list;
                foreach (DefinedName definition in list.Elements<DefinedName>()) Add(definition);
                LastChild = list.LastChild;
            }

            internal DefinedNames List { get; }
            internal OpenXmlElement? LastChild { get; private set; }

            internal bool IsCurrent(DefinedNames list) => ReferenceEquals(List, list) && ReferenceEquals(LastChild, list.LastChild);

            internal void Removed() => LastChild = List.LastChild;

            internal void Appended(DefinedName definition) {
                Add(definition);
                LastChild = definition;
            }

            internal List<DefinedName> Find(string name, uint? sheetIndex) {
                var found = new List<DefinedName>();
                if (!_byName.TryGetValue(name, out List<DefinedName>? definitions)) return found;
                definitions.RemoveAll(definition => !ReferenceEquals(definition.Parent, List)
                    || !string.Equals(definition.Name?.Value, name, StringComparison.OrdinalIgnoreCase));
                foreach (DefinedName definition in definitions) {
                    if (IsSameDefinedName(definition, name, sheetIndex)) found.Add(definition);
                }
                return found;
            }

            private void Add(DefinedName definition) {
                string key = definition.Name?.Value ?? string.Empty;
                if (!_byName.TryGetValue(key, out List<DefinedName>? definitions)) {
                    definitions = new List<DefinedName>(1);
                    _byName.Add(key, definitions);
                }
                definitions.Add(definition);
            }
        }
    }
}

using System;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        public void CreateNamedRange(string name, string a1Range, bool workbookScope = false) {
            _excelDocument.CreateNamedRange(name, this, a1Range, workbookScope);
        }

        public string? GetNamedRange(string name, bool workbookScope = false) {
            var reference = workbookScope ? _excelDocument.GetNamedRange(name) : _excelDocument.GetNamedRange(name, this);
            if (reference != null && !workbookScope) {
                var idx = reference.IndexOf('!');
                if (idx >= 0 && idx < reference.Length - 1) {
                    return reference.Substring(idx + 1);
                }
            }
            return reference;
        }

        public bool DeleteNamedRange(string name, bool workbookScope = false) {
            return workbookScope ? _excelDocument.DeleteNamedRange(name) : _excelDocument.DeleteNamedRange(name, this);
        }
    }
}

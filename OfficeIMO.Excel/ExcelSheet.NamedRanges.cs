namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        public void SetNamedRange(string name, string range, bool save = true) {
            _excelDocument.SetNamedRange(name, range, this, save);
        }

        public string? GetNamedRange(string name) {
            return _excelDocument.GetNamedRange(name, this);
        }

        public System.Collections.Generic.IReadOnlyDictionary<string, string> GetAllNamedRanges() {
            return _excelDocument.GetAllNamedRanges(this);
        }

        public bool RemoveNamedRange(string name, bool save = true) {
            return _excelDocument.RemoveNamedRange(name, this, save);
        }
    }
}


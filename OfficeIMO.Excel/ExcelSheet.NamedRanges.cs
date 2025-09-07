namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        public void SetNamedRange(string name, string range) {
            _excelDocument.SetNamedRange(name, range, this);
        }

        public string? GetNamedRange(string name) {
            return _excelDocument.GetNamedRange(name, this);
        }

        public bool RemoveNamedRange(string name) {
            return _excelDocument.RemoveNamedRange(name, this);
        }
    }
}


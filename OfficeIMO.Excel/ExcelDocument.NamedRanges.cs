using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel {
    public partial class ExcelDocument {
        private static readonly Regex _a1Regex = new Regex("^\\$?[A-Za-z]{1,3}\\$?[1-9][0-9]*(?:(:\\$?[A-Za-z]{1,3}\\$?[1-9][0-9]*))?$", RegexOptions.Compiled);

        private static void ValidateA1(string range) {
            if (string.IsNullOrWhiteSpace(range) || !_a1Regex.IsMatch(range)) {
                throw new ArgumentException("Range must be a valid A1 reference.", nameof(range));
            }
        }

        private static string EscapeSheetName(string sheetName) {
            return sheetName.Contains(' ') ? $"'{sheetName}'" : sheetName;
        }

        public void CreateNamedRange(string name, ExcelSheet sheet, string a1Range, bool workbookScope = true) {
            ValidateA1(a1Range);
            Locking.ExecuteWrite(EnsureLock(), () => {
                var workbookPart = _spreadSheetDocument.WorkbookPart ?? throw new InvalidOperationException("WorkbookPart is null");
                var definedNames = workbookPart.Workbook.DefinedNames ?? workbookPart.Workbook.AppendChild(new DefinedNames());
                var definedName = new DefinedName { Name = name };
                if (!workbookScope) {
                    var sheets = workbookPart.Workbook.Sheets!.OfType<Sheet>().ToList();
                    var index = sheets.FindIndex(s => s.Name != null && s.Name.Value == sheet.Name);
                    if (index >= 0) {
                        definedName.LocalSheetId = (uint)index;
                    }
                }
                definedName.Text = $"{EscapeSheetName(sheet.Name)}!{a1Range}";
                definedNames.AppendChild(definedName);
            });
        }

        public string? GetNamedRange(string name, ExcelSheet? sheet = null) {
            return Locking.ExecuteRead(EnsureLock(), () => {
                var workbookPart = _spreadSheetDocument.WorkbookPart ?? throw new InvalidOperationException("WorkbookPart is null");
                var definedNames = workbookPart.Workbook.DefinedNames;
                if (definedNames == null) return null;
                IEnumerable<DefinedName> names = definedNames.Elements<DefinedName>().Where(d => d.Name?.Value == name);
                if (sheet != null) {
                    var sheets = workbookPart.Workbook.Sheets!.OfType<Sheet>().ToList();
                    var index = sheets.FindIndex(s => s.Name != null && s.Name.Value == sheet.Name);
                    names = names.Where(d => d.LocalSheetId != null && d.LocalSheetId.Value == (uint)index);
                } else {
                    names = names.Where(d => d.LocalSheetId == null);
                }
                return names.FirstOrDefault()?.Text;
            });
        }

        public bool DeleteNamedRange(string name, ExcelSheet? sheet = null) {
            return Locking.ExecuteWrite(EnsureLock(), () => {
                var workbookPart = _spreadSheetDocument.WorkbookPart ?? throw new InvalidOperationException("WorkbookPart is null");
                var definedNames = workbookPart.Workbook.DefinedNames;
                if (definedNames == null) return false;
                IEnumerable<DefinedName> names = definedNames.Elements<DefinedName>().Where(d => d.Name?.Value == name);
                if (sheet != null) {
                    var sheets = workbookPart.Workbook.Sheets!.OfType<Sheet>().ToList();
                    var index = sheets.FindIndex(s => s.Name != null && s.Name.Value == sheet.Name);
                    names = names.Where(d => d.LocalSheetId != null && d.LocalSheetId.Value == (uint)index);
                } else {
                    names = names.Where(d => d.LocalSheetId == null);
                }
                var target = names.FirstOrDefault();
                if (target == null) {
                    return false;
                }
                target.Remove();
                if (!definedNames.Any()) {
                    definedNames.Remove();
                }
                return true;
            });
        }
    }
}

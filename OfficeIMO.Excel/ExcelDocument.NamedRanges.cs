using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel {
    public partial class ExcelDocument {
        private static readonly Regex _a1Regex = new Regex("^\\$?[A-Za-z]{1,3}\\$?[1-9][0-9]*(?:(:\\$?[A-Za-z]{1,3}\\$?[1-9][0-9]*))?$", RegexOptions.Compiled | RegexOptions.CultureInvariant);

        private static void ValidateA1(string range) {
            if (string.IsNullOrWhiteSpace(range) || !_a1Regex.IsMatch(range)) {
                throw new ArgumentException("Range must be a valid A1 reference.", nameof(range));
            }
        }

        private static string EscapeSheetName(string sheetName) {
            var escaped = sheetName.Replace("'", "''");
            return escaped.IndexOfAny(new[] { ' ', '!', '\'', '[', ']' }) >= 0 ? $"'{escaped}'" : escaped;
        }

        private static int GetSheetIndex(WorkbookPart workbookPart, string sheetName) {
            var index = 0;
            foreach (var sheet in workbookPart.Workbook.Sheets!.Elements<Sheet>()) {
                if (sheet.Name != null && sheet.Name.Value == sheetName) {
                    return index;
                }
                index++;
            }
            return -1;
        }

        public void CreateNamedRange(string name, ExcelSheet sheet, string a1Range, bool workbookScope = true) {
            ValidateA1(a1Range);
            Locking.ExecuteWrite(EnsureLock(), () => {
                var workbookPart = _spreadSheetDocument.WorkbookPart ?? throw new InvalidOperationException("WorkbookPart is null");
                var definedNames = workbookPart.Workbook.DefinedNames ?? workbookPart.Workbook.AppendChild(new DefinedNames());
                var index = workbookScope ? -1 : GetSheetIndex(workbookPart, sheet.Name);
                if (!workbookScope && index < 0) {
                    throw new ArgumentException($"Worksheet '{sheet.Name}' not found.", nameof(sheet));
                }
                var localSheetId = workbookScope ? (uint?)null : (uint)index;

                var existing = definedNames.Elements<DefinedName>().FirstOrDefault(d => d.Name?.Value == name && d.LocalSheetId == localSheetId);
                if (existing != null) {
                    throw new InvalidOperationException($"Named range '{name}' already exists.");
                }

                var definedName = new DefinedName { Name = name, Text = $"{EscapeSheetName(sheet.Name)}!{a1Range}" };
                if (localSheetId != null) {
                    definedName.LocalSheetId = localSheetId.Value;
                }
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
                    var index = GetSheetIndex(workbookPart, sheet.Name);
                    if (index < 0) {
                        return null;
                    }
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
                    var index = GetSheetIndex(workbookPart, sheet.Name);
                    if (index < 0) {
                        return false;
                    }
                    names = names.Where(d => d.LocalSheetId != null && d.LocalSheetId.Value == (uint)index);
                } else {
                    names = names.Where(d => d.LocalSheetId == null);
                }
                var target = names.FirstOrDefault();
                if (target == null) {
                    return false;
                }
                target.Remove();
                if (!definedNames.Elements<DefinedName>().Any()) {
                    definedNames.Remove();
                }
                return true;
            });
        }
    }
}

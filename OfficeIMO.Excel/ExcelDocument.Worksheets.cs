using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Validation;
using OfficeIMO.Excel.Utilities;
using OfficeIMO.Shared;
using System.IO.Packaging;
using System.Threading;
using System.Threading.Tasks;
using System.Xml;
using System;
using System.Diagnostics;
using System.IO;

namespace OfficeIMO.Excel {
    public partial class ExcelDocument : IDisposable, IAsyncDisposable {

        /// <summary>
        /// Creates a new Excel document with a single worksheet.
        /// </summary>
        /// <param name="filePath">Path to the new file.</param>
        /// <param name="worksheetName">Name of the worksheet.</param>
        /// <returns>Created <see cref="ExcelDocument"/> instance.</returns>
        public static ExcelDocument Create(string filePath, string worksheetName) {
            ExcelDocument excelDocument = Create(filePath);
            // Prefer a sanitized sheet name for convenience in the common Create(path, name) flow
            excelDocument.AddWorksheet(worksheetName, SheetNameValidationMode.Sanitize);
            return excelDocument;
        }

        /// <summary>
        /// Adds a worksheet to the document.
        /// </summary>
        /// <param name="worksheetName">Worksheet name.</param>
        /// <returns>Created <see cref="ExcelSheet"/> instance.</returns>
        public ExcelSheet AddWorksheet(string worksheetName = "") {
            return AddWorksheet(worksheetName, SheetNameValidationMode.Sanitize);
        }

        /// <summary>
        /// Adds a worksheet to the document with control over name validation.
        /// </summary>
        /// <param name="worksheetName">Requested worksheet name.</param>
        /// <param name="validationMode">How to validate the sheet name: None (no checks), Sanitize (coerce), or Strict (throw on invalid).</param>
        /// <returns>Created <see cref="ExcelSheet"/> instance.</returns>
        public ExcelSheet AddWorksheet(string worksheetName, SheetNameValidationMode validationMode) {
            if (!_materializingDeferredDataSetImport) {
                MaterializeDeferredDataSetImport();
            }

            return Locking.ExecuteWrite(EnsureLock(), () => {
                EnsureSheetCacheInitialized(_lock);
                string name = ValidateOrSanitizeSheetName(worksheetName, validationMode, currentSheetName: null);
                ExcelSheet excelSheet = new ExcelSheet(this, _workBookPart, _spreadSheetDocument, name);
                MarkSheetCacheDirty();
                MarkRequiresSavePreflight();
                return excelSheet;
            });
        }

        internal void RenameWorksheet(ExcelSheet sheet, string worksheetName, SheetNameValidationMode validationMode) {
            if (sheet == null) throw new ArgumentNullException(nameof(sheet));

            Locking.ExecuteWrite(EnsureLock(), () => {
                string currentName = sheet.Name;
                string validatedName = ValidateOrSanitizeSheetName(worksheetName, validationMode, currentName);
                if (string.Equals(currentName, validatedName, StringComparison.Ordinal)) {
                    return;
                }

                var target = WorkbookRoot.Sheets?
                    .OfType<DocumentFormat.OpenXml.Spreadsheet.Sheet>()
                    .FirstOrDefault(s => ReferenceEquals(s, sheet.SheetElement)
                                         || string.Equals(s.Name?.Value, currentName, StringComparison.Ordinal));
                if (target == null) {
                    throw new ArgumentException("Worksheet not found in workbook.", nameof(sheet));
                }

                target.Name = validatedName;
                UpdateSheetNameReferences(currentName, validatedName);
                WorkbookRoot.Save();
                MarkRequiresSavePreflight();
            });
        }

        private string ValidateOrSanitizeSheetName(string name, SheetNameValidationMode mode, string? currentSheetName) {
            // Collect existing names (case-insensitive)
            var existing = new System.Collections.Generic.HashSet<string>(System.StringComparer.OrdinalIgnoreCase);
            foreach (var s in WorkbookRoot.Sheets?.OfType<DocumentFormat.OpenXml.Spreadsheet.Sheet>() ?? System.Linq.Enumerable.Empty<DocumentFormat.OpenXml.Spreadsheet.Sheet>()) {
                var existingName = s.Name?.Value;
                if (string.IsNullOrEmpty(existingName)) continue;
                if (!string.IsNullOrEmpty(currentSheetName) && string.Equals(existingName, currentSheetName, StringComparison.OrdinalIgnoreCase)) continue;
                existing.Add(existingName!);
            }

            if (mode == SheetNameValidationMode.None) {
                // Preserve historical behavior: default to "Sheet1" when empty
                if (string.IsNullOrEmpty(name)) name = "Sheet1";
                return name;
            }

            // Rules common to Sanitize/Strict
            static bool ContainsInvalidChars(string s) {
                foreach (char c in s) {
                    if (c == ':' || c == '\\' || c == '/' || c == '?' || c == '*' || c == '[' || c == ']') return true;
                }
                return false;
            }

            string baseName = name ?? string.Empty;
            baseName = baseName.Trim();
            baseName = baseName.Trim('\'', ' ');

            if (mode == SheetNameValidationMode.Strict) {
                if (string.IsNullOrEmpty(baseName)) throw new System.ArgumentException("Worksheet name cannot be empty.", nameof(name));
                if (baseName.Length > 31) throw new System.ArgumentException("Worksheet name cannot exceed 31 characters.", nameof(name));
                if (ContainsInvalidChars(baseName)) throw new System.ArgumentException("Worksheet name contains invalid characters (: \\ / ? * [ ]).", nameof(name));
                if (existing.Contains(baseName)) throw new System.ArgumentException($"Worksheet name '{baseName}' already exists.", nameof(name));
                return baseName;
            }

            // Sanitize
            var sb = new System.Text.StringBuilder(baseName.Length > 0 ? baseName.Length : 5);
            foreach (char c in baseName) {
                if (c == ':' || c == '\\' || c == '/' || c == '?' || c == '*' || c == '[' || c == ']') sb.Append('_');
                else sb.Append(c);
            }
            string cleaned = sb.ToString().Trim();
            // Collapse multiple underscores and trim leading/trailing underscores for nicer names
            cleaned = _multipleUnderscoresRegex.Replace(cleaned, "_");
            cleaned = cleaned.Trim('_');
            if (cleaned.Length == 0) cleaned = GenerateDefaultSheetName(existing);
            if (cleaned.Length > 31) cleaned = cleaned.Substring(0, 31);

            // Ensure uniqueness by appending (2), (3), ...
            string candidate = cleaned;
            int n = 2;
            while (existing.Contains(candidate)) {
                string suffix = " (" + n.ToString(System.Globalization.CultureInfo.InvariantCulture) + ")";
                int maxBase = 31 - suffix.Length;
                string basePart = cleaned.Length > maxBase ? cleaned.Substring(0, maxBase) : cleaned;
                candidate = basePart + suffix;
                n++;
            }
            return candidate;
        }

        private static string GenerateDefaultSheetName(System.Collections.Generic.ISet<string> existing) {
            int n = 1;
            string candidate = "Sheet1";
            while (existing.Contains(candidate)) {
                n++;
                candidate = "Sheet" + n.ToString(System.Globalization.CultureInfo.InvariantCulture);
            }

            return candidate;
        }
    }
}

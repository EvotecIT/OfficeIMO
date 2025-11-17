using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Xml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
#if DEBUG
using DocumentFormat.OpenXml.Validation;
#endif

namespace OfficeIMO.Excel {
    internal static class WorksheetIntegrityValidator {
        internal static void Validate(WorksheetPart worksheetPart, ExecutionPolicy policy, string sheetName) {
            if (worksheetPart == null) throw new ArgumentNullException(nameof(worksheetPart));
            if (policy == null) throw new ArgumentNullException(nameof(policy));

            var mode = policy.WorksheetValidation;
            if (mode == WorksheetValidationMode.Disabled) {
                return;
            }

            if (mode == WorksheetValidationMode.DiagnosticsOnly && !policy.AreDiagnosticsRequested) {
                return;
            }

            try {
                EnsureInvariants(worksheetPart, sheetName);

#if DEBUG
                if (policy.UseOpenXmlValidatorInDebug &&
                    (mode == WorksheetValidationMode.Always
                        || mode == WorksheetValidationMode.DebugOnly
                        || (mode == WorksheetValidationMode.DiagnosticsOnly && policy.AreDiagnosticsRequested))) {
                    var validator = new OpenXmlValidator();
                    var firstError = validator.Validate(worksheetPart.Worksheet).FirstOrDefault();
                    if (firstError != null) {
                        throw new InvalidOperationException($"OpenXmlValidator failure in worksheet '{sheetName}': {firstError.Description}");
                    }
                }
#endif
            } catch (InvalidOperationException) {
                throw;
            } catch (Exception ex) {
                throw new InvalidOperationException($"Worksheet '{sheetName}' failed structural validation: {ex.Message}", ex);
            }
        }

        internal static TimeSpan MeasureTargetedValidation(WorksheetPart worksheetPart, int iterations, string sheetName) {
            if (iterations <= 0) throw new ArgumentOutOfRangeException(nameof(iterations));

            var sw = Stopwatch.StartNew();
            for (int i = 0; i < iterations; i++) {
                EnsureInvariants(worksheetPart, sheetName);
            }
            sw.Stop();
            return sw.Elapsed;
        }

        internal static TimeSpan MeasureLegacyOuterXml(WorksheetPart worksheetPart, int iterations) {
            if (iterations <= 0) throw new ArgumentOutOfRangeException(nameof(iterations));

            var sw = Stopwatch.StartNew();
            for (int i = 0; i < iterations; i++) {
                using var sr = new StringReader(worksheetPart.Worksheet.OuterXml);
                using var reader = XmlReader.Create(sr, new XmlReaderSettings { IgnoreWhitespace = true });
                while (reader.Read()) { }
            }
            sw.Stop();
            return sw.Elapsed;
        }

        private static void EnsureInvariants(WorksheetPart worksheetPart, string sheetName) {
            var sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();
            if (sheetData == null) {
                return;
            }

            uint previousRowIndex = 0;
            foreach (var row in sheetData.Elements<Row>()) {
                if (row.RowIndex?.Value is not uint rowIndex || rowIndex == 0) {
                    throw new InvalidOperationException($"Worksheet '{sheetName}' contains a row without a valid RowIndex.");
                }

                if (rowIndex <= previousRowIndex) {
                    throw new InvalidOperationException($"Worksheet '{sheetName}' contains non-increasing row indices (row {rowIndex}).");
                }

                previousRowIndex = rowIndex;

                int lastColumnIndex = 0;
                foreach (var cell in row.Elements<Cell>()) {
                    var reference = cell.CellReference?.Value;
                    if (string.IsNullOrWhiteSpace(reference)) {
                        throw new InvalidOperationException($"Row {rowIndex} in worksheet '{sheetName}' contains a cell without a reference.");
                    }

                    var parsedRowIndex = GetRowIndex(reference!);
                    if (parsedRowIndex != rowIndex) {
                        throw new InvalidOperationException($"Cell '{reference}' in worksheet '{sheetName}' does not match its containing row {rowIndex}.");
                    }

                    int columnIndex = GetColumnIndex(reference!);
                    if (columnIndex <= lastColumnIndex) {
                        throw new InvalidOperationException($"Row {rowIndex} in worksheet '{sheetName}' has unsorted or duplicate cells at column index {columnIndex}.");
                    }

                    lastColumnIndex = columnIndex;
                }
            }
        }

        private static int GetColumnIndex(string cellReference) {
            int columnIndex = 0;
            foreach (char ch in cellReference) {
                if (!char.IsLetter(ch)) {
                    break;
                }

                columnIndex = (columnIndex * 26) + (char.ToUpperInvariant(ch) - 'A' + 1);
            }
            return columnIndex;
        }

        private static uint GetRowIndex(string cellReference) {
            int start = 0;
            while (start < cellReference.Length && char.IsLetter(cellReference[start])) {
                start++;
            }

            if (start >= cellReference.Length) {
                throw new InvalidOperationException($"Cell reference '{cellReference}' is missing a row component.");
            }

            uint rowIndex = 0;
            for (int i = start; i < cellReference.Length; i++) {
                char ch = cellReference[i];
                if (ch < '0' || ch > '9') {
                    throw new InvalidOperationException($"Cell reference '{cellReference}' has an invalid row component.");
                }

                uint digit = (uint)(ch - '0');
                if (rowIndex > (uint.MaxValue - digit) / 10) {
                    throw new InvalidOperationException($"Cell reference '{cellReference}' has an invalid row component.");
                }

                rowIndex = (rowIndex * 10) + digit;
            }

            if (rowIndex == 0) {
                throw new InvalidOperationException($"Cell reference '{cellReference}' has an invalid row component.");
            }

            return rowIndex;
        }
    }
}

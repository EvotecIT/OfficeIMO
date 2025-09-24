using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Globalization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel
{
    /// <summary>
    /// Planner for SharedStrings to avoid DOM mutations during parallel compute.
    /// Collects distinct strings concurrently, applies them under document lock,
    /// and fixes up prepared cell values to reference shared string indices.
    /// </summary>
    internal sealed class SharedStringPlanner
    {
        private readonly ConcurrentDictionary<string, byte> _distinct = new();
        private Dictionary<string, int>? _finalIndex;

        /// <summary>
        /// Records a string for shared string table inclusion and returns the sanitized version.
        /// </summary>
        /// <param name="s">The string to record (null becomes empty string).</param>
        /// <returns>The sanitized string that will be used for indexing.</returns>
        public string Note(string? s)
        {
            var normalized = NormalizeText(s);

            // Clamp and strip illegal XML control characters defensively
            var safe = Utilities.ExcelSanitizer.SanitizeString(normalized);
            _distinct.TryAdd(safe, 0);
            return safe;
        }

        /// <summary>
        /// Apply collected strings to the document's SharedStringTable and build final index mapping.
        /// Must be called inside a serialized apply stage (under document write lock).
        /// </summary>
        public void ApplyTo(ExcelDocument doc)
        {
            if (_distinct.IsEmpty)
            {
                _finalIndex = new Dictionary<string, int>(0);
                return;
            }

            var map = new Dictionary<string, int>(StringComparer.Ordinal);
            foreach (var s in _distinct.Keys)
            {
                int idx = doc.GetSharedStringIndex(s);
                map[s] = idx;
            }
            _finalIndex = map;
        }

        /// <summary>
        /// Fixes a prepared cell tuple in-place, replacing SharedString text with its index.
        /// </summary>
        public void Fixup(
            ref (
                int Row,
                int Col,
                CellValue Val,
                EnumValue<DocumentFormat.OpenXml.Spreadsheet.CellValues> Type) prepared)
        {
            if (_finalIndex is null) return;
            if (prepared.Type?.Value != DocumentFormat.OpenXml.Spreadsheet.CellValues.SharedString) return;

            // prepared.Val.Text currently holds the raw string; replace with index text
            var text = NormalizeText(prepared.Val?.Text);
            var sanitized = Utilities.ExcelSanitizer.SanitizeString(text);
            if (_finalIndex.TryGetValue(sanitized, out int idx))
            {
                prepared.Val = new CellValue(idx.ToString(CultureInfo.InvariantCulture));
            }
            else
            {
                // Fallback: if not found (shouldn't happen), keep as string cell
                prepared.Type = DocumentFormat.OpenXml.Spreadsheet.CellValues.String;
            }
        }

        /// <summary>
        /// Applies planner to document and fixes all prepared cells.
        /// Must be called inside serialized apply stage.
        /// </summary>
        public void ApplyAndFixup(
            (
                int Row,
                int Col,
                CellValue Val,
                EnumValue<DocumentFormat.OpenXml.Spreadsheet.CellValues> Type)[] prepared,
            ExcelDocument doc)
        {
            ApplyTo(doc);
            for (int i = 0; i < prepared.Length; i++)
            {
                Fixup(ref prepared[i]);
            }
        }

        private static string NormalizeText(string? text) => text ?? string.Empty;
    }
}

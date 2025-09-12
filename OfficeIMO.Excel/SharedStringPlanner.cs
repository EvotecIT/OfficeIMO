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

        public void Note(string s)
        {
            if (s is null) return;
            if (s.Length > 32767)
            {
                throw new ArgumentException("String exceeds Excel's limit of 32,767 characters", nameof(s));
            }

            _distinct.TryAdd(s, 0);
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
        public void Fixup(ref (int Row, int Col, CellValue Val, EnumValue<DocumentFormat.OpenXml.Spreadsheet.CellValues> Type) prepared)
        {
            if (_finalIndex is null) return;
            if (prepared.Type?.Value != DocumentFormat.OpenXml.Spreadsheet.CellValues.SharedString) return;

            // prepared.Val.Text currently holds the raw string; replace with index text
            var text = prepared.Val?.Text ?? string.Empty;
            if (text is null) text = string.Empty;
            if (_finalIndex.TryGetValue(text, out int idx))
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
        public void ApplyAndFixup((int Row, int Col, CellValue Val, EnumValue<DocumentFormat.OpenXml.Spreadsheet.CellValues> Type)[] prepared, ExcelDocument doc)
        {
            ApplyTo(doc);
            for (int i = 0; i < prepared.Length; i++)
            {
                Fixup(ref prepared[i]);
            }
        }
    }
}

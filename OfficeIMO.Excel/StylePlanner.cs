using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel {
    /// <summary>
    /// Minimal planner for number formats/styles.
    /// Collects distinct number format strings during compute, then creates
    /// NumberingFormats and CellFormats once in the apply stage.
    /// </summary>
    internal sealed class StylePlanner {
        private readonly HashSet<string> _formats = new(StringComparer.Ordinal);
        private readonly object _formatsLock = new();
        private Dictionary<string, uint>? _cellFormatIndexByFormat;

        public void NoteNumberFormat(string? format) {
            if (string.IsNullOrWhiteSpace(format)) return;
            lock (_formatsLock) {
                _formats.Add(format!);
            }
        }

        public void ApplyTo(ExcelDocument doc) {
            if (_formats.Count == 0) {
                _cellFormatIndexByFormat = new Dictionary<string, uint>(0, StringComparer.Ordinal);
                return;
            }

            var workbookPart = doc.WorkbookPartRoot ?? throw new InvalidOperationException("WorkbookPart is null");
            WorkbookStylesPart? stylesPart = workbookPart.WorkbookStylesPart;
            if (stylesPart == null) {
                stylesPart = workbookPart.AddNewPart<WorkbookStylesPart>();
            }

            Stylesheet stylesheet = stylesPart.Stylesheet ??= new Stylesheet();

            // Ensure required collections exist
            stylesheet.Fonts ??= new Fonts(new DocumentFormat.OpenXml.Spreadsheet.Font());
            if (stylesheet.Fonts.Count == null) stylesheet.Fonts.Count = (uint)stylesheet.Fonts.Count();

            stylesheet.Fills ??= new Fills(new Fill(new PatternFill { PatternType = PatternValues.None }));
            if (stylesheet.Fills.Count == null) stylesheet.Fills.Count = (uint)stylesheet.Fills.Count();

            stylesheet.Borders ??= new Borders(new Border());
            if (stylesheet.Borders.Count == null) stylesheet.Borders.Count = (uint)stylesheet.Borders.Count();

            stylesheet.CellStyleFormats ??= new CellStyleFormats(new CellFormat());
            if (stylesheet.CellStyleFormats.Count == null) stylesheet.CellStyleFormats.Count = (uint)stylesheet.CellStyleFormats.Count();

            stylesheet.CellFormats ??= new CellFormats(new CellFormat());
            if (stylesheet.CellFormats.Count == null || stylesheet.CellFormats.Count.Value == 0) {
                stylesheet.CellFormats.Count = 1; // default format at index 0
            }

            stylesheet.NumberingFormats ??= new NumberingFormats();

            // Build or reuse NumberingFormats
            var numberFormats = stylesheet.NumberingFormats.Elements<NumberingFormat>().ToList();
            uint nextId = numberFormats.Any() ? numberFormats.Max(n => n.NumberFormatId!.Value) + 1U : 164U;
            uint numberFormatCount = stylesheet.NumberingFormats.Count?.Value ?? (uint)numberFormats.Count;
            var numFmtIdByFormat = new Dictionary<string, uint>(StringComparer.Ordinal);
            var existingNumberFormatIdByFormat = new Dictionary<string, uint>(numberFormats.Count, StringComparer.Ordinal);
            foreach (var numberFormat in numberFormats) {
                if (numberFormat.FormatCode?.Value is string code && numberFormat.NumberFormatId?.Value is uint id) {
                    existingNumberFormatIdByFormat[code] = id;
                }
            }

            foreach (var fmt in _formats) {
                if (!existingNumberFormatIdByFormat.TryGetValue(fmt, out uint id)) {
                    id = nextId++;
                    var numberingFormat = new NumberingFormat {
                        NumberFormatId = id,
                        FormatCode = fmt
                    };
                    stylesheet.NumberingFormats.Append(numberingFormat);
                    stylesheet.NumberingFormats.Count = ++numberFormatCount;
                }

                numFmtIdByFormat[fmt] = id;
            }

            // Create (or reuse) CellFormats that apply the numbering format
            var cellFormats = stylesheet.CellFormats.Elements<CellFormat>().ToList();
            uint cellFormatCount = stylesheet.CellFormats.Count?.Value ?? (uint)cellFormats.Count;
            var cellFormatIndexByFormat = new Dictionary<string, uint>(StringComparer.Ordinal);
            foreach (var kvp in numFmtIdByFormat) {
                string fmt = kvp.Key;
                uint id = kvp.Value;

                int idx = cellFormats.FindIndex(cf => cf.NumberFormatId != null && cf.NumberFormatId.Value == id && cf.ApplyNumberFormat != null && cf.ApplyNumberFormat.Value);
                if (idx == -1) {
                    var cf = new CellFormat {
                        NumberFormatId = id,
                        FontId = 0U,
                        FillId = 0U,
                        BorderId = 0U,
                        FormatId = 0U,
                        ApplyNumberFormat = true
                    };
                    stylesheet.CellFormats.Append(cf);
                    stylesheet.CellFormats.Count = ++cellFormatCount;
                    idx = cellFormats.Count;
                    cellFormats.Add(cf);
                }
                cellFormatIndexByFormat[fmt] = (uint)idx;
            }

            stylesPart.Stylesheet.Save();
            _cellFormatIndexByFormat = cellFormatIndexByFormat;
        }

        public bool TryGetCellFormatIndex(string? format, out uint index) {
            index = 0;
            if (string.IsNullOrWhiteSpace(format) || _cellFormatIndexByFormat == null)
                return false;
            return _cellFormatIndexByFormat.TryGetValue(format!, out index);
        }
    }
}


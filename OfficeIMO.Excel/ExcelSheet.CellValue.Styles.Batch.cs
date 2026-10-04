using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Threading;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        private static readonly AsyncLocal<StyleBatch?> CurrentStyleBatch = new();

        /// <summary>Indexes style primitives and saves the stylesheet once while a converter writes one sheet.</summary>
        internal IDisposable BeginStyleBatch() {
            var workbookPart = _excelDocument.WorkbookPartRoot ?? throw new InvalidOperationException("WorkbookPart is null");
            var stylesPart = workbookPart.WorkbookStylesPart ?? workbookPart.AddNewPart<WorkbookStylesPart>();
            Stylesheet stylesheet = stylesPart.Stylesheet ??= new Stylesheet();
            EnsureDefaultStylePrimitives(stylesheet);
            var batch = new StyleBatch(stylesPart, CurrentStyleBatch.Value);
            CurrentStyleBatch.Value = batch;
            return batch;
        }

        private static StyleBatch? BatchFor(Stylesheet stylesheet) =>
            CurrentStyleBatch.Value is { } batch && ReferenceEquals(batch.Stylesheet, stylesheet) ? batch : null;

        private static void SaveStylesheet(WorkbookStylesPart stylesPart) {
            if (stylesPart.Stylesheet is { } stylesheet && BatchFor(stylesheet) is not null) return;
            stylesPart.Stylesheet?.Save();
        }

        private sealed class StyleBatch : IDisposable {
            private readonly WorkbookStylesPart _part;
            private readonly StyleBatch? _parent;
            private List<CellFormat>? _cellFormats;
            private Dictionary<string, uint>? _cellFormatIndexes;
            private List<Fill>? _fills;
            private Dictionary<string, uint>? _fillIndexes;
            private List<DocumentFormat.OpenXml.Spreadsheet.Font>? _fonts;
            private Dictionary<string, uint>? _fontIndexes;
            private Dictionary<string, uint>? _numberFormatIds;
            private int _numberFormatCount = -1;
            private uint _nextNumberFormatId;
            private bool _disposed;

            internal StyleBatch(WorkbookStylesPart part, StyleBatch? parent) {
                _part = part;
                _parent = parent;
                Stylesheet = part.Stylesheet ?? throw new InvalidOperationException("Stylesheet is null");
            }

            internal Stylesheet Stylesheet { get; }

            internal CellFormat? CellFormat(uint index) {
                RefreshCellFormats();
                return index < _cellFormats!.Count ? _cellFormats[(int)index] : null;
            }

            internal uint CellFormatIndex(CellFormat candidate) {
                RefreshCellFormats();
                string key = candidate.OuterXml;
                if (_cellFormatIndexes!.TryGetValue(key, out uint index)) return index;
                var formats = Stylesheet.CellFormats!;
                index = (uint)_cellFormats!.Count;
                formats.Append(candidate);
                _cellFormats.Add(candidate);
                _cellFormatIndexes.Add(key, index);
                formats.Count = (uint)_cellFormats.Count;
                return index;
            }

            internal uint FillIndex(Fill candidate) {
                var fills = Stylesheet.Fills ??= new Fills();
                if (_fills is null || _fills.Count != fills.ChildElements.Count) {
                    _fills = fills.Elements<Fill>().ToList();
                    _fillIndexes = IndexXml(_fills);
                }
                string key = candidate.OuterXml;
                if (_fillIndexes!.TryGetValue(key, out uint index)) return index;
                index = (uint)_fills.Count;
                fills.Append(candidate);
                _fills.Add(candidate);
                _fillIndexes.Add(key, index);
                fills.Count = (uint)_fills.Count;
                return index;
            }

            internal DocumentFormat.OpenXml.Spreadsheet.Font? Font(uint index) {
                RefreshFonts();
                return index < _fonts!.Count ? _fonts[(int)index] : null;
            }

            internal uint FontIndex(DocumentFormat.OpenXml.Spreadsheet.Font candidate) {
                RefreshFonts();
                string key = candidate.OuterXml;
                if (_fontIndexes!.TryGetValue(key, out uint index)) return index;
                var fonts = Stylesheet.Fonts!;
                index = (uint)_fonts!.Count;
                fonts.Append(candidate);
                _fonts.Add(candidate);
                _fontIndexes.Add(key, index);
                fonts.Count = (uint)_fonts.Count;
                return index;
            }

            internal uint NumberFormatId(string code) {
                var formats = Stylesheet.NumberingFormats ??= new NumberingFormats();
                if (_numberFormatIds is null || _numberFormatCount != formats.ChildElements.Count) {
                    _numberFormatIds = new Dictionary<string, uint>(StringComparer.Ordinal);
                    uint maximumId = 163;
                    foreach (NumberingFormat format in formats.Elements<NumberingFormat>()) {
                        if (format.FormatCode?.Value is string existing && format.NumberFormatId?.Value is uint id) {
                            if (!_numberFormatIds.ContainsKey(existing)) _numberFormatIds.Add(existing, id);
                            maximumId = Math.Max(maximumId, id);
                        }
                    }
                    _numberFormatCount = formats.ChildElements.Count;
                    _nextNumberFormatId = checked(maximumId + 1);
                }
                if (_numberFormatIds.TryGetValue(code, out uint existingId)) return existingId;
                uint newId = _nextNumberFormatId++;
                formats.Append(new NumberingFormat { NumberFormatId = newId, FormatCode = code });
                _numberFormatIds.Add(code, newId);
                formats.Count = (uint)++_numberFormatCount;
                return newId;
            }

            private void RefreshCellFormats() {
                var formats = Stylesheet.CellFormats ??= new CellFormats(new CellFormat());
                if (_cellFormats is not null && _cellFormats.Count == formats.ChildElements.Count) return;
                _cellFormats = formats.Elements<CellFormat>().ToList();
                _cellFormatIndexes = IndexXml(_cellFormats);
            }

            private void RefreshFonts() {
                var fonts = Stylesheet.Fonts ??= new Fonts(new DocumentFormat.OpenXml.Spreadsheet.Font());
                if (_fonts is not null && _fonts.Count == fonts.ChildElements.Count) return;
                _fonts = fonts.Elements<DocumentFormat.OpenXml.Spreadsheet.Font>().ToList();
                _fontIndexes = IndexXml(_fonts);
            }

            private static Dictionary<string, uint> IndexXml<T>(List<T> elements) where T : DocumentFormat.OpenXml.OpenXmlElement {
                var indexes = new Dictionary<string, uint>(elements.Count, StringComparer.Ordinal);
                for (int index = 0; index < elements.Count; index++) {
                    string key = elements[index].OuterXml;
                    if (!indexes.ContainsKey(key)) indexes.Add(key, (uint)index);
                }
                return indexes;
            }

            public void Dispose() {
                if (_disposed) return;
                if (!ReferenceEquals(CurrentStyleBatch.Value, this))
                    throw new InvalidOperationException("Style batches must be disposed in their creation order.");
                CurrentStyleBatch.Value = _parent;
                _disposed = true;
                Stylesheet.Save();
            }
        }
    }
}

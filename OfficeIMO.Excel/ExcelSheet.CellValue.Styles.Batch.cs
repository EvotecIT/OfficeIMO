using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Threading;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        private static readonly AsyncLocal<StyleBatch?> CurrentStyleBatch = new();

        /// <summary>
        /// Applies or clears text wrapping for selected cells under one workbook
        /// write lock and saves the stylesheet once. Other cell formatting is preserved.
        /// </summary>
        /// <param name="cells">1-based row and column coordinates. Duplicate coordinates are applied once.</param>
        /// <param name="wrapText">Whether text should wrap within the selected cells.</param>
        /// <exception cref="ArgumentNullException">The cell selection is null.</exception>
        /// <exception cref="ArgumentOutOfRangeException">A coordinate is outside the worksheet. The selection is validated before editing.</exception>
        public void CellWrapTextFor(IEnumerable<(int Row, int Column)> cells, bool wrapText = true) {
            if (cells == null) throw new ArgumentNullException(nameof(cells));
            var selected = cells.Distinct().ToArray();
            foreach (var cell in selected) {
                if (cell.Row < 1 || cell.Row > A1.MaxRows || cell.Column < 1 || cell.Column > A1.MaxColumns) {
                    throw new ArgumentOutOfRangeException(nameof(cells), "Cell coordinates must be within the worksheet.");
                }
            }
            if (selected.Length == 0) return;

            WriteLockConditional(() => {
                using (BeginStyleBatch()) {
                    foreach (var cell in selected) CellWrapText(cell.Row, cell.Column, wrapText);
                }
            });
        }

        /// <summary>Indexes style primitives and saves the stylesheet once while a converter writes one sheet.</summary>
        internal IDisposable BeginStyleBatch() {
            var workbookPart = _excelDocument.WorkbookPartRoot ?? throw new InvalidOperationException("WorkbookPart is null");
            var stylesPart = workbookPart.WorkbookStylesPart ?? workbookPart.AddNewPart<WorkbookStylesPart>();
            Stylesheet stylesheet = stylesPart.Stylesheet ??= new Stylesheet();
            EnsureDefaultStylePrimitives(stylesheet);
            var batch = new StyleBatch(stylesPart, CurrentStyleBatch.Value, stylesheet) { PrimitivesEnsured = true };
            CurrentStyleBatch.Value = batch;
            return batch;
        }

        /// <summary>
        /// Starts a style scope for <see cref="Batch"/> and range edits when the workbook has a styles part. The scope
        /// binds to the stylesheet only when a style call reaches it and saves the stylesheet only when it changed, so
        /// a batch that writes values alone neither loads nor rewrites the stylesheet. Number formats keep the rule of
        /// a single call, so cell formats are identical with or without the batch.
        /// </summary>
        private IDisposable? TryBeginStyleBatch() {
            WorkbookStylesPart? stylesPart = _excelDocument.WorkbookPartRoot?.WorkbookStylesPart;
            if (stylesPart is null) return null;
            var batch = new StyleBatch(stylesPart, CurrentStyleBatch.Value, stylesheet: null);
            CurrentStyleBatch.Value = batch;
            return batch;
        }

        /// <summary>
        /// Runs a range edit under one write lock and one lazily bound style scope. Unlike <see cref="Batch"/>, it
        /// joins a write lock that another sheet's batch already holds.
        /// </summary>
        internal void RangeEdit(Action edit) {
            WriteLockConditional(() => {
                using (TryBeginStyleBatch()) {
                    edit();
                }
            });
        }

        private static StyleBatch? BatchFor(Stylesheet stylesheet) =>
            CurrentStyleBatch.Value is { } batch && batch.Matches(stylesheet) ? batch : null;

        private static void SaveStylesheet(WorkbookStylesPart stylesPart) {
            if (stylesPart.Stylesheet is { } stylesheet && BatchFor(stylesheet) is { } batch) {
                batch.MarkModified();
                return;
            }
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
            private Dictionary<uint, uint>? _wrapTextFormatIndexes;
            private Dictionary<(uint BaseIndex, string Operation), uint>? _derivedFormatIndexes;
            private int _numberFormatCount = -1;
            private uint _nextNumberFormatId;
            private readonly bool _saveOnDispose;
            private Stylesheet? _stylesheet;
            private bool _modified;
            private bool _disposed;

            /// <param name="part">Styles part the batch indexes.</param>
            /// <param name="parent">Batch to restore when this one ends.</param>
            /// <param name="stylesheet">The loaded stylesheet for an eager batch that always saves it when it ends;
            /// null for a lazy scope that binds on the first style call and saves only a changed stylesheet.</param>
            internal StyleBatch(WorkbookStylesPart part, StyleBatch? parent, Stylesheet? stylesheet) {
                _part = part;
                _parent = parent;
                _stylesheet = stylesheet;
                _saveOnDispose = stylesheet is not null;
            }

            internal Stylesheet Stylesheet => _stylesheet ?? throw new InvalidOperationException("The style batch is not bound to a stylesheet.");

            /// <summary>Whether the default fonts, fills, borders and formats were ensured for this batch.</summary>
            internal bool PrimitivesEnsured { get; set; }

            /// <summary>Whether number formats are resolved through the batch index. Lazy scopes keep the rule of a single call.</summary>
            internal bool IndexesNumberFormats => _saveOnDispose;

            /// <summary>Whether this batch serves <paramref name="stylesheet"/>; a lazy scope binds to it on first use.</summary>
            internal bool Matches(Stylesheet stylesheet) {
                if (_stylesheet is not null) return ReferenceEquals(_stylesheet, stylesheet);
                if (!ReferenceEquals(stylesheet.OpenXmlPart, _part)) return false;
                _stylesheet = stylesheet;
                return true;
            }

            /// <summary>Records that the stylesheet changed, so the batch saves it when it ends.</summary>
            internal void MarkModified() => _modified = true;

            /// <summary>
            /// Looks up the format index an <paramref name="operation"/> produced earlier in this batch for a cell that
            /// started from <paramref name="baseIndex"/>. Formats, fonts and fills are only appended while the batch is
            /// open, so the same starting format and operation always resolve to the same index.
            /// </summary>
            internal bool TryGetDerivedFormat(uint baseIndex, string operation, out uint index) {
                RefreshCellFormats();
                return _derivedFormatIndexes!.TryGetValue((baseIndex, operation), out index);
            }

            /// <summary>Records the format index <paramref name="operation"/> produced for <paramref name="baseIndex"/>.</summary>
            internal void SetDerivedFormat(uint baseIndex, string operation, uint index) {
                RefreshCellFormats();
                _derivedFormatIndexes![(baseIndex, operation)] = index;
            }

            /// <summary>
            /// Returns the wrap-text format for cells that start from <paramref name="baseIndex"/>, with the same
            /// first-match rule and appended format as <see cref="ApplyWrapText(Cell)"/> outside a batch.
            /// </summary>
            internal uint WrapTextFormatIndex(uint baseIndex) {
                RefreshCellFormats();
                if (_wrapTextFormatIndexes!.TryGetValue(baseIndex, out uint cached)) return cached;
                CellFormat baseFormat = baseIndex < _cellFormats!.Count ? _cellFormats[(int)baseIndex] : CreateDefaultCellFormat();
                int index = FindWrapTextFormat(_cellFormats!, baseFormat);
                if (index < 0) {
                    CellFormat candidate = CreateWrapTextFormat(baseFormat);
                    var formats = Stylesheet.CellFormats!;
                    index = _cellFormats!.Count;
                    formats.Append(candidate);
                    _cellFormats.Add(candidate);
                    string key = candidate.OuterXml;
                    if (!_cellFormatIndexes!.ContainsKey(key)) _cellFormatIndexes.Add(key, (uint)index);
                    formats.Count = (uint)_cellFormats.Count;
                    _modified = true;
                }
                _wrapTextFormatIndexes[baseIndex] = (uint)index;
                return (uint)index;
            }

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
                _modified = true;
                return index;
            }

            internal uint FillIndex(Fill candidate) {
                var fills = Stylesheet.Fills ??= new Fills();
                if (_fills is null || _fills.Count != fills.ChildElements.Count) {
                    _fills = fills.Elements<Fill>().ToList();
                    _fillIndexes = IndexXml(_fills);
                    _derivedFormatIndexes?.Clear();
                }
                string key = candidate.OuterXml;
                if (_fillIndexes!.TryGetValue(key, out uint index)) return index;
                index = (uint)_fills.Count;
                fills.Append(candidate);
                _fills.Add(candidate);
                _fillIndexes.Add(key, index);
                fills.Count = (uint)_fills.Count;
                _modified = true;
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
                _modified = true;
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
                _modified = true;
                return newId;
            }

            private void RefreshCellFormats() {
                var formats = Stylesheet.CellFormats ??= new CellFormats(new CellFormat());
                if (_cellFormats is not null && _cellFormats.Count == formats.ChildElements.Count) return;
                _cellFormats = formats.Elements<CellFormat>().ToList();
                _cellFormatIndexes = IndexXml(_cellFormats);
                _wrapTextFormatIndexes = new Dictionary<uint, uint>();
                _derivedFormatIndexes = new Dictionary<(uint BaseIndex, string Operation), uint>();
            }

            private void RefreshFonts() {
                var fonts = Stylesheet.Fonts ??= new Fonts(new DocumentFormat.OpenXml.Spreadsheet.Font());
                if (_fonts is not null && _fonts.Count == fonts.ChildElements.Count) return;
                _fonts = fonts.Elements<DocumentFormat.OpenXml.Spreadsheet.Font>().ToList();
                _fontIndexes = IndexXml(_fonts);
                _derivedFormatIndexes?.Clear();
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
                if (_stylesheet is null || !(_saveOnDispose || _modified)) return;
                // A stylesheet replaced inside the batch was saved by whoever replaced it; never write the old one back.
                if (ReferenceEquals(_part.Stylesheet, _stylesheet)) _stylesheet.Save();
            }
        }
    }
}

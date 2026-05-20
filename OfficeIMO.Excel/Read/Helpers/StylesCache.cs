using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel {
    internal sealed class StylesCache {
        private static readonly bool[] EmptyDateStyleIndexes = Array.Empty<bool>();
        private bool[] _dateStyleIndexes = EmptyDateStyleIndexes;

        private StylesCache() { }

        public bool HasDateStyles { get; private set; }

        public static StylesCache Build(SpreadsheetDocument doc) {
            var cache = new StylesCache();
            var sp = doc.WorkbookPart!.WorkbookStylesPart;
            if (sp?.Stylesheet == null) return cache;

            var nf = new Dictionary<uint, string>();
            var numbering = sp.Stylesheet.NumberingFormats;
            if (numbering != null) {
                foreach (var n in numbering.Elements<NumberingFormat>()) {
                    if (n.NumberFormatId?.Value is uint id && n.FormatCode?.Value is string code)
                        nf[id] = code;
                }
            }

            static bool IsBuiltInDate(uint id)
                => id is 14 or 15 or 16 or 17 or 18 or 19 or 20 or 21 or 22
                    or 27 or 30 or 36 or 45 or 46 or 47;

            var xfs = sp.Stylesheet.CellFormats;
            if (xfs != null) {
                var cellFormats = xfs.Elements<CellFormat>().ToList();
                if (cellFormats.Count > 0) {
                    cache._dateStyleIndexes = new bool[cellFormats.Count];
                }

                for (int idx = 0; idx < cellFormats.Count; idx++) {
                    var cf = cellFormats[idx];
                    var nId = (uint)(cf.NumberFormatId?.Value ?? 0);
                    bool dateLike = IsBuiltInDate(nId) || (nf.TryGetValue(nId, out var code) && ExcelNumberFormatClassifier.LooksLikeDateFormat(code));
                    if (dateLike) {
                        cache._dateStyleIndexes[idx] = true;
                        cache.HasDateStyles = true;
                    }
                }
            }

            return cache;
        }

        public bool IsDateLike(uint styleIndex) => styleIndex < (uint)_dateStyleIndexes.Length && _dateStyleIndexes[styleIndex];
    }
}


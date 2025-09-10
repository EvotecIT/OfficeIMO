using System.Collections.Generic;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel
{
    internal sealed class StylesCache
    {
        private readonly HashSet<uint> _dateStyleIdx = new();

        private StylesCache() { }

        public static StylesCache Build(SpreadsheetDocument doc)
        {
            var cache = new StylesCache();
            var sp = doc.WorkbookPart!.WorkbookStylesPart;
            if (sp?.Stylesheet == null) return cache;

            var nf = new Dictionary<uint, string>();
            var numbering = sp.Stylesheet.NumberingFormats;
            if (numbering != null)
            {
                foreach (var n in numbering.Elements<NumberingFormat>())
                {
                    if (n.NumberFormatId?.Value is uint id && n.FormatCode?.Value is string code)
                        nf[id] = code;
                }
            }

            static bool IsBuiltInDate(uint id)
                => id is 14 or 15 or 16 or 17 or 18 or 19 or 20 or 21 or 22
                    or 27 or 30 or 36 or 45 or 46 or 47;

            var xfs = sp.Stylesheet.CellFormats;
            if (xfs != null)
            {
                uint idx = 0;
                foreach (var cf in xfs.Elements<CellFormat>())
                {
                    var nId = (uint)(cf.NumberFormatId?.Value ?? 0);
                    bool dateLike = IsBuiltInDate(nId) || (nf.TryGetValue(nId, out var code) && LooksLikeDateFormat(code));
                    if (dateLike) cache._dateStyleIdx.Add(idx);
                    idx++;
                }
            }

            return cache;
        }

        public bool IsDateLike(uint styleIndex) => _dateStyleIdx.Contains(styleIndex);

        private static bool LooksLikeDateFormat(string code)
        {
            var lower = code.ToLowerInvariant();
            if (lower.IndexOf('d') >= 0 || lower.IndexOf('y') >= 0 || lower.IndexOf('h') >= 0 || lower.IndexOf('s') >= 0)
                return true;
            if (lower.Contains('m') && (lower.Contains('d') || lower.Contains('y') || lower.Contains('h')))
                return true;
            return false;
        }
    }
}


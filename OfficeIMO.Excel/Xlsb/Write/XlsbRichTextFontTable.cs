using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel.Xlsb.Model;

namespace OfficeIMO.Excel.Xlsb.Write;

/// <summary>Turns rich cell text into RichStr formatting runs and the extra BrtFont records those runs reference.</summary>
internal sealed class XlsbRichTextFontTable {
    private readonly Font[] _baseFonts;
    private readonly CellFormat[] _cellFormats;
    private readonly SharedStringTable? _sharedStrings;
    private readonly Dictionary<string, ushort> _fontIndexes = new(StringComparer.Ordinal);
    private readonly List<Font> _addedFonts = new();

    internal XlsbRichTextFontTable(Stylesheet stylesheet, SharedStringTable? sharedStrings) {
        _baseFonts = stylesheet.Fonts?.Elements<Font>().ToArray() ?? Array.Empty<Font>();
        _cellFormats = stylesheet.CellFormats?.Elements<CellFormat>().ToArray() ?? Array.Empty<CellFormat>();
        _sharedStrings = sharedStrings;
        for (int index = 0; index < _baseFonts.Length && index <= ushort.MaxValue; index++) {
            string key = _baseFonts[index].OuterXml;
            if (!_fontIndexes.ContainsKey(key)) _fontIndexes.Add(key, (ushort)index);
        }
    }

    /// <summary>Fonts that rich runs need beyond the workbook's own font collection, in index order.</summary>
    internal IReadOnlyList<Font> AddedFonts => _addedFonts;

    /// <summary>Returns the runs of a cell whose text has run formatting, or null for plain text.</summary>
    internal XlsbTextRun[]? CreateRuns(Cell cell, uint styleIndex, string text) {
        Run[] items = GetRuns(cell, _sharedStrings);
        if (!HasFormattedRuns(items)) return null;

        ushort cellFont = GetCellFont(styleIndex);
        var result = new List<XlsbTextRun>(items.Length);
        int position = 0;
        foreach (Run run in items) {
            int length = Utilities.ExcelSanitizer.DecodeXString(run.Text?.Text ?? string.Empty).Length;
            if (length == 0) continue;
            ushort font = run.RunProperties == null || !run.RunProperties.HasChildren ? cellFont : GetOrAddFont(run.RunProperties);
            if (result.Count == 0 || result[result.Count - 1].Font != font) result.Add(new XlsbTextRun(checked((ushort)position), font));
            position += length;
        }
        if (position != text.Length) {
            throw new NotSupportedException("Native XLSB generation found rich cell text whose runs do not match the cell text.");
        }
        return result.Count == 0 || (result.Count == 1 && result[0].Font == cellFont) ? null : result.ToArray();
    }

    /// <summary>Whether a cell's inline or shared text carries run formatting.</summary>
    internal static bool HasFormattedRuns(Cell cell, SharedStringTable? sharedStrings) => HasFormattedRuns(GetRuns(cell, sharedStrings));

    private static bool HasFormattedRuns(Run[] runs) => runs.Any(run => run.RunProperties?.HasChildren == true);

    private static Run[] GetRuns(Cell cell, SharedStringTable? sharedStrings) {
        IEnumerable<Run>? runs = cell.InlineString?.Elements<Run>();
        if (runs == null && cell.DataType?.Value == CellValues.SharedString
            && int.TryParse(cell.CellValue?.InnerText, System.Globalization.NumberStyles.None, System.Globalization.CultureInfo.InvariantCulture, out int index)) {
            runs = sharedStrings?.Elements<SharedStringItem>().ElementAtOrDefault(index)?.Elements<Run>();
        }
        return runs?.ToArray() ?? Array.Empty<Run>();
    }

    private ushort GetCellFont(uint styleIndex) {
        uint font = styleIndex < _cellFormats.Length ? _cellFormats[styleIndex].FontId?.Value ?? 0U : 0U;
        return checked((ushort)font);
    }

    private ushort GetOrAddFont(RunProperties properties) {
        Font font = CreateFont(properties);
        string key = font.OuterXml;
        if (_fontIndexes.TryGetValue(key, out ushort existing)) return existing;
        int index = _baseFonts.Length + _addedFonts.Count;
        if (index > ushort.MaxValue) throw new NotSupportedException("Native XLSB generation found more rich-text fonts than the BrtFont index can address.");
        _addedFonts.Add(font);
        _fontIndexes.Add(key, (ushort)index);
        return (ushort)index;
    }

    // CT_RPrElt carries the CT_Font properties under run names; emit them in CT_Font's sequence.
    private static Font CreateFont(RunProperties properties) {
        var font = new Font();
        void Add<T>() where T : OpenXmlElement {
            T? value = properties.GetFirstChild<T>();
            if (value != null) font.Append(value.CloneNode(true));
        }
        Add<Bold>();
        Add<Italic>();
        Add<Strike>();
        Add<Condense>();
        Add<Extend>();
        Add<Outline>();
        Add<Shadow>();
        Add<Underline>();
        Add<VerticalTextAlignment>();
        Add<FontSize>();
        Add<Color>();
        if (properties.GetFirstChild<RunFont>()?.Val?.Value is string name) font.Append(new FontName { Val = name });
        if (properties.GetFirstChild<FontFamily>()?.Val?.Value is int family) font.Append(new FontFamilyNumbering { Val = family });
        if (properties.GetFirstChild<RunPropertyCharSet>()?.Val?.Value is int charset) font.Append(new FontCharSet { Val = charset });
        Add<FontScheme>();
        return font;
    }
}

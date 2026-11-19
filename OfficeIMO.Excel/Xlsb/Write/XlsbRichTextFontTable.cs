using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel.Xlsb.Model;

namespace OfficeIMO.Excel.Xlsb.Write;

/// <summary>Turns rich cell text into RichStr formatting runs and the extra BrtFont records those runs reference.</summary>
internal sealed class XlsbRichTextFontTable {
    private readonly Font[] _baseFonts;
    private readonly CellFormat[] _cellFormats;
    private readonly Dictionary<string, ushort> _fontIndexes = new(StringComparer.Ordinal);
    private readonly List<Font> _addedFonts = new();

    internal XlsbRichTextFontTable(Stylesheet stylesheet, XlsbCellTextRuns textRuns) {
        _baseFonts = stylesheet.Fonts?.Elements<Font>().ToArray() ?? Array.Empty<Font>();
        _cellFormats = stylesheet.CellFormats?.Elements<CellFormat>().ToArray() ?? Array.Empty<CellFormat>();
        TextRuns = textRuns;
        for (int index = 0; index < _baseFonts.Length && index <= ushort.MaxValue; index++) {
            string key = _baseFonts[index].OuterXml;
            if (!_fontIndexes.ContainsKey(key)) _fontIndexes.Add(key, (ushort)index);
        }
    }

    /// <summary>The shared-string snapshot the runs are read from.</summary>
    internal XlsbCellTextRuns TextRuns { get; }

    /// <summary>Fonts that rich runs need beyond the workbook's own font collection, in index order.</summary>
    internal IReadOnlyList<Font> AddedFonts => _addedFonts;

    /// <summary>Returns the runs of a cell whose text has run formatting, or null for plain text.</summary>
    internal XlsbTextRun[]? CreateRuns(Cell cell, uint styleIndex, string text) {
        Run[] items = TextRuns.GetRuns(cell);
        if (!XlsbCellTextRuns.HasFormattedRuns(items)) return null;

        ushort cellFont = GetCellFont(styleIndex);
        var result = new List<XlsbTextRun>(items.Length);
        int position = 0;
        foreach (Run run in items) {
            int length = Utilities.ExcelSanitizer.DecodeXString(run.Text?.Text ?? string.Empty).Length;
            if (length == 0) continue;
            ushort font = run.RunProperties == null || !run.RunProperties.HasChildren ? cellFont : GetOrAddFont(run.RunProperties, cellFont);
            if (result.Count == 0 || result[result.Count - 1].Font != font) result.Add(new XlsbTextRun(checked((ushort)position), font));
            position += length;
        }
        if (position != text.Length) {
            throw new NotSupportedException("Native XLSB generation found rich cell text whose runs do not match the cell text.");
        }
        return result.Count == 0 || (result.Count == 1 && result[0].Font == cellFont) ? null : result.ToArray();
    }

    private ushort GetCellFont(uint styleIndex) {
        uint font = styleIndex < _cellFormats.Length ? _cellFormats[styleIndex].FontId?.Value ?? 0U : 0U;
        return checked((ushort)font);
    }

    private ushort GetOrAddFont(RunProperties properties, ushort cellFont) {
        Font font = CreateFont(properties, cellFont < _baseFonts.Length ? _baseFonts[cellFont] : null);
        string key = font.OuterXml;
        if (_fontIndexes.TryGetValue(key, out ushort existing)) return existing;
        int index = _baseFonts.Length + _addedFonts.Count;
        if (index > ushort.MaxValue) throw new NotSupportedException("Native XLSB generation found more rich-text fonts than the BrtFont index can address.");
        _addedFonts.Add(font);
        _fontIndexes.Add(key, (ushort)index);
        return (ushort)index;
    }

    // CT_RPrElt carries the CT_Font properties under run names; emit them in CT_Font's sequence. A run that omits
    // the typeface or size keeps the cell font's, as a missing ExcelRichTextRun font or size inherits it.
    private static Font CreateFont(RunProperties properties, Font? cellFont) {
        var font = new Font();
        void Add<T>() where T : OpenXmlElement {
            T? value = properties.GetFirstChild<T>();
            if (value != null) font.Append(value.CloneNode(true));
        }
        void Inherit<T>() where T : OpenXmlElement {
            T? value = cellFont?.GetFirstChild<T>();
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
        if (properties.GetFirstChild<FontSize>() != null) Add<FontSize>(); else Inherit<FontSize>();
        Add<Color>();
        RunFont? name = properties.GetFirstChild<RunFont>();
        if (name?.Val?.Value is string typeface) {
            font.Append(new FontName { Val = typeface });
            if (properties.GetFirstChild<FontFamily>()?.Val?.Value is int family) font.Append(new FontFamilyNumbering { Val = family });
            if (properties.GetFirstChild<RunPropertyCharSet>()?.Val?.Value is int charset) font.Append(new FontCharSet { Val = charset });
            Add<FontScheme>();
        } else {
            // The family, character set and scheme describe the typeface, so they follow it from the cell font.
            Inherit<FontName>();
            Inherit<FontFamilyNumbering>();
            Inherit<FontCharSet>();
            Inherit<FontScheme>();
        }
        return font;
    }
}

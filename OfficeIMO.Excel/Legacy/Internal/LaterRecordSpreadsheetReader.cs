using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;

namespace OfficeIMO.Excel.Legacy;

internal enum LaterSpreadsheetProfile { Lotus, QuattroWb1, WorksWindows }

/// <summary>Generation-specific record envelopes projected into the existing legacy workbook model.</summary>
internal sealed partial class LaterRecordSpreadsheetReader {
    private readonly byte[] _data;
    private readonly OfficeLegacyImportLimits _limits;
    private readonly CancellationToken _cancellation;
    private readonly LaterSpreadsheetProfile _profile;
    private readonly LegacySpreadsheetModel _model = new() { Quality = OfficeLegacyImportQuality.Structured, PreserveSheetNames = true };
    private readonly Dictionary<int, string> _names = new();
    private readonly Dictionary<(int Sheet, int Row, int Column), LegacySpreadsheetCell> _cells = new();
    private readonly HashSet<(int Sheet, int Row, int Column)> _commentOnly = new();
    private readonly HashSet<(int Sheet, int Row, int Column)> _formulas = new();
    private readonly HashSet<string> _losses = new();
    private int _text;
    private int _sheet;
    private LaterRecordSpreadsheetReader(byte[] data, OfficeLegacyImportLimits limits, CancellationToken cancellation, LaterSpreadsheetProfile profile) {
        _data = data; _limits = limits; _cancellation = cancellation; _profile = profile;
    }
    internal static bool IsLotus(byte[] data) => data.Length >= 30 && U16(data, 0) == 0 && U16(data, 2) == 26 &&
        (U16(data, 4) == 0x1000 || U16(data, 4) == 0x1002 || U16(data, 4) == 0x1005);
    internal static bool IsQuattro(byte[] data) => data.Length >= 6 && U16(data, 0) == 0 && U16(data, 2) == 2 && U16(data, 4) == 0x1001;
    internal static bool IsWorks(byte[] data) => data.Length >= 6 && U16(data, 0) == 0xff && U16(data, 2) == 2 && U16(data, 4) == 0x0404;
    internal static LegacySpreadsheetModel Read(byte[] data, OfficeLegacyImportLimits limits, CancellationToken cancellation, LaterSpreadsheetProfile profile) =>
        new LaterRecordSpreadsheetReader(data, limits, cancellation, profile).Read();

    private LegacySpreadsheetModel Read() {
        var records = new List<(ushort Type, int Offset, int Length)>();
        int p = 0, count = 0; bool eof = false;
        if (_profile == LaterSpreadsheetProfile.Lotus) {
            int lastSheet = _data[14];
            for (int id = 0; id <= lastSheet; id++) EnsureSheet(id);
        }
        while (p < _data.Length) {
            _cancellation.ThrowIfCancellationRequested();
            if (++count > _limits.MaxRecords) throw new InvalidDataException("Legacy workbook exceeds the record limit.");
            Require(p, 4, _data.Length);
            ushort type = U16(_data, p); int length = U16(_data, p + 2), payload = p + 4;
            Require(payload, length, _data.Length);
            if (p != 0 && (type == 0 || type == 0xff && _profile == LaterSpreadsheetProfile.WorksWindows))
                throw new InvalidDataException("Duplicate workbook BOF.");
            if (type == 1) { if (length != 0) throw new InvalidDataException("Invalid workbook EOF."); eof = true; p = payload; break; }
            if (_profile == LaterSpreadsheetProfile.Lotus && type == 2 || _profile == LaterSpreadsheetProfile.QuattroWb1 && type == 0x4b)
                throw new InvalidDataException("Encrypted legacy workbook is not supported.");
            if (_profile == LaterSpreadsheetProfile.Lotus && type == 0x23) {
                Require(payload, 5, payload + length); SetName(_data[payload + 2], Text(payload + 4, length - 4));
            } else if (_profile == LaterSpreadsheetProfile.QuattroWb1 && type == 0xca) {
                if (length != 1) throw new InvalidDataException("Invalid Quattro sheet selector.");
                _sheet = _data[payload]; EnsureSheet(_sheet);
            } else if (_profile == LaterSpreadsheetProfile.QuattroWb1 && type == 0xcc) SetName(_sheet, Text(payload, length));
            else if (IsCell(type)) records.Add((type, payload, length));
            else if (p != 0) {
                Loss("LATER_RECORDS_UNSUPPORTED", "Source formatting, chart, name, view and application records outside the qualified cell profile are not reconstructed.");
                if (type == 0x97) _model.InertContent |= OfficeLegacyInertContentKind.ExternalLinks;
                if (type == 0x38e || type == 0x10d) _model.InertContent |= OfficeLegacyInertContentKind.EmbeddedObjects;
            }
            p = payload + length;
        }
        if (!eof) throw new InvalidDataException("Legacy workbook has no complete EOF record.");
        if (p != _data.Length) {
            if (_profile != LaterSpreadsheetProfile.Lotus && _data.Skip(p).Any(value => value != 0 && value != 0x1a))
                throw new InvalidDataException("Unexpected data following workbook EOF.");
            // Lotus stores document properties and application objects after the main record stream.
            if (_profile == LaterSpreadsheetProfile.Lotus) {
                Loss("LOTUS_TRAILING_APPLICATION_DATA", "Application data following the Lotus main-stream EOF is retained inert and is not reconstructed.");
                _model.InertContent |= OfficeLegacyInertContentKind.EmbeddedObjects;
            }
        }
        foreach (var record in records) { _cancellation.ThrowIfCancellationRequested(); ReadCell(record.Type, record.Offset, record.Length); }
        if (_names.Count == 0) EnsureSheet(0);
        foreach (var entry in _names.OrderBy(item => item.Key)) {
            var sheet = new LegacySpreadsheetSheet(entry.Value);
            sheet.Cells.AddRange(_cells.Where(cell => cell.Key.Sheet == entry.Key).OrderBy(cell => cell.Key.Row).ThenBy(cell => cell.Key.Column).Select(cell => cell.Value));
            _model.Sheets.Add(sheet);
        }
        _model.RecoveredCellCount = _cells.Count;
        return _model;
    }

    private bool IsCell(ushort type) => _profile == LaterSpreadsheetProfile.Lotus ?
        type == 0x16 || type == 0x17 || type == 0x18 || type == 0x19 || type == 0x1a || type == 0x25 || type == 0x26 || type == 0x27 || type == 0x28 :
        type >= 0xc && type <= 0x10 || type == 0x33 && _profile == LaterSpreadsheetProfile.QuattroWb1 ||
        (type == 0x545b || type == 0x36) && _profile == LaterSpreadsheetProfile.WorksWindows;

    private void ReadCell(ushort type, int p, int length) {
        bool lotus = _profile == LaterSpreadsheetProfile.Lotus;
        int header = lotus ? 4 : 6, end = p + length; Require(p, header, end);
        int row = lotus ? U16(_data, p) : U16(_data, p + 2), column = lotus ? _data[p + 3] : _data[p], sheet = lotus ? _data[p + 2] : _data[p + 1];
        if (!lotus && row > 32767) throw new InvalidDataException("Unsupported signed source row address.");
        if (_profile == LaterSpreadsheetProfile.WorksWindows && sheet != 0) throw new InvalidDataException("Unsupported Works sheet identifier.");
        EnsureSheet(sheet); var key = (sheet, row, column);
        object? value = null; string? formula = null, comment = null; ExcelHorizontalAlignment? alignment = null;
        int q = p + header;
        bool text = lotus ? type == 0x16 || type == 0x1a || type == 0x26 : type == 0xf || type == 0x33 || type == 0x36;
        if (text) {
            if (q == end) throw new InvalidDataException("Truncated text cell.");
            if (lotus || _profile == LaterSpreadsheetProfile.QuattroWb1) {
                byte first = _data[q];
                if (first == '\'' || first == '^' || first == '"' || first == '\\') {
                    alignment = first == '^' ? ExcelHorizontalAlignment.Center : first == '"' ? ExcelHorizontalAlignment.Right : ExcelHorizontalAlignment.Left; q++;
                } else if (!lotus) { if (first != 0) Loss("LATER_LABEL_ALIGNMENT", "A source label alignment is not supported."); q++; }
            }
            value = Text(q, end - q);
            if (type == 0x26 && lotus) { comment = (string)value; value = null; }
        } else if (type != 0xc) {
            LaterNumberKind kind = lotus ? type == 0x18 ? LaterNumberKind.Compact16 : type == 0x25 ? LaterNumberKind.Compact32 :
                type == 0x17 || type == 0x19 ? LaterNumberKind.Extended80 : LaterNumberKind.Double64 :
                type == 0x545b ? LaterNumberKind.Single32 : LaterNumberKind.Double64;
            int size = !lotus && type == 0xd ? 2 : kind == LaterNumberKind.Extended80 ? 10 : kind == LaterNumberKind.Double64 ? 8 : kind == LaterNumberKind.Compact16 ? 2 : 4;
            Require(q, size, end);
            double number = !lotus && type == 0xd ? (short)U16(_data, q) : LaterSpreadsheetNumbers.Read(_data, q, kind);
            bool isFormula = lotus ? type == 0x19 || type == 0x28 : type == 0x10;
            if (!double.IsNaN(number) && !double.IsInfinity(number)) value = number;
            else if (!isFormula) throw new InvalidDataException("Non-finite numeric cell value.");
            else Loss("LATER_FORMULA_NON_NUMERIC_CACHE", "A formula has a source error or non-numeric cache; supported formulas or subsequent string-result records are retained.");
            q += size;
            if (isFormula) { formula = Formula(q, end, row, column, sheet, type == 0x19); }
            else if (q != end) throw new InvalidDataException("Numeric cell size does not match its record.");
        } else if (q != end) throw new InvalidDataException("Invalid blank cell size.");
        if (_cells.TryGetValue(key, out LegacySpreadsheetCell? previous)) {
            if (_commentOnly.Remove(key) && comment == null) { comment = previous.Comment; }
            else if ((lotus && type == 0x1a || !lotus && type == 0x33) && _formulas.Contains(key)) { formula = previous.Formula; comment = previous.Comment; alignment = previous.Alignment; }
            else if (comment != null && previous.Comment == null) { value = previous.Value; formula = previous.Formula; alignment = previous.Alignment; }
            else if (type == 0x36 && previous.Value is string prefix && !_formulas.Contains(key)) {
                string suffix = (string)value!;
                if (prefix.Length > 32767 - suffix.Length) throw new InvalidDataException("Continued cell text exceeds the XLSX limit.");
                value = prefix + suffix; comment = previous.Comment; alignment = previous.Alignment;
            }
            else throw new InvalidDataException("Duplicate source cell address.");
        } else if (type == 0x36 || lotus && type == 0x1a || !lotus && type == 0x33) throw new InvalidDataException("Source continuation or cached string lacks its preceding cell.");
        else if (comment != null) _commentOnly.Add(key);
        if (lotus ? type == 0x19 || type == 0x28 : type == 0x10) _formulas.Add(key);
        if (!_cells.ContainsKey(key) && _cells.Count >= _limits.MaxItems) throw new InvalidDataException("Legacy workbook exceeds the cell limit.");
        _cells[key] = new LegacySpreadsheetCell(row + 1, column + 1, value, formula, comment: comment, alignment: alignment);
    }

    private string? Formula(int p, int end, int row, int column, int sheet, bool extended) {
        var context = new LaterFormulaContext { IsLotus = _profile == LaterSpreadsheetProfile.Lotus, ExtendedNumbers = extended, Sheet = sheet, SheetName = EnsureSheet };
        if (_profile == LaterSpreadsheetProfile.WorksWindows) {
            Require(p, 2, end); int size = U16(_data, p); p += 2;
            if (size != end - p) throw new InvalidDataException("Works formula size does not match its record.");
        } else if (_profile == LaterSpreadsheetProfile.QuattroWb1) {
            Require(p, 6, end); p += 2; // Cached-result state.
            int size = U16(_data, p), tokens = U16(_data, p + 2); p += 4;
            if (size != end - p || tokens > size) throw new InvalidDataException("Invalid Quattro formula reference-table envelope.");
            try { context.References = QuattroReferences(p + tokens, end, row, column, sheet); }
            catch (InvalidDataException) { Loss("LATER_FORMULA_CACHED_FALLBACK", "Unsupported formula references retain their cached result without creating a live formula."); return null; }
            end = p + tokens;
        }
        if (WkFormulaDecoder.TryDecode(_data, p, end - p, row, column, _limits, _limits.MaxTextCharacters - _text, _cancellation, out string? formula, out _, context)) {
            CountText(formula!.Length); return formula;
        }
        Loss("LATER_FORMULA_CACHED_FALLBACK", "Unsupported or invalid formula tokens retain their cached result without creating a live formula."); return null;
    }

    private string EnsureSheet(int id) {
        if (id < 0 || id > 255) throw new InvalidDataException("Formula sheet reference is outside the source profile.");
        if (!_names.TryGetValue(id, out string? name)) {
            if (_names.Count >= _limits.MaxItems) throw new InvalidDataException("Legacy workbook exceeds the sheet limit.");
            string basis = "Sheet" + (id + 1).ToString(CultureInfo.InvariantCulture); name = basis; int suffix = 2;
            while (_names.Values.Contains(name, StringComparer.OrdinalIgnoreCase)) name = basis + " (" + (suffix++).ToString(CultureInfo.InvariantCulture) + ")";
            _names[id] = name;
        }
        return name;
    }
    private void SetName(int id, string name) {
        EnsureSheet(id);
        string trimmed = name.Trim();
        if (trimmed != name) Loss("LATER_SHEET_NAME", "A source sheet name was trimmed for XLSX; formula references use its final projected name.");
        name = trimmed;
        if (name.Length == 0 || name.Length > 31 || name.IndexOfAny(new[] { '[', ']', ':', '*', '?', '/', '\\' }) >= 0 ||
            name[0] == '\'' || name[name.Length - 1] == '\'' || _names.Any(item => item.Key != id && string.Equals(item.Value, name, StringComparison.OrdinalIgnoreCase))) {
            Loss("LATER_SHEET_NAME", "A source sheet name cannot be used safely in XLSX; its stable generated name is used."); return;
        }
        _names[id] = name;
    }
    private string Text(int p, int length) {
        int zero = Array.IndexOf(_data, (byte)0, p, length); if (zero < 0) throw new InvalidDataException("Unterminated source text.");
        int count = zero - p; if (count > 32767) throw new InvalidDataException("Source cell text exceeds the XLSX limit."); CountText(count);
        var text = new StringBuilder(count);
        for (int i = p; i < zero; i++) {
            if ((i & 255) == 0) _cancellation.ThrowIfCancellationRequested();
            byte value = _data[i];
            if (value > 127 || value < 32 && value != 9 && value != 10 && value != 13) { text.Append('\uFFFD'); Loss("LATER_TEXT_ENCODING", "Source characters outside the qualified ASCII profile are replaced explicitly."); }
            else text.Append((char)value);
        }
        return text.ToString();
    }
    private void CountText(int length) { if (length > _limits.MaxTextCharacters - _text) throw new InvalidDataException("Legacy workbook exceeds the text limit."); _text += length; }
    private void Loss(string code, string message) { if (_losses.Add(code)) _model.Findings.Add(LegacySpreadsheetAdapterBase.LossFinding(code, "Legacy workbook", message)); }
    private static ushort U16(byte[] data, int p) => (ushort)(data[p] | data[p + 1] << 8);
    private static void Require(int p, int size, int end) { if (p < 0 || size < 0 || p > end - size) throw new InvalidDataException("Truncated legacy workbook record."); }
}

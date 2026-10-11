using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;

namespace OfficeIMO.Excel.Legacy;

internal sealed partial class LaterRecordSpreadsheetReader {
    private IReadOnlyList<(string Text, bool Range)> QuattroReferences(int p, int end, int row, int column, int sheet) {
        var references = new List<(string, bool)>();
        while (p < end) {
            _cancellation.ThrowIfCancellationRequested(); Require(p, 2, end);
            int type = U16(_data, p); p += 2;
            if ((type & 0x0fff) != 0 || type >> 12 > 1) {
                if ((type & 0x3ff) != 0) _model.InertContent |= OfficeLegacyInertContentKind.ExternalLinks;
                throw new InvalidDataException("Named, deleted, external and collection references are outside the qualified formula profile.");
            }
            bool range = type == 0x1000;
            string first = QuattroReference(ref p, end, row, column, sheet);
            string value = range ? LaterFormulaContext.Range(first, QuattroReference(ref p, end, row, column, sheet)) : first;
            if (references.Count >= _limits.MaxItems) throw new InvalidDataException("Formula exceeds the reference limit.");
            references.Add((value, range));
        }
        return references;
    }
    private string QuattroReference(ref int p, int end, int row, int column, int sheet) {
        Require(p, 4, end); int c = _data[p], s = _data[p + 1], bits = U16(_data, p + 2); p += 4;
        bool relativeColumn = (bits & 0x4000) != 0, relativeRow = (bits & 0x2000) != 0;
        if ((bits & 0x8000) != 0) s = sheet + (sbyte)s;
        if (relativeColumn) c = column + (sbyte)c;
        int r = bits & 0x1fff; if (relativeRow) r = row + (r >= 4096 ? r - 8192 : r);
        if (c < 0 || c > 255 || r < 0 || r > 8191) throw new InvalidDataException("Invalid Quattro formula address.");
        string name = s == sheet ? "" : "'" + EnsureSheet(s).Replace("'", "''") + "'!";
        string letters = ""; for (int n = c + 1; n > 0; n /= 26) { n--; letters = (char)('A' + n % 26) + letters; }
        return name + (relativeColumn ? "" : "$") + letters + (relativeRow ? "" : "$") + (r + 1).ToString(CultureInfo.InvariantCulture);
    }
}

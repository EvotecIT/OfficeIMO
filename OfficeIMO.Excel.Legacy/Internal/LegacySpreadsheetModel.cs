using System;
using System.Collections.Generic;
using System.Threading;

namespace OfficeIMO.Excel.Legacy;

internal sealed class LegacySpreadsheetModel {
    internal List<LegacySpreadsheetSheet> Sheets { get; } = new();
    internal Dictionary<string, string> Metadata { get; } = new(StringComparer.OrdinalIgnoreCase);
    internal List<LegacySpreadsheetChartMetadata> Charts { get; } = new();
    internal List<LegacySpreadsheetName> Names { get; } = new();
    internal List<OfficeCompatibilityFinding> Findings { get; } = new();
    internal OfficeLegacyImportQuality Quality { get; set; } = OfficeLegacyImportQuality.Salvage;
    internal OfficeLegacyInertContentKind InertContent { get; set; }
    internal int RecoveredCellCount { get; set; }
}

internal sealed class LegacySpreadsheetSheet {
    internal LegacySpreadsheetSheet(string name) => Name = name;
    internal string Name { get; }
    internal List<LegacySpreadsheetCell> Cells { get; } = new();
}

internal sealed class LegacySpreadsheetCell {
    internal LegacySpreadsheetCell(int row, int column, object? value, string? formula = null, byte sourceFormat = 0, string? numberFormat = null, string? comment = null, OfficeIMO.Excel.ExcelHorizontalAlignment? alignment = null) {
        Row = row;
        Column = column;
        Value = value;
        Formula = formula;
        SourceFormat = sourceFormat;
        NumberFormat = numberFormat;
        Comment = comment;
        Alignment = alignment;
    }
    internal int Row { get; }
    internal int Column { get; }
    internal object? Value { get; }
    internal string? Formula { get; }
    internal byte SourceFormat { get; }
    internal string? NumberFormat { get; }
    internal string? Comment { get; }
    internal OfficeIMO.Excel.ExcelHorizontalAlignment? Alignment { get; }
}

internal sealed class LegacySpreadsheetName {
    internal LegacySpreadsheetName(string name, string sheetName, int firstRow, int firstColumn, int lastRow, int lastColumn) {
        Name = name; SheetName = sheetName; FirstRow = firstRow; FirstColumn = firstColumn; LastRow = lastRow; LastColumn = lastColumn;
    }
    internal string Name { get; }
    internal string? ProjectedName { get; set; }
    internal string SheetName { get; }
    internal int FirstRow { get; }
    internal int FirstColumn { get; }
    internal int LastRow { get; }
    internal int LastColumn { get; }
}

internal interface ILegacySpreadsheetAdapter {
    LegacySpreadsheetFormat Format { get; }
    string ProfileId { get; }
    string GetProfileId(byte[] data, OfficeLegacyImportLimits limits, CancellationToken cancellationToken);
    int Probe(byte[] data, string? sourceName, OfficeLegacyImportLimits limits, CancellationToken cancellationToken, out string reason);
    LegacySpreadsheetModel Parse(byte[] data, OfficeLegacyImportLimits limits, CancellationToken cancellationToken);
}

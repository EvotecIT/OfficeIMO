using System;
using System.Collections.Generic;
using OfficeIMO.Excel;

namespace OfficeIMO.Excel.Legacy;

/// <summary>Legacy spreadsheet families recognized by the managed importer.</summary>
public enum LegacySpreadsheetFormat {
    /// <summary>Selected Lotus 1-2-3 WK/123 profiles.</summary>
    Lotus123,
    /// <summary>Selected Borland/Corel Quattro Pro WQ/WB/QPW profiles.</summary>
    QuattroPro,
    /// <summary>Selected Microsoft Multiplan DOS profiles.</summary>
    Multiplan,
    /// <summary>Selected Microsoft Works spreadsheet profiles.</summary>
    MicrosoftWorks
}

/// <summary>Describes one bounded legacy-spreadsheet source profile match.</summary>
public sealed class LegacySpreadsheetDetection {
    internal LegacySpreadsheetDetection(LegacySpreadsheetFormat format, string profileId, int confidence, string reason) {
        Format = format;
        ProfileId = profileId;
        Confidence = confidence;
        Reason = reason;
    }
    /// <summary>Gets the detected product family.</summary>
    public LegacySpreadsheetFormat Format { get; }
    /// <summary>Gets the stable adapter/profile identifier.</summary>
    public string ProfileId { get; }
    /// <summary>Gets confidence from 0 through 100.</summary>
    public int Confidence { get; }
    /// <summary>Gets the bounded evidence used for detection.</summary>
    public string Reason { get; }
}

/// <summary>Options for safe read-only legacy-spreadsheet import.</summary>
public sealed class LegacySpreadsheetImportOptions {
    /// <summary>Gets or sets hard resource limits.</summary>
    public OfficeLegacyImportLimits Limits { get; set; } = new();
    /// <summary>Gets or sets an explicit family for a known weak or damaged source.</summary>
    public LegacySpreadsheetFormat? FormatHint { get; set; }
    /// <summary>Gets or sets the source name used for extension-assisted detection.</summary>
    public string? SourceName { get; set; }
    /// <summary>Gets or sets whether salvage-quality output must be rejected.</summary>
    public bool RequireStructured { get; set; }
}

/// <summary>Recovered legacy chart metadata that was not converted into a misleading live chart.</summary>
public sealed class LegacySpreadsheetChartMetadata {
    internal LegacySpreadsheetChartMetadata(string sheetName, string sourceRecord, int payloadLength) {
        SheetName = sheetName;
        SourceRecord = sourceRecord;
        PayloadLength = payloadLength;
    }
    /// <summary>Gets the projected worksheet name.</summary>
    public string SheetName { get; }
    /// <summary>Gets the source record identifier.</summary>
    public string SourceRecord { get; }
    /// <summary>Gets the bounded source payload length.</summary>
    public int PayloadLength { get; }
}

/// <summary>Describes a recovered source cell, including its cached value and translated formula.</summary>
public sealed class LegacySpreadsheetCellContent {
    internal LegacySpreadsheetCellContent(string sheetName, LegacySpreadsheetCell source) {
        SheetName = sheetName; Row = source.Row; Column = source.Column; CachedValue = source.Value; Formula = source.Formula;
        SourceFormat = source.SourceFormat; NumberFormat = source.NumberFormat; Comment = source.Comment;
        Alignment = source.Alignment;
    }
    /// <summary>Gets the projected sheet name.</summary>
    public string SheetName { get; }
    /// <summary>Gets the 1-based row.</summary>
    public int Row { get; }
    /// <summary>Gets the 1-based column.</summary>
    public int Column { get; }
    /// <summary>Gets the finite cached value retained from the source.</summary>
    public object? CachedValue { get; }
    /// <summary>Gets the safely translated formula, or null when cached-value fallback was required.</summary>
    public string? Formula { get; }
    /// <summary>Gets the original format byte.</summary>
    public byte SourceFormat { get; }
    /// <summary>Gets the mapped Excel number format.</summary>
    public string? NumberFormat { get; }
    /// <summary>Gets a recovered comment.</summary>
    public string? Comment { get; }
    /// <summary>Gets the recovered horizontal alignment.</summary>
    public ExcelHorizontalAlignment? Alignment { get; }
}

/// <summary>Describes a safely recovered source named range.</summary>
public sealed class LegacySpreadsheetNameContent {
    internal LegacySpreadsheetNameContent(LegacySpreadsheetName source) {
        Name = source.Name; SheetName = source.SheetName; FirstRow = source.FirstRow; FirstColumn = source.FirstColumn; LastRow = source.LastRow; LastColumn = source.LastColumn;
        ProjectedName = source.ProjectedName;
    }
    /// <summary>Gets the source name.</summary>
    public string Name { get; }
    /// <summary>Gets the exact workbook name that was projected, or null when validation or collision handling retained the source name as metadata only.</summary>
    public string? ProjectedName { get; }
    /// <summary>Gets the projected sheet name.</summary>
    public string SheetName { get; }
    /// <summary>Gets the first 1-based row.</summary>
    public int FirstRow { get; }
    /// <summary>Gets the first 1-based column.</summary>
    public int FirstColumn { get; }
    /// <summary>Gets the last 1-based row.</summary>
    public int LastRow { get; }
    /// <summary>Gets the last 1-based column.</summary>
    public int LastColumn { get; }
}

/// <summary>Owns an imported editable workbook and its source-loss report.</summary>
public sealed class LegacySpreadsheetImportResult : IDisposable {
    internal LegacySpreadsheetImportResult(ExcelDocument document, LegacySpreadsheetDetection detection, OfficeLegacyImportReport report,
        IReadOnlyDictionary<string, string> metadata, IReadOnlyList<LegacySpreadsheetChartMetadata> charts,
        IReadOnlyList<LegacySpreadsheetCellContent> cells, IReadOnlyList<LegacySpreadsheetNameContent> names) {
        Document = document;
        Detection = detection;
        Report = report;
        Metadata = metadata;
        Charts = charts;
        Cells = cells;
        Names = names;
    }
    /// <summary>Gets the normal OfficeIMO workbook used by XLSX and converter packages.</summary>
    public ExcelDocument Document { get; }
    /// <summary>Gets detected family and profile information.</summary>
    public LegacySpreadsheetDetection Detection { get; }
    /// <summary>Gets structured/salvage quality, inert-content flags, and explicit losses.</summary>
    public OfficeLegacyImportReport Report { get; }
    /// <summary>Gets recovered workbook metadata and names that could not be safely projected.</summary>
    public IReadOnlyDictionary<string, string> Metadata { get; }
    /// <summary>Gets bounded chart metadata discovered in the source.</summary>
    public IReadOnlyList<LegacySpreadsheetChartMetadata> Charts { get; }
    /// <summary>Gets recovered source cells, including cached values retained beside translated formulas.</summary>
    public IReadOnlyList<LegacySpreadsheetCellContent> Cells { get; }
    /// <summary>Gets recovered source names and whether each was projected into the workbook.</summary>
    public IReadOnlyList<LegacySpreadsheetNameContent> Names { get; }
    /// <inheritdoc />
    public void Dispose() => Document.Dispose();
}

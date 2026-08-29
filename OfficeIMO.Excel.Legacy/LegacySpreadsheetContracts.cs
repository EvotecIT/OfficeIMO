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

/// <summary>Owns an imported editable workbook and its source-loss report.</summary>
public sealed class LegacySpreadsheetImportResult : IDisposable {
    internal LegacySpreadsheetImportResult(ExcelDocument document, LegacySpreadsheetDetection detection, OfficeLegacyImportReport report,
        IReadOnlyDictionary<string, string> metadata, IReadOnlyList<LegacySpreadsheetChartMetadata> charts) {
        Document = document;
        Detection = detection;
        Report = report;
        Metadata = metadata;
        Charts = charts;
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
    /// <inheritdoc />
    public void Dispose() => Document.Dispose();
}

namespace OfficeIMO.Reader.Excel;

/// <summary>Adds Excel workbook support to a modular Reader builder.</summary>
public static class OfficeDocumentReaderBuilderExcelExtensions {
    /// <summary>Stable Excel handler identifier.</summary>
    public const string HandlerId = "officeimo.reader.excel";
    /// <summary>Stable legacy-spreadsheet handler identifier.</summary>
    public const string LegacyHandlerId = "officeimo.reader.excel.legacy";

    /// <summary>Adds every Excel format classified by <see cref="global::OfficeIMO.Excel.ExcelFormatCatalog"/>.</summary>
    public static OfficeDocumentReaderBuilder AddExcelHandler(
        this OfficeDocumentReaderBuilder builder,
        ReaderExcelOptions? options = null,
        bool replaceExisting = false) {
        if (builder == null) throw new ArgumentNullException(nameof(builder));
        ReaderExcelOptions configured = ExcelReaderAdapter.Clone(options);
        return builder.AddHandler(new ReaderHandlerRegistration {
            Origin = ReaderHandlerOrigin.OfficeIMO,
            Id = HandlerId,
            DisplayName = "Excel Reader",
            Description = "OfficeIMO.Excel workbook projection with bounded row and table extraction.",
            Kind = ReaderInputKind.Excel,
            Extensions = global::OfficeIMO.Excel.ExcelFormatCatalog.All.Select(format => format.Extension).ToArray(),
            ReadDocumentPath = (path, readerOptions, token) => ExcelReaderAdapter.ReadDocument(path, readerOptions, configured, token),
            ReadDocumentStream = (stream, sourceName, readerOptions, token) => ExcelReaderAdapter.ReadDocument(stream, sourceName, readerOptions, configured, token),
            ProbeStream = (stream, sourceName, readerOptions, token) => ExcelReaderAdapter.ProbeEncryptedOpenXml(stream, sourceName, readerOptions, token),
            WarningBehavior = ReaderWarningBehavior.Mixed,
            DeterministicOutput = true
        }, replaceExisting);
    }

    /// <summary>Adds safe read-only handlers for selected legacy spreadsheet families.</summary>
    public static OfficeDocumentReaderBuilder AddLegacySpreadsheetHandler(
        this OfficeDocumentReaderBuilder builder,
        global::OfficeIMO.Excel.Legacy.LegacySpreadsheetImportOptions? importOptions = null,
        ReaderExcelOptions? options = null,
        bool replaceExisting = false) {
        if (builder == null) throw new ArgumentNullException(nameof(builder));
        ReaderExcelOptions configured = ExcelReaderAdapter.Clone(options);
        global::OfficeIMO.Excel.Legacy.LegacySpreadsheetImportOptions? configuredImport = LegacySpreadsheetReaderAdapter.Clone(importOptions);
        return builder.AddHandler(new ReaderHandlerRegistration {
            Origin = ReaderHandlerOrigin.OfficeIMO,
            Id = LegacyHandlerId,
            DisplayName = "Legacy Spreadsheet Reader",
            Description = "Bounded Lotus 1-2-3, Quattro Pro, Multiplan, and Works spreadsheet import.",
            Kind = ReaderInputKind.Excel,
            UseDetectedKindFallback = false,
            Extensions = new[] { ".wk1", ".wk2", ".wk3", ".wk4", ".123", ".wq1", ".wq2", ".wb1", ".wb2", ".wb3", ".qpw", ".mp", ".mp1", ".mp2", ".mp3", ".wks", ".xlr" },
            ReadDocumentPath = (path, readerOptions, token) => LegacySpreadsheetReaderAdapter.ReadDocument(path, readerOptions, configured, configuredImport, token),
            ReadDocumentStream = (stream, sourceName, readerOptions, token) => LegacySpreadsheetReaderAdapter.ReadDocument(stream, sourceName, readerOptions, configured, configuredImport, token),
            WarningBehavior = ReaderWarningBehavior.Mixed,
            DeterministicOutput = true,
            DefaultMaxInputBytes = 64L * 1024L * 1024L
        }, replaceExisting);
    }
}

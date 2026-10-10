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
            FormatQualifications = global::OfficeIMO.Excel.ExcelFormatCatalog.All
                .Select(format => new ReaderFormatQualification(format.Extension, format.Id, ReaderFormatSupport.ReadConvert,
                    profile: format.Encoding.ToString(),
                    preservation: new[] { "Readable text and supported structured content" },
                    limitations: new[] { "Semantic extraction; source package, layout, macros and signatures are not reproduced", "Configured projection limits and owner compatibility boundaries apply" },
                    evidence: new[] { "OfficeIMO.Reader.Excel/README.md", "OfficeIMO.Reader.Tests/Reader.DocumentReadResult.cs" })).ToArray(),
            Extensions = global::OfficeIMO.Excel.ExcelFormatCatalog.All.Select(format => format.Extension).ToArray(),
            SupportsIncrementalPath = true,
            ReadPath = (path, readerOptions, token) => ExcelReaderAdapter.ReadIncremental(path, readerOptions, configured, token),
            ReadDocumentPath = (path, readerOptions, token) => ExcelReaderAdapter.ReadDocument(path, readerOptions, configured, token),
            ReadDocumentStream = (stream, sourceName, readerOptions, token) => ExcelReaderAdapter.ReadDocument(stream, sourceName, readerOptions, configured, token),
            ProbeStream = (stream, sourceName, readerOptions, token) => ExcelReaderAdapter.ProbeEncryptedOpenXml(stream, sourceName, readerOptions, token),
            WarningBehavior = ReaderWarningBehavior.Mixed,
            DeterministicOutput = true
        }, replaceExisting);
    }

    /// <summary>Adds the normal and legacy spreadsheet handlers with one immutable option set for every legacy route.</summary>
    public static OfficeDocumentReaderBuilder AddExcelAndLegacyHandlers(
        this OfficeDocumentReaderBuilder builder,
        global::OfficeIMO.Excel.Legacy.LegacySpreadsheetImportOptions? legacyImportOptions = null,
        ReaderExcelOptions? options = null,
        bool replaceExisting = false) {
        AddExcelHandler(builder, options, replaceExisting);
        return AddLegacySpreadsheetHandler(builder, legacyImportOptions, options, replaceExisting);
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
        long legacyMaxInputBytes = configuredImport?.Limits.MaxInputBytes ?? new global::OfficeIMO.OfficeLegacyImportLimits().MaxInputBytes;
        return builder.AddHandler(new ReaderHandlerRegistration {
            Origin = ReaderHandlerOrigin.OfficeIMO,
            Id = LegacyHandlerId,
            DisplayName = "Legacy Spreadsheet Reader",
            Description = "Bounded Lotus 1-2-3, Quattro Pro, Multiplan, Works, SYLK and DIF spreadsheet import.",
            Kind = ReaderInputKind.Excel,
            UseDetectedKindFallback = false,
            Extensions = new[] { ".wk1", ".wk2", ".wk3", ".wk4", ".123", ".wq1", ".wq2", ".wb1", ".wb2", ".wb3", ".qpw", ".mp", ".mp1", ".mp2", ".mp3", ".wks", ".xlr", ".slk", ".dif" },
            FormatQualifications = new[] {
                new ReaderFormatQualification(".slk", "sylk-stored-values", ReaderFormatSupport.ReadConvert,
                    preservation: new[] { "Cell coordinates, explicit blanks, text, finite numbers, booleans and valid stored formula values" },
                    limitations: new[] { "Formula expressions and shared/matrix/table behavior and formatting are omitted with loss reports; error markers become literal text", "SYLK character escapes other than line breaks and the historical SCALC3 string dialect are rejected", "No source-format writing or source expression evaluation" },
                    evidence: new[] { "OfficeIMO.Excel/README.md#profile-coverage", "OfficeIMO.LegacyImport.Tests/TextSpreadsheetImportTests.cs", "OfficeIMO.LegacyImport.Tests/Fixtures/TextSpreadsheets/README.md" }),
                new ReaderFormatQualification(".dif", "dif-row-values", ReaderFormatSupport.ReadConvert,
                    preservation: new[] { "Source row order, text, finite numbers and booleans" },
                    limitations: new[] { "DIF contains stored values, without formulas or source formatting", "Error markers become literal text with loss reports", "Unsupported topics and dimension mismatches are reported; no source-format writing" },
                    evidence: new[] { "OfficeIMO.Excel/README.md#profile-coverage", "OfficeIMO.LegacyImport.Tests/TextSpreadsheetImportTests.cs", "OfficeIMO.LegacyImport.Tests/Fixtures/TextSpreadsheets/README.md" })
            },
            ReadDocumentPath = (path, readerOptions, token) => LegacySpreadsheetReaderAdapter.ReadDocument(path, readerOptions, configured, configuredImport, token),
            ReadDocumentStream = (stream, sourceName, readerOptions, token) => LegacySpreadsheetReaderAdapter.ReadDocument(stream, sourceName, readerOptions, configured, configuredImport, token),
            ExtensionValidationProbeStream = (stream, sourceName, readerOptions, token) => LegacySpreadsheetReaderAdapter.Probe(stream, sourceName, readerOptions, configuredImport, token),
            WarningBehavior = ReaderWarningBehavior.Mixed,
            DeterministicOutput = true,
            DefaultMaxInputBytes = legacyMaxInputBytes,
            MaxInputBytesCeiling = legacyMaxInputBytes
        }, replaceExisting);
    }
}

namespace OfficeIMO.Reader.Excel;

/// <summary>Adds Excel workbook support to a modular Reader builder.</summary>
public static class OfficeDocumentReaderBuilderExcelExtensions {
    /// <summary>Stable Excel handler identifier.</summary>
    public const string HandlerId = "officeimo.reader.excel";

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
}

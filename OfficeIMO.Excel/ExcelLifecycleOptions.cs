using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Drawing;
using OfficeIMO.Excel.Xlsb;

namespace OfficeIMO.Excel {
    /// <summary>Controls creation and persistence of an Excel workbook.</summary>
    public sealed class ExcelCreateOptions : DocumentCreateOptions {
        /// <summary>Controls the Open XML workbook type when no destination extension is available.</summary>
        public SpreadsheetDocumentType DocumentType { get; set; } = SpreadsheetDocumentType.Workbook;
    }

    /// <summary>Controls access, persistence, and package behavior when loading an Excel workbook.</summary>
    public sealed class ExcelLoadOptions : DocumentLoadOptions {
        /// <summary>
        /// Maximum workbook bytes buffered by load APIs. Default: 512 MiB. Set to null to disable this compatibility guard.
        /// </summary>
        public long? MaxInputBytes { get; set; } = 512L * 1024L * 1024L;

        /// <summary>Provides optional low-level Open XML package settings.</summary>
        public OpenSettings? OpenSettings { get; set; }

        /// <summary>Provides optional resource limits and reporting controls for XLSB sources.</summary>
        public XlsbImportOptions? XlsbImportOptions { get; set; }
    }

    /// <summary>Controls creation of a workbook from a template package.</summary>
    public sealed class ExcelTemplateCreateOptions : DocumentCreateOptions {
        /// <summary>Controls whether an existing destination is replaced.</summary>
        public bool Overwrite { get; set; } = true;

        /// <summary>Provides optional low-level Open XML package settings.</summary>
        public OpenSettings? OpenSettings { get; set; }
    }
}

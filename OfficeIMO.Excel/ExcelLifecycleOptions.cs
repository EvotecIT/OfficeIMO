using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Drawing;
using OfficeIMO.Excel.Xlsb;

namespace OfficeIMO.Excel {
    /// <summary>Controls creation and persistence of an Excel workbook.</summary>
    public sealed class ExcelCreateOptions : DocumentCreateOptions {
    }

    /// <summary>Controls access, persistence, and package behavior when loading an Excel workbook.</summary>
    public sealed class ExcelLoadOptions : DocumentLoadOptions {
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

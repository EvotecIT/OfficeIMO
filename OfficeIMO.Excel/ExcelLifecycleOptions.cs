using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Core;

namespace OfficeIMO.Excel {
    /// <summary>Controls creation and persistence of an Excel workbook.</summary>
    public sealed class ExcelCreateOptions : DocumentCreateOptions {
    }

    /// <summary>Controls access, persistence, and package behavior when loading an Excel workbook.</summary>
    public sealed class ExcelLoadOptions : DocumentLoadOptions {
        /// <summary>Provides optional low-level Open XML package settings.</summary>
        public OpenSettings? OpenSettings { get; set; }
    }

    /// <summary>Controls creation of a workbook from a template package.</summary>
    public sealed class ExcelTemplateCreateOptions : DocumentCreateOptions {
        /// <summary>Controls whether an existing destination is replaced.</summary>
        public bool Overwrite { get; set; } = true;

        /// <summary>Provides optional low-level Open XML package settings.</summary>
        public OpenSettings? OpenSettings { get; set; }
    }
}

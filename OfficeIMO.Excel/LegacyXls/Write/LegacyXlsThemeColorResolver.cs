using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Excel.Utilities;

namespace OfficeIMO.Excel.LegacyXls.Write {
    /// <summary>Thin BIFF adapter over the workbook's canonical SpreadsheetML color resolver.</summary>
    internal sealed class LegacyXlsThemeColorResolver {
        private readonly WorkbookPart? _workbookPart;

        private LegacyXlsThemeColorResolver(WorkbookPart? workbookPart) {
            _workbookPart = workbookPart;
        }

        internal static LegacyXlsThemeColorResolver Create(ExcelDocument document) {
            if (document == null) throw new ArgumentNullException(nameof(document));
            return new LegacyXlsThemeColorResolver(document.WorkbookPartRoot);
        }

        internal bool TryResolve(uint themeIndex, double? tint, out string? argb) {
            argb = ExcelThemeColorResolver.ResolveTheme(themeIndex, tint, _workbookPart);
            return argb != null;
        }
    }
}

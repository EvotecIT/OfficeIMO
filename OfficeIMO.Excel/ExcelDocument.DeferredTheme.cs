using System.Text;
using OfficeIMO.Core.Internal.OpenXml;

namespace OfficeIMO.Excel {
    /// <summary>Theme and document-type changes to a workbook that has not opened its SDK package.</summary>
    public partial class ExcelDocument {
        private const string ThemeRelationshipType = OfficeXmlNamespaces.TransitionalRelationshipTypePrefix + "theme";
        private const string VbaProjectRelationshipType = "http://schemas.microsoft.com/office/2006/relationships/vbaProject";

        /// <summary>
        /// Replaces the theme part of a workbook that has not opened its SDK package, which already has a theme and a
        /// stylesheet; false, having changed nothing, otherwise.
        /// </summary>
        private bool TrySetDeferredWorkbookThemeXml(string themeXml) {
            if (_deferredWorkbook is not { } workbook) return false;
            string? themePart;
            using (var source = new MemoryStream(_deferredPackageBytes!, writable: false))
            using (OfficeOpcPackage package = OfficeOpcPackage.Open(source, leaveOpen: true)) {
                themePart = package.GetRelatedPartName(workbook.WorkbookPartName, ThemeRelationshipType);
                string? stylesPart = package.GetRelatedPartName(workbook.WorkbookPartName, StylesRelationshipType);
                if (themePart == null || !package.PartExists(themePart) || stylesPart == null || !package.PartExists(stylesPart)) return false;
            }
            byte[] bytes;
            using (var input = new MemoryStream(_deferredPackageBytes!, writable: false))
            using (OfficeOpcEditablePackage editable = OfficeOpcEditablePackage.Open(input, leaveOpen: true)) {
                byte[] theme = Encoding.UTF8.GetBytes(themeXml);
                editable.ReplacePart(themePart, theme, theme.Length);
                using var output = new MemoryStream(_deferredPackageBytes!.Length + theme.Length);
                editable.Save(output);
                bytes = output.ToArray();
            }
            _deferredPackageBytes = bytes;
            ForgetDeferredTables();
            _deferredPackageWrittenHere = true;
            MarkPackageDirty();
            return true;
        }

        /// <summary>The theme part of a workbook that has not opened its SDK package, or null when it has none.</summary>
        private string? ReadDeferredThemeXml() {
            if (_deferredWorkbook is not { } workbook) return null;
            using var source = new MemoryStream(_deferredPackageBytes!, writable: false);
            using OfficeOpcPackage package = OfficeOpcPackage.Open(source, leaveOpen: true);
            string? themePart = package.GetRelatedPartName(workbook.WorkbookPartName, ThemeRelationshipType);
            if (themePart == null || !package.PartExists(themePart)) return null;
            using OfficePooledBytes content = package.ReadPart(themePart);
            return Encoding.UTF8.GetString(content.Array, 0, content.Length);
        }

        /// <summary>
        /// The document type a workbook that has not opened its SDK package has, from its main content type, and whether
        /// it holds a VBA project; false when the content type is not one the owned path knows.
        /// </summary>
        private bool TryReadDeferredDocumentType(out ExcelDocumentType type, out bool hasMacros) {
            type = ExcelDocumentType.Workbook;
            hasMacros = false;
            if (_deferredWorkbook is not { } workbook) return false;
            string? contentType = ReadDeferredMainContentType();
            ExcelDocumentType? known = null;
            foreach (ExcelDocumentType candidate in new[] { ExcelDocumentType.Workbook, ExcelDocumentType.Template,
                         ExcelDocumentType.MacroEnabledWorkbook, ExcelDocumentType.MacroEnabledTemplate, ExcelDocumentType.AddIn }) {
                if (string.Equals(CreatedMainContentType(candidate), contentType, StringComparison.OrdinalIgnoreCase)) known = candidate;
            }
            if (known == null) return false;
            type = known.Value;
            using var source = new MemoryStream(_deferredPackageBytes!, writable: false);
            using OfficeOpcPackage package = OfficeOpcPackage.Open(source, leaveOpen: true);
            hasMacros = package.GetRelatedPartName(workbook.WorkbookPartName, VbaProjectRelationshipType) != null;
            return true;
        }
    }
}

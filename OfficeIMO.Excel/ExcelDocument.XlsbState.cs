namespace OfficeIMO.Excel {
    public partial class ExcelDocument {
        private static void EnsureXlsbFileTargetSupported(string path) {
            if (!ExcelDocumentLoadRouting.HasXlsbExtension(path)) {
                return;
            }

            throw new NotSupportedException("Native XLSB saving is not available in this build. The target was rejected before writing so an XLSX package is never emitted with an .xlsb extension.");
        }

        private static void EnsureXlsbStreamTargetSupported(ExcelFileFormat format) {
            if (format != ExcelFileFormat.Xlsb) {
                return;
            }

            throw new NotSupportedException("Native XLSB stream saving is not available in this build.");
        }
    }
}

using OfficeIMO.Drawing;

namespace OfficeIMO.Excel {
    public partial class ExcelDocument {
        /// <summary>
        /// Reopens an in-memory binary-workbook projection through the normal read-only Open XML path.
        /// Projection itself requires a writable package, but callers that requested read-only access must
        /// receive the same package access and save guards as callers loading an Open XML workbook.
        /// </summary>
        private static ExcelDocument ReopenProjectedWorkbookReadOnly(ExcelDocument projectedDocument) {
            if (projectedDocument == null) throw new ArgumentNullException(nameof(projectedDocument));

            byte[] packageBytes;
            try {
                packageBytes = projectedDocument.ToBytes(ExcelFileFormat.Xlsx);
            } finally {
                projectedDocument.Dispose();
            }

            return LoadFromByteArray(
                packageBytes,
                new ExcelLoadOptions { AccessMode = DocumentAccessMode.ReadOnly },
                filePath: null);
        }
    }
}

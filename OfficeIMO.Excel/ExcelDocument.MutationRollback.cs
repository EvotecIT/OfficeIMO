using DocumentFormat.OpenXml.Packaging;

namespace OfficeIMO.Excel {
    public partial class ExcelDocument {
        /// <summary>The package a transaction returns to, which the document already holds, so nothing is copied for a rollback that rarely happens.</summary>
        internal sealed class SourcePackageRollback {
            internal SourcePackageRollback(byte[] package, Dictionary<uint, WorksheetPart> sheetParts) {
                Package = package;
                SheetParts = sheetParts;
            }

            internal byte[] Package { get; }
            internal Dictionary<uint, WorksheetPart> SheetParts { get; }
        }

        /// <summary>
        /// Returns the rollback for a workbook that still is its source package, or null when a snapshot is needed.
        /// A snapshot measures serialized parts against <paramref name="maximumCharacters"/>, so only a package far
        /// inside that budget skips it. Serialization can differ from the expanded size, but not by a factor of two.
        /// </summary>
        internal SourcePackageRollback? TryCaptureSourcePackageRollback(long maximumCharacters) {
            if (ReadSnapshotIdentity is not byte[] package || _pendingRowDeletes != null) return null;
            long? expandedSize = ExcelPackageEntrySplicer.MeasureExpandedSize(package);
            if (expandedSize == null || expandedSize.Value > maximumCharacters / 2) return null;
            return new SourcePackageRollback(package, CaptureActiveSheetPartsForReload());
        }

        internal void RestoreSourcePackage(SourcePackageRollback rollback, IReadOnlyList<ExcelDrawingHandle.ReloadSnapshot> drawingHandles) {
            _sheetPartsBeforePackageClose = rollback.SheetParts;
            _drawingHandlesBeforePackageClose = drawingHandles;
            try { _spreadSheetDocument.Dispose(); } catch { }
            ReloadFromBytes(rollback.Package);
        }
    }
}

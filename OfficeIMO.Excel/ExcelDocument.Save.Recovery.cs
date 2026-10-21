namespace OfficeIMO.Excel {
    public partial class ExcelDocument {
        /// <summary>Reopens a serialized workbook after a file-save failure without hiding a recovery failure.</summary>
        /// <remarks>
        /// The caller must rethrow the original save exception after successful recovery. Once a serialized
        /// snapshot exists, the package may already be closed and another writer must not be attempted.
        /// A failure can also occur after the destination was committed; recovery does not roll back that file.
        /// </remarks>
        private void RestoreFailedFileSave(byte[] packageBytes, Exception saveFailure, bool clearDirectDataSetCandidate = false) {
            try {
                if (clearDirectDataSetCandidate) ClearDirectDataSetSaveCandidate();
                ReloadFromBytes(packageBytes, simplePackageContentKnown: true);
            } catch (Exception recoveryFailure) {
                throw new AggregateException(
                    "The file save failed and the workbook could not be restored for further editing.",
                    saveFailure, recoveryFailure);
            }
        }
    }
}

using OfficeIMO.Core.Internal;

namespace OfficeIMO.Excel;

public partial class ExcelDocument {
    private static bool NeedsStreamingFinalization(ExcelWriterOptions options) =>
        options.DocumentProperties?.Count > 0 || options.SaveOptions != null;

    private static void FinalizeStreamingFile(string path, ExcelWriterOptions options) {
        ExcelSaveOptions? save = options.SaveOptions;
        CheckStreamingTemporaryBudget(path, save);
        if (save == null || !RequiresStreamingSavePass(save)) return;
        if (save.MaxInMemoryPackageBytes is long memoryLimit && new FileInfo(path).Length > memoryLimit)
            throw new InvalidOperationException("The requested native save policy exceeds MaxInMemoryPackageBytes.");

        // These explicitly requested policies use the existing save owner, including its validation,
        // formula, encryption and format rules. Default row ingestion never enters this DOM phase.
        string finalPath = OfficeFileCommit.CreateTemporaryPath(path);
        try {
            using (var document = Load(path))
            using (var destination = new FileStream(finalPath, FileMode.CreateNew, FileAccess.ReadWrite, FileShare.None)) {
                document.Save(destination, save);
            }
            CheckStreamingTemporaryBudget(finalPath, save);
            OfficeFileCommit.CommitTemporaryFileAtomically(finalPath, path);
        } finally {
            if (File.Exists(finalPath)) File.Delete(finalPath);
        }
    }

    private static void CheckStreamingTemporaryBudget(string path, ExcelSaveOptions? options) {
        if (options?.MaxTemporaryPackageBytes is long maximum && new FileInfo(path).Length > maximum)
            throw new InvalidOperationException("The completed package exceeds MaxTemporaryPackageBytes.");
    }

    private static bool RequiresStreamingSavePass(ExcelSaveOptions options) =>
        options.Format != ExcelFileFormat.Xlsx || options.ResolvePassword(null) != null || options.ValidateOpenXml
        || options.SafePreflight || options.SafeRepairDefinedNames || options.DisableFastPackageWriter
        || options.EvaluateFormulasBeforeSave || options.ClearCachedFormulaResultsBeforeSave
        || options.MarkFormulasDirtyBeforeSave || options.ForceFullCalculationOnOpen;
}

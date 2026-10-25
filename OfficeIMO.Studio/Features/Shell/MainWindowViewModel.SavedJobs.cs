namespace OfficeIMO.Studio.Features.Shell;

public sealed partial class MainWindowViewModel {
    /// <summary>The desktop host brings the shared workbench into the visible document's navigation.</summary>
    internal Action? SavedFolderJobReviewRequested { get; set; }

    private void InitializeSavedJobReview() {
        Jobs.ReviewSavedJob = async (path, fingerprint, token) => {
            if (!await ConversionWorkbench.BatchExport.ReviewSavedJobAsync(path, fingerprint, token))
                throw new InvalidOperationException(ConversionWorkbench.BatchExport.Status);
            ConversionWorkbench.IsFolderJobReviewVisible = true;
            ShowConversionWorkbenchCommand.Execute(null);
            SavedFolderJobReviewRequested?.Invoke();
        };
    }
}

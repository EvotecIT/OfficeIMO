using OfficeIMO.GoogleWorkspace;
using OfficeIMO.GoogleWorkspace.Drive;

namespace OfficeIMO.Word.GoogleDocs {
    public sealed partial class GoogleDocsExporter : IGoogleDocsExporter {
        private static async Task<GoogleDriveFile?> ApplyDrivePlacementAsync(
            GoogleDriveClient driveClient,
            string? fileId,
            GoogleDriveFileLocation location,
            TranslationReport report,
            CancellationToken cancellationToken) {
            if (string.IsNullOrWhiteSpace(fileId)) {
                return null;
            }

            if (!string.IsNullOrWhiteSpace(location.FolderId)) {
                return await driveClient.MoveFileAsync(fileId!, location.FolderId!, report, cancellationToken).ConfigureAwait(false);
            }

            return await driveClient.GetFileAsync(fileId!, report: report, cancellationToken: cancellationToken).ConfigureAwait(false);
        }

        private static async Task ValidateDrivePlacementAsync(
            GoogleDriveClient driveClient,
            GoogleDriveFileLocation location,
            TranslationReport report,
            CancellationToken cancellationToken) {
            if (string.IsNullOrWhiteSpace(location.FolderId)) return;
            await driveClient.ResolveFolderAsync(location.FolderId!, location.DriveId, report, cancellationToken).ConfigureAwait(false);
        }

        private static string? BuildDocumentWebViewLink(string? documentId) {
            return string.IsNullOrWhiteSpace(documentId)
                ? null
                : $"https://docs.google.com/document/d/{documentId}/edit";
        }
    }

}

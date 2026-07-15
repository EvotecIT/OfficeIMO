using OfficeIMO.GoogleWorkspace;
using OfficeIMO.GoogleWorkspace.Drive;

namespace OfficeIMO.Word.GoogleDocs {
    public sealed partial class GoogleDocsExporter {
        private static async Task ApplyCommentsAsync(
            WordDocument document,
            GoogleDriveClient drive,
            string fileId,
            GoogleDocsSaveOptions options,
            TranslationReport report,
            CancellationToken cancellationToken) {
            if (options.Comments != GoogleDocsCommentMode.UnanchoredDriveComments || document.Comments.Count == 0) return;
            foreach (WordComment comment in document.Comments.Where(comment => comment.ParentComment == null)) {
                GoogleDriveComment created = await drive.CreateCommentAsync(
                    fileId,
                    FormatComment(comment),
                    anchor: null,
                    report,
                    cancellationToken).ConfigureAwait(false);
                if (string.IsNullOrWhiteSpace(created.Id)) continue;
                foreach (WordComment reply in comment.Replies) {
                    await drive.CreateReplyAsync(fileId, created.Id!, FormatComment(reply), report: report, cancellationToken: cancellationToken).ConfigureAwait(false);
                }
                if (comment.IsResolved == true) {
                    await drive.CreateReplyAsync(fileId, created.Id!, string.Empty, action: "resolve", report: report, cancellationToken: cancellationToken).ConfigureAwait(false);
                }
            }
            report.Add(
                TranslationSeverity.Warning,
                "Comments",
                $"Created {document.Comments.Count} Word comment/reply item(s) through Drive. Google editors display them as unanchored discussions.",
                code: "DOCS.COMMENT.UNANCHORED_CREATED",
                action: TranslationAction.Flatten,
                count: document.Comments.Count,
                targetId: fileId);
        }

        private static string FormatComment(WordComment comment) {
            string prefix = string.IsNullOrWhiteSpace(comment.Author) ? string.Empty : comment.Author + ": ";
            return prefix + (comment.Text ?? string.Empty);
        }
    }
}

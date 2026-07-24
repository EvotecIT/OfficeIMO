using OfficeIMO.GoogleWorkspace;
using OfficeIMO.GoogleWorkspace.Drive;

namespace OfficeIMO.Word.GoogleDocs {
    public sealed partial class GoogleDocsExporter {
        private static async Task ApplyCommentsAsync(
            WordDocument document,
            GoogleDriveClient drive,
            string fileId,
            GoogleDocsSaveOptions options,
            bool reconcileExistingComments,
            TranslationReport report,
            CancellationToken cancellationToken) {
            if (options.Comments != GoogleDocsCommentMode.UnanchoredDriveComments) return;
            var existingComments = reconcileExistingComments
                ? await ListCommentsAsync(drive, fileId, report, cancellationToken).ConfigureAwait(false)
                : new List<GoogleDriveComment>();
            IReadOnlyList<WordComment> wordComments = document.Comments;
            CommentThreadEntry[] entries = wordComments
                .Select(comment => new CommentThreadEntry(comment, comment.ParaId,
                    comment.ParentParaId, comment.IsResolved))
                .ToArray();
            Dictionary<string, CommentThreadEntry[]> repliesByParent = entries
                .Where(entry => !string.IsNullOrWhiteSpace(entry.ParentParaId))
                .GroupBy(entry => entry.ParentParaId!, StringComparer.Ordinal)
                .ToDictionary(group => group.Key, group => group.ToArray(), StringComparer.Ordinal);
            var claimedReplyParents = new HashSet<string>(StringComparer.Ordinal);
            var usedCommentIds = new HashSet<string>(StringComparer.Ordinal);
            int createdCount = 0;
            int reusedCount = 0;
            foreach (CommentThreadEntry entry in entries.Where(entry => string.IsNullOrWhiteSpace(entry.ParentParaId))) {
                WordComment comment = entry.Comment;
                string rootContent = FormatComment(comment);
                GoogleDriveComment? target = existingComments.FirstOrDefault(candidate =>
                    !candidate.Deleted
                    && !string.IsNullOrWhiteSpace(candidate.Id)
                    && !usedCommentIds.Contains(candidate.Id!)
                    && string.Equals(candidate.Content, rootContent, StringComparison.Ordinal));
                if (target == null) {
                    target = await drive.CreateCommentAsync(
                        fileId,
                        rootContent,
                        anchor: null,
                        report,
                        cancellationToken).ConfigureAwait(false);
                    createdCount++;
                } else {
                    usedCommentIds.Add(target.Id!);
                    reusedCount++;
                }

                if (string.IsNullOrWhiteSpace(target.Id)) continue;
                var usedReplyIds = new HashSet<string>(StringComparer.Ordinal);
                IReadOnlyList<CommentThreadEntry> replies = !string.IsNullOrWhiteSpace(entry.ParaId)
                    && claimedReplyParents.Add(entry.ParaId!)
                    && repliesByParent.TryGetValue(entry.ParaId!, out CommentThreadEntry[]? groupedReplies)
                        ? groupedReplies
                        : Array.Empty<CommentThreadEntry>();
                foreach (CommentThreadEntry replyEntry in replies) {
                    WordComment reply = replyEntry.Comment;
                    string replyContent = FormatComment(reply);
                    GoogleDriveReply? existingReply = target.Replies.FirstOrDefault(candidate =>
                        !candidate.Deleted
                        && !string.IsNullOrWhiteSpace(candidate.Id)
                        && !usedReplyIds.Contains(candidate.Id!)
                        && string.Equals(candidate.Content, replyContent, StringComparison.Ordinal));
                    if (existingReply != null) {
                        usedReplyIds.Add(existingReply.Id!);
                        reusedCount++;
                        continue;
                    }

                    await drive.CreateReplyAsync(fileId, target.Id!, replyContent, report: report, cancellationToken: cancellationToken).ConfigureAwait(false);
                    createdCount++;
                }
                if (entry.IsResolved.HasValue && entry.IsResolved.Value != target.Resolved) {
                    await drive.CreateReplyAsync(
                        fileId,
                        target.Id!,
                        string.Empty,
                        action: entry.IsResolved.Value ? "resolve" : "reopen",
                        report: report,
                        cancellationToken: cancellationToken).ConfigureAwait(false);
                }
            }
            if (createdCount > 0) {
                report.Add(
                    TranslationSeverity.Warning,
                    "Comments",
                    $"Created {createdCount} Word comment/reply item(s) through Drive. Google editors display them as unanchored discussions.",
                    code: "DOCS.COMMENT.UNANCHORED_CREATED",
                    action: TranslationAction.Flatten,
                    count: createdCount,
                    targetId: fileId);
            }
            if (reusedCount > 0) {
                report.Add(
                    TranslationSeverity.Info,
                    "Comments",
                    $"Reused {reusedCount} matching Drive comment/reply item(s) while replacing the document.",
                    code: "DOCS.COMMENT.UNANCHORED_REUSED",
                    action: TranslationAction.Preserve,
                    count: reusedCount,
                    targetId: fileId);
            }
            // Drive comments do not expose a reliable OfficeIMO provenance marker. Unmatched
            // collaborator discussions are therefore never deleted during content replacement.
        }

        private readonly struct CommentThreadEntry {
            internal CommentThreadEntry(WordComment comment, string? paraId, string? parentParaId,
                bool? isResolved) {
                Comment = comment;
                ParaId = paraId;
                ParentParaId = parentParaId;
                IsResolved = isResolved;
            }

            internal WordComment Comment { get; }
            internal string? ParaId { get; }
            internal string? ParentParaId { get; }
            internal bool? IsResolved { get; }
        }

        private static async Task<List<GoogleDriveComment>> ListCommentsAsync(
            GoogleDriveClient drive,
            string fileId,
            TranslationReport report,
            CancellationToken cancellationToken) {
            var comments = new List<GoogleDriveComment>();
            string? pageToken = null;
            do {
                GoogleDriveCommentList page = await drive.ListCommentsAsync(fileId, pageToken, report, cancellationToken).ConfigureAwait(false);
                comments.AddRange(page.Comments);
                pageToken = page.NextPageToken;
            } while (!string.IsNullOrWhiteSpace(pageToken));
            return comments;
        }

        private static string FormatComment(WordComment comment) {
            string prefix = string.IsNullOrWhiteSpace(comment.Author) ? string.Empty : comment.Author + ": ";
            return prefix + (comment.Text ?? string.Empty);
        }
    }
}

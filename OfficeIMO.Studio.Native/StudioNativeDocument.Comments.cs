using OfficeIMO.Pdf;

namespace OfficeIMO.Studio.Native;

/// <summary>
/// Comment threads: each comment with its nested replies and review state, read from the same reply catalog the
/// desktop Studio's Comments panel shows. Replying and resolving are undoable edits through <see cref="EditAsync"/>.
/// </summary>
internal sealed partial class StudioNativeDocument {
    /// <summary>
    /// The document's comment threads in page order. Popups, links and form widgets aren't comments, so they are left
    /// out. Each thread keeps an id across revisions, so a host keeps its selection when a reply rewrites the page.
    /// </summary>
    internal async Task<StudioNativeComments> GetCommentsAsync(CancellationToken cancellationToken) {
        await EnterAsync(cancellationToken).ConfigureAwait(false);
        try {
            IReadOnlyList<PdfAnnotation> annotations = _workspace.AnnotationMetadata;
            PdfAnnotationReviewCatalog catalog = PdfAnnotationReviewCatalog.Build(annotations, cancellationToken: cancellationToken);
            IReadOnlyList<PdfPageLayoutInfo>? layouts = null;
            var threads = new List<StudioNativeCommentThread>();
            foreach (PdfAnnotationReviewThread thread in catalog.Threads) {
                cancellationToken.ThrowIfCancellationRequested();
                PdfAnnotation root = thread.Root.Annotation;
                if (root.Subtype is "Popup" or "Link" or "Widget" || root.ObjectNumber is not int number) continue;
                var entries = new List<StudioNativeComment>();
                Add(thread.Root, 0);
                StudioNativeRectangle? bounds = null;
                if (root.PageNumber is int page && page >= 1 && page <= _workspace.Pages.Count && root.Width != 0 && root.Height != 0) {
                    layouts ??= _workspace.CreateDocumentSnapshot().GetPageLayouts();
                    PdfSelectionQuad quad = layouts[page - 1].MapUserSpaceRectangleToVisual(root.X1, root.Y1, root.X2, root.Y2);
                    bounds = new StudioNativeRectangle(quad.Left, quad.Top, quad.Width, quad.Height);
                }
                PdfAnnotationReviewState state = root.Review?.StandardState ?? PdfAnnotationReviewState.None;
                threads.Add(new StudioNativeCommentThread(IdentityOf(number), number, root.PageNumber,
                    root.Subtype, bounds, StateName(state), state == PdfAnnotationReviewState.Completed, thread.IsOrphanedReply, entries));

                void Add(PdfAnnotationReviewEntry entry, int depth) {
                    PdfAnnotation annotation = entry.Annotation;
                    entries.Add(new StudioNativeComment(annotation.ObjectNumber, annotation.Title, annotation.Contents ?? annotation.RichContentsPlainText,
                        annotation.Modified, depth));
                    foreach (PdfAnnotationReviewEntry reply in entry.Replies) Add(reply, depth + 1);
                }
            }
            return new StudioNativeComments(StudioNativeAbi.Version, _workspace.Revision, _workspace.CanEditAnnotations, threads);
        } finally { _gate.Release(); }
    }

    private Task ReplyToComment(StudioNativeCommentReply reply, CancellationToken cancellationToken) {
        if (string.IsNullOrWhiteSpace(reply.Text)) throw new ArgumentException("Write the reply first.", nameof(reply));
        RequireComment(reply.ObjectNumber);
        StudioNativeColor color = reply.Color ?? new StudioNativeColor(1D, 0.8D, 0.1D);
        return _workspace.AddAnnotationReplyAsync(reply.ObjectNumber, reply.Text, reply.Author ?? string.Empty,
            global::OfficeIMO.OfficeColor.FromNormalizedRgb(color.Red, color.Green, color.Blue), cancellationToken);
    }

    private Task SetCommentState(StudioNativeCommentState change, CancellationToken cancellationToken) {
        RequireComment(change.ObjectNumber);
        PdfAnnotationReviewState state = change.State switch {
            "completed" => PdfAnnotationReviewState.Completed, "none" => PdfAnnotationReviewState.None,
            "accepted" => PdfAnnotationReviewState.Accepted, "rejected" => PdfAnnotationReviewState.Rejected,
            "cancelled" => PdfAnnotationReviewState.Cancelled,
            _ => throw new ArgumentException($"“{change.State}” isn't a review state.", nameof(change))
        };
        return _workspace.SetAnnotationReviewStateAsync(change.ObjectNumber, state, cancellationToken);
    }

    private void RequireComment(int objectNumber) {
        if (!_workspace.CanEditAnnotations) throw new NotSupportedException("This document's comments can't be changed.");
        if (!_workspace.AnnotationMetadata.Any(annotation => annotation.ObjectNumber == objectNumber))
            throw new ArgumentException("The comment does not exist.", nameof(objectNumber));
    }

    // Identities follow an annotation through rewrites; one read before identities were built falls back to its number.
    private string IdentityOf(int objectNumber) {
        try { return _workspace.GetAnnotationIdentity(objectNumber).ToString("N"); }
        catch (KeyNotFoundException) { return "object-" + objectNumber.ToString(System.Globalization.CultureInfo.InvariantCulture); }
    }

    private static string StateName(PdfAnnotationReviewState state) => state switch {
        PdfAnnotationReviewState.Completed => "completed", PdfAnnotationReviewState.Accepted => "accepted",
        PdfAnnotationReviewState.Rejected => "rejected", PdfAnnotationReviewState.Cancelled => "cancelled",
        PdfAnnotationReviewState.Marked => "marked", PdfAnnotationReviewState.Unmarked => "unmarked", _ => "none"
    };
}

/// <summary>A comment or one of its replies; <c>depth</c> is 0 for the comment and counts replies to replies.</summary>
internal sealed record StudioNativeComment(int? ObjectNumber, string? Author, string? Contents, string? Modified, int Depth);

/// <summary>A comment and its replies. <c>id</c> stays the same across revisions; <c>bounds</c> are in points of the page as displayed.</summary>
internal sealed record StudioNativeCommentThread(string Id, int ObjectNumber, int? Page, string Subtype, StudioNativeRectangle? Bounds,
    string State, bool Resolved, bool Orphaned, IReadOnlyList<StudioNativeComment> Entries);

internal sealed record StudioNativeComments(int AbiVersion, long Revision, bool CanReply, IReadOnlyList<StudioNativeCommentThread> Threads);

/// <summary>A reply to the comment with <c>objectNumber</c>, written by <c>author</c>.</summary>
internal sealed record StudioNativeCommentReply(int ObjectNumber, string? Text, string? Author = null, StudioNativeColor? Color = null);

/// <summary>A comment's review state: completed (resolved), none (open), accepted, rejected or cancelled.</summary>
internal sealed record StudioNativeCommentState(int ObjectNumber, string? State);

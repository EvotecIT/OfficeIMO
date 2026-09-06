using OfficeIMO.Pdf;
using OfficeIMO.Studio.Infrastructure.Localization;

namespace OfficeIMO.Studio.Features.Comments;

/// <summary>Presentation of a thread from the PDF engine's bounded review catalog.</summary>
public sealed class CommentThreadViewModel {
    internal CommentThreadViewModel(PdfAnnotationReviewThread thread, IStudioLocalizer localizer) {
        Annotation = thread.Root.Annotation;
        IsResolved = Annotation.Review?.StandardState == PdfAnnotationReviewState.Completed;
        State = localizer.Get("Comments.State." + (Annotation.Review?.StandardState?.ToString() ?? "None"));
        Label = localizer.Format("Comments.ThreadLabel", Annotation.PageNumber, Annotation.Title ?? localizer.Get("Comments.UnknownAuthor"), Contents);
        IsOrphaned = thread.IsOrphanedReply;
        var entries = new List<CommentEntryViewModel>();
        Add(thread.Root, 0);
        Entries = entries;
        void Add(PdfAnnotationReviewEntry entry, int depth) {
            entries.Add(new(entry.Annotation.Title ?? localizer.Get("Comments.UnknownAuthor"), entry.Annotation.Contents ?? string.Empty,
                depth == 0 ? string.Empty : localizer.Format("Comments.ReplyDepth", depth)));
            foreach (var reply in entry.Replies) Add(reply, depth + 1);
        }
    }

    internal PdfAnnotation Annotation { get; }
    public string Contents => Annotation.Contents ?? string.Empty;
    public string Label { get; }
    public string State { get; }
    public bool IsResolved { get; }
    public bool IsOrphaned { get; }
    public IReadOnlyList<CommentEntryViewModel> Entries { get; }

    internal bool Matches(CommentThreadViewModel other) {
        var a = Annotation;
        var b = other.Annotation;
        if (!string.IsNullOrEmpty(a.Name) && !string.IsNullOrEmpty(b.Name)) return a.Name == b.Name;
        return a.PageNumber == b.PageNumber && a.Subtype == b.Subtype && a.Title == b.Title && a.Contents == b.Contents
            && a.X1 == b.X1 && a.Y1 == b.Y1 && a.X2 == b.X2 && a.Y2 == b.Y2;
    }
}

/// <summary>A comment or nested reply displayed with its author and depth.</summary>
public sealed record CommentEntryViewModel(string Author, string Contents, string DepthLabel) {
    public bool IsReply => DepthLabel.Length != 0;
}

/// <summary>A localized comment filter choice.</summary>
public sealed record CommentStatusChoice(string Id, string Label);

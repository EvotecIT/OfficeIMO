namespace OfficeIMO.Pdf;

/// <summary>Typed counts for a PDF sanitization preview or verified result.</summary>
public sealed class PdfSanitizationCategoryCounts {
    internal PdfSanitizationCategoryCounts(
        int userMetadata,
        int embeddedFiles,
        int actions,
        int commentsAndMarkup,
        int bookmarks,
        int optionalContent) {
        UserMetadata = userMetadata;
        EmbeddedFiles = embeddedFiles;
        Actions = actions;
        CommentsAndMarkup = commentsAndMarkup;
        Bookmarks = bookmarks;
        OptionalContent = optionalContent;
    }

    /// <summary>Number of selected user-authored Info fields and XMP packets.</summary>
    public int UserMetadata { get; }
    /// <summary>Number of selected logical file attachments.</summary>
    public int EmbeddedFiles { get; }
    /// <summary>Number of selected action findings.</summary>
    public int Actions { get; }
    /// <summary>Number of selected comment and markup annotations.</summary>
    public int CommentsAndMarkup { get; }
    /// <summary>Number of selected outline/bookmark entries.</summary>
    public int Bookmarks { get; }
    /// <summary>Number of selected optional-content layer definitions.</summary>
    public int OptionalContent { get; }
    /// <summary>Total number of selected logical items.</summary>
    public int Total => UserMetadata + EmbeddedFiles + Actions + CommentsAndMarkup + Bookmarks + OptionalContent;

    /// <summary>Returns the count for one atomic content category.</summary>
    public int GetCount(PdfSanitizationContentKind kind) => kind switch {
        PdfSanitizationContentKind.None => 0,
        PdfSanitizationContentKind.UserMetadata => UserMetadata,
        PdfSanitizationContentKind.EmbeddedFiles => EmbeddedFiles,
        PdfSanitizationContentKind.Actions => Actions,
        PdfSanitizationContentKind.CommentsAndMarkup => CommentsAndMarkup,
        PdfSanitizationContentKind.Bookmarks => Bookmarks,
        PdfSanitizationContentKind.OptionalContent => OptionalContent,
        _ => throw new ArgumentOutOfRangeException(nameof(kind), kind, "A single PDF sanitization content category is required.")
    };
}

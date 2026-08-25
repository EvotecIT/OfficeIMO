namespace OfficeIMO.Pdf;

/// <summary>Options for adding a reply to an existing annotation.</summary>
public sealed class PdfAnnotationReplyOptions {
    /// <summary>Reply author stored in /T.</summary>
    public string? Author { get; set; }

    /// <summary>Reply subject stored in /Subj.</summary>
    public string? Subject { get; set; }

    /// <summary>Optional initial review state.</summary>
    public PdfAnnotationReviewState? ReviewState { get; set; }

    /// <summary>Optional 18-point reply icon rectangle. Defaults to the lower-left corner of the parent.</summary>
    public IReadOnlyList<double>? Rectangle { get; set; }

    /// <summary>Optional RGB annotation color.</summary>
    public IReadOnlyList<double>? Color { get; set; }

    /// <summary>Text annotation icon name.</summary>
    public string IconName { get; set; } = "Comment";

    /// <summary>Creates a linked popup for the reply.</summary>
    public bool CreatePopup { get; set; }

    /// <summary>Initial popup open state.</summary>
    public bool PopupOpen { get; set; }

    /// <summary>Preferred full-rewrite or append-only mutation mode.</summary>
    public PdfMutationExecutionPreference ExecutionPreference { get; set; } = PdfMutationExecutionPreference.Automatic;
}

/// <summary>Resource limits for building annotation reply-thread trees.</summary>
public sealed class PdfAnnotationReviewCatalogOptions {
    /// <summary>Maximum nested reply depth. Defaults to 128 and cannot exceed 512.</summary>
    public int MaximumThreadDepth { get; set; } = 128;

    /// <summary>Maximum /IRT relationships represented by one catalog.</summary>
    public int MaximumRelationships { get; set; } = 100_000;
}

/// <summary>One annotation and its nested replies in a review thread.</summary>
public sealed class PdfAnnotationReviewEntry {
    internal PdfAnnotationReviewEntry(PdfAnnotation annotation, IReadOnlyList<PdfAnnotationReviewEntry> replies) {
        Annotation = annotation;
        Replies = replies;
    }

    /// <summary>The annotation represented by this entry.</summary>
    public PdfAnnotation Annotation { get; }

    /// <summary>Replies whose /IRT points to this annotation.</summary>
    public IReadOnlyList<PdfAnnotationReviewEntry> Replies { get; }
}

/// <summary>A root annotation and all of its replies.</summary>
public sealed class PdfAnnotationReviewThread {
    internal PdfAnnotationReviewThread(PdfAnnotationReviewEntry root, bool isOrphanedReply) {
        Root = root;
        IsOrphanedReply = isOrphanedReply;
    }

    /// <summary>Root entry for the thread.</summary>
    public PdfAnnotationReviewEntry Root { get; }

    /// <summary>True when the root declares /IRT but its parent is absent or invalid.</summary>
    public bool IsOrphanedReply { get; }
}

/// <summary>Validated thread view over the annotations in a PDF artifact.</summary>
public sealed class PdfAnnotationReviewCatalog {
    private PdfAnnotationReviewCatalog(
        IReadOnlyList<PdfAnnotationReviewThread> threads,
        int annotationCount,
        int replyCount,
        int orphanedReplyCount) {
        Threads = threads;
        AnnotationCount = annotationCount;
        ReplyCount = replyCount;
        OrphanedReplyCount = orphanedReplyCount;
    }

    /// <summary>Root review threads in page and object order.</summary>
    public IReadOnlyList<PdfAnnotationReviewThread> Threads { get; }

    /// <summary>Total annotations represented by the catalog.</summary>
    public int AnnotationCount { get; }

    /// <summary>Annotations that declare an /IRT relationship.</summary>
    public int ReplyCount { get; }

    /// <summary>Replies whose parent could not be represented as a valid thread ancestor.</summary>
    public int OrphanedReplyCount { get; }

    /// <summary>Reads all annotations and builds their reply threads.</summary>
    public static PdfAnnotationReviewCatalog Read(
        byte[] pdf,
        PdfReadOptions? readOptions = null,
        PdfAnnotationReviewCatalogOptions? options = null) {
        Guard.NotNull(pdf, nameof(pdf));
        return Build(PdfInspector.Inspect(pdf, readOptions).Annotations, options);
    }

    /// <summary>Builds reply threads from already-read annotations.</summary>
    public static PdfAnnotationReviewCatalog Build(
        IReadOnlyList<PdfAnnotation> annotations,
        PdfAnnotationReviewCatalogOptions? options = null) {
        Guard.NotNull(annotations, nameof(annotations));
        PdfAnnotationReviewCatalogOptions effectiveOptions = options ?? new PdfAnnotationReviewCatalogOptions();
        if (effectiveOptions.MaximumThreadDepth <= 0 || effectiveOptions.MaximumThreadDepth > 512) {
            throw new ArgumentOutOfRangeException(nameof(options), "Maximum annotation thread depth must be between 1 and 512.");
        }
        if (effectiveOptions.MaximumRelationships <= 0) {
            throw new ArgumentOutOfRangeException(nameof(options), "Maximum annotation relationships must be positive.");
        }

        List<PdfAnnotation> ordered = annotations
            .Where(static annotation => annotation is not null)
            .OrderBy(static annotation => annotation.PageNumber ?? int.MaxValue)
            .ThenBy(static annotation => annotation.ObjectNumber ?? int.MaxValue)
            .ToList();
        Dictionary<int, PdfAnnotation> byObjectNumber = ordered
            .Where(static annotation => annotation.ObjectNumber.HasValue)
            .GroupBy(static annotation => annotation.ObjectNumber!.Value)
            .ToDictionary(static group => group.Key, static group => group.First());
        var children = new Dictionary<int, List<PdfAnnotation>>();
        foreach (PdfAnnotation annotation in ordered) {
            int? parentObjectNumber = annotation.Review?.InReplyToObjectNumber;
            if (!parentObjectNumber.HasValue || !byObjectNumber.ContainsKey(parentObjectNumber.Value) || annotation.ObjectNumber == parentObjectNumber) continue;
            if (!children.TryGetValue(parentObjectNumber.Value, out List<PdfAnnotation>? entries)) {
                entries = new List<PdfAnnotation>();
                children[parentObjectNumber.Value] = entries;
            }
            entries.Add(annotation);
        }
        int replyCount = ordered.Count(static annotation => annotation.Review?.InReplyToObjectNumber.HasValue == true);
        if (replyCount > effectiveOptions.MaximumRelationships) {
            throw new InvalidOperationException("Annotation review catalog exceeded the configured relationship limit.");
        }

        var visited = new HashSet<PdfAnnotation>();
        var threads = new List<PdfAnnotationReviewThread>();
        foreach (PdfAnnotation annotation in ordered) {
            int? parentObjectNumber = annotation.Review?.InReplyToObjectNumber;
            bool hasValidParent = parentObjectNumber.HasValue &&
                parentObjectNumber != annotation.ObjectNumber &&
                byObjectNumber.ContainsKey(parentObjectNumber.Value);
            if (hasValidParent) continue;
            threads.Add(new PdfAnnotationReviewThread(
                BuildEntry(annotation, children, visited, new HashSet<int>(), effectiveOptions.MaximumThreadDepth),
                parentObjectNumber.HasValue));
        }

        foreach (PdfAnnotation annotation in ordered) {
            if (visited.Contains(annotation)) continue;
            threads.Add(new PdfAnnotationReviewThread(
                BuildEntry(annotation, children, visited, new HashSet<int>(), effectiveOptions.MaximumThreadDepth),
                isOrphanedReply: true));
        }

        int orphanCount = threads.Count(static thread => thread.IsOrphanedReply);
        return new PdfAnnotationReviewCatalog(threads.AsReadOnly(), ordered.Count, replyCount, orphanCount);
    }

    private static PdfAnnotationReviewEntry BuildEntry(
        PdfAnnotation annotation,
        IReadOnlyDictionary<int, List<PdfAnnotation>> children,
        HashSet<PdfAnnotation> visited,
        HashSet<int> path,
        int maximumThreadDepth) {
        visited.Add(annotation);
        if (!annotation.ObjectNumber.HasValue || !path.Add(annotation.ObjectNumber.Value)) {
            return new PdfAnnotationReviewEntry(annotation, Array.Empty<PdfAnnotationReviewEntry>());
        }

        var replies = new List<PdfAnnotationReviewEntry>();
        if (children.TryGetValue(annotation.ObjectNumber.Value, out List<PdfAnnotation>? candidates)) {
            if (path.Count >= maximumThreadDepth && candidates.Count > 0) {
                throw new InvalidOperationException("Annotation review thread exceeded the configured depth limit.");
            }
            foreach (PdfAnnotation candidate in candidates) {
                if (candidate.ObjectNumber.HasValue && path.Contains(candidate.ObjectNumber.Value)) continue;
                replies.Add(BuildEntry(candidate, children, visited, new HashSet<int>(path), maximumThreadDepth));
            }
        }
        return new PdfAnnotationReviewEntry(annotation, replies.AsReadOnly());
    }
}

/// <summary>High-level annotation reply and review-state operations.</summary>
public static class PdfAnnotationReviewEditor {
    /// <summary>Adds a Text annotation reply whose /IRT points to an existing indirect annotation.</summary>
    public static PdfAnnotationEditResult AddReply(
        byte[] pdf,
        int parentObjectNumber,
        string contents,
        PdfAnnotationReplyOptions? options = null,
        PdfReadOptions? readOptions = null) {
        Guard.NotNull(pdf, nameof(pdf));
        Guard.PositiveInteger(parentObjectNumber, nameof(parentObjectNumber));
        if (string.IsNullOrWhiteSpace(contents)) throw new ArgumentException("Reply contents cannot be empty.", nameof(contents));

        PdfAnnotation? parent = PdfInspector.Inspect(pdf, readOptions).Annotations
            .FirstOrDefault(annotation => annotation.ObjectNumber == parentObjectNumber);
        if (parent is null || !parent.PageNumber.HasValue) {
            throw new ArgumentException("Reply parent annotation object was not found on a page.", nameof(parentObjectNumber));
        }

        PdfAnnotationReplyOptions effective = options ?? new PdfAnnotationReplyOptions();
        IReadOnlyList<double> rectangle = effective.Rectangle ?? new[] {
            parent.X1,
            parent.Y1,
            parent.X1 + 18D,
            parent.Y1 + 18D
        };
        return PdfAnnotationEditor.AddAnnotation(pdf, new PdfAnnotationCreateOptions {
            PageNumber = parent.PageNumber.Value,
            Subtype = "Text",
            Rectangle = rectangle,
            Contents = contents,
            Title = effective.Author,
            Subject = effective.Subject,
            Color = effective.Color,
            IconName = effective.IconName,
            InReplyToObjectNumber = parentObjectNumber,
            ReplyType = "R",
            ReviewState = effective.ReviewState,
            CreatePopup = effective.CreatePopup,
            PopupOpen = effective.PopupOpen,
            GenerateAppearance = false,
            ExecutionPreference = effective.ExecutionPreference
        }, readOptions);
    }

    /// <summary>Sets a standard review state on one indirect annotation and validates readback.</summary>
    public static PdfAnnotationEditResult SetState(
        byte[] pdf,
        int annotationObjectNumber,
        PdfAnnotationReviewState state,
        PdfMutationExecutionPreference executionPreference = PdfMutationExecutionPreference.Automatic,
        bool allowResidualDataInAppendOnly = false,
        PdfReadOptions? readOptions = null) {
        return PdfAnnotationEditor.UpdateAnnotation(pdf, annotationObjectNumber, new PdfAnnotationUpdateOptions {
            ReviewState = state,
            ExecutionPreference = executionPreference,
            AllowResidualDataInAppendOnly = allowResidualDataInAppendOnly
        }, readOptions);
    }
}

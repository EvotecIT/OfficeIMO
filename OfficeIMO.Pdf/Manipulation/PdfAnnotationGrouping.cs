namespace OfficeIMO.Pdf;

/// <summary>Resolves the standard PDF markup grouping relationship without treating comment replies as group members.</summary>
public static class PdfAnnotationGrouping {
    /// <summary>Returns an annotation and its same-page standard group. Ungrouped annotations are returned alone.</summary>
    /// <remarks>Malformed, nested, cross-page, or missing-parent relationships are not expanded.</remarks>
    public static IReadOnlyList<PdfAnnotation> GetMembers(IReadOnlyList<PdfAnnotation> annotations, int objectNumber) {
        Guard.NotNull(annotations, nameof(annotations));
        PdfAnnotation selected = annotations.SingleOrDefault(annotation => annotation.ObjectNumber == objectNumber)
            ?? throw new ArgumentException("The annotation was not found.", nameof(objectNumber));
        int primaryNumber = selected.Review is { IsGroup: true, InReplyToObjectNumber: int parent } ? parent : objectNumber;
        PdfAnnotation? primary = annotations.FirstOrDefault(annotation => annotation.ObjectNumber == primaryNumber);
        if (primary is null || primary.PageNumber != selected.PageNumber || primary.Review?.InReplyToObjectNumber is not null)
            return new[] { selected };
        return annotations.Where(annotation => annotation.PageNumber == primary.PageNumber &&
            (annotation.ObjectNumber == primaryNumber || annotation.Review is { IsGroup: true } review && review.InReplyToObjectNumber == primaryNumber)).ToArray();
    }
}

/// <summary>Changes a selected set's position within a page's annotation painting order.</summary>
public enum PdfAnnotationOrderChange {
    /// <summary>Moves selected annotations one position toward the front, preserving their relative order.</summary>
    Raise,
    /// <summary>Moves selected annotations one position toward the back, preserving their relative order.</summary>
    Lower,
    /// <summary>Moves selected annotations to the front, preserving their relative order.</summary>
    BringToFront,
    /// <summary>Moves selected annotations to the back, preserving their relative order.</summary>
    SendToBack
}

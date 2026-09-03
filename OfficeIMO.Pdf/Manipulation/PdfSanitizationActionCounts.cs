namespace OfficeIMO.Pdf;

/// <summary>Typed counts of action findings in a PDF sanitization preview or result.</summary>
public sealed class PdfSanitizationActionCounts {
    internal PdfSanitizationActionCounts(IReadOnlyList<PdfSanitizationFinding> findings) {
        for (int i = 0; i < findings.Count; i++) {
            PdfSanitizationActionKind? kind = findings[i].ActionKind;
            if (!kind.HasValue) continue;
            _counts.TryGetValue(kind.Value, out int count);
            _counts[kind.Value] = count + 1;
            Total++;
        }
    }

    private readonly Dictionary<PdfSanitizationActionKind, int> _counts = new();

    /// <summary>Total number of action findings.</summary>
    public int Total { get; private set; }

    /// <summary>Number of JavaScript action findings.</summary>
    public int JavaScript => GetCount(PdfSanitizationActionKind.JavaScript);
    /// <summary>Number of URI action or catalog URI-base findings.</summary>
    public int Uri => GetCount(PdfSanitizationActionKind.Uri);
    /// <summary>Number of Launch action findings.</summary>
    public int Launch => GetCount(PdfSanitizationActionKind.Launch);
    /// <summary>Number of SubmitForm action findings.</summary>
    public int SubmitForm => GetCount(PdfSanitizationActionKind.SubmitForm);
    /// <summary>Number of GoToR action findings.</summary>
    public int GoToR => GetCount(PdfSanitizationActionKind.GoToR);
    /// <summary>Number of GoToE action findings.</summary>
    public int GoToE => GetCount(PdfSanitizationActionKind.GoToE);
    /// <summary>Number of ImportData action findings.</summary>
    public int ImportData => GetCount(PdfSanitizationActionKind.ImportData);
    /// <summary>Number of Movie action findings.</summary>
    public int Movie => GetCount(PdfSanitizationActionKind.Movie);
    /// <summary>Number of Rendition action findings.</summary>
    public int Rendition => GetCount(PdfSanitizationActionKind.Rendition);
    /// <summary>Number of RichMedia action findings.</summary>
    public int RichMedia => GetCount(PdfSanitizationActionKind.RichMedia);

    /// <summary>Returns the count for one atomic action kind.</summary>
    public int GetCount(PdfSanitizationActionKind kind) => _counts.TryGetValue(kind, out int count) ? count : 0;
}

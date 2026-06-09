namespace OfficeIMO.Pdf;

/// <summary>
/// Represents one label/value row for generated document metadata, invoice facts, definition lists, and similar two-column PDF layouts.
/// </summary>
public sealed class PdfKeyValueRow {
    /// <summary>Creates a plain text label/value row.</summary>
    public PdfKeyValueRow(string? key, string? value)
        : this(new[] { TextRun.Normal(key ?? string.Empty) }, new[] { TextRun.Normal(value ?? string.Empty) }) {
    }

    /// <summary>Creates a rich text label/value row.</summary>
    public PdfKeyValueRow(IEnumerable<TextRun> keyRuns, IEnumerable<TextRun> valueRuns) {
        KeyRuns = SnapshotRuns(keyRuns, nameof(keyRuns));
        ValueRuns = SnapshotRuns(valueRuns, nameof(valueRuns));
    }

    /// <summary>Rich text runs used by the label cell.</summary>
    public IReadOnlyList<TextRun> KeyRuns { get; }

    /// <summary>Rich text runs used by the value cell.</summary>
    public IReadOnlyList<TextRun> ValueRuns { get; }

    /// <summary>Creates a plain text label/value row.</summary>
    public static PdfKeyValueRow Text(string? key, string? value) => new PdfKeyValueRow(key, value);

    /// <summary>Creates a rich text label/value row.</summary>
    public static PdfKeyValueRow Rich(IEnumerable<TextRun> keyRuns, IEnumerable<TextRun> valueRuns) => new PdfKeyValueRow(keyRuns, valueRuns);

    internal PdfTableCell[] ToTableCells() {
        return new[] {
            KeyRuns.Count == 0 ? PdfTableCell.TextCell(string.Empty) : PdfTableCell.RichTextCell(KeyRuns),
            ValueRuns.Count == 0 ? PdfTableCell.TextCell(string.Empty) : PdfTableCell.RichTextCell(ValueRuns)
        };
    }

    private static System.Collections.ObjectModel.ReadOnlyCollection<TextRun> SnapshotRuns(IEnumerable<TextRun> runs, string paramName) {
        Guard.NotNull(runs, paramName);
        var snapshot = new List<TextRun>();
        foreach (TextRun run in runs) {
            if (run == null) {
                throw new ArgumentException("PDF key/value row text runs cannot contain null entries.", paramName);
            }

            snapshot.Add(run);
        }

        return snapshot.AsReadOnly();
    }
}

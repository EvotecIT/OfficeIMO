namespace OfficeIMO.Pdf;

/// <summary>One printable label in a label-sheet recipe.</summary>
public sealed class PdfLabel {
    /// <summary>Creates a label with primary, secondary, and optional machine-readable text.</summary>
    public PdfLabel(string text, string? secondaryText = null, string? codeText = null) {
        Guard.NotNullOrWhiteSpace(text, nameof(text));
        Text = text;
        SecondaryText = secondaryText;
        CodeText = codeText;
    }

    /// <summary>Primary label text.</summary>
    public string Text { get; }
    /// <summary>Optional secondary line.</summary>
    public string? SecondaryText { get; }
    /// <summary>Optional searchable code value.</summary>
    public string? CodeText { get; }
    internal string ToDisplayText() => string.Join("\n", new[] { Text, SecondaryText, CodeText }.Where(static value => !string.IsNullOrWhiteSpace(value)));
}

/// <summary>A simple printable label grid implemented as a normal flow table.</summary>
public sealed class PdfLabelSheetComponent : IPdfComponent {
    private readonly PdfLabel[] _labels;

    /// <summary>Creates a label sheet with a fixed logical column count.</summary>
    public PdfLabelSheetComponent(IEnumerable<PdfLabel> labels, int columns = 3) {
        Guard.NotNull(labels, nameof(labels));
        if (columns < 1 || columns > 12) throw new ArgumentOutOfRangeException(nameof(columns));
        _labels = labels.ToArray();
        if (_labels.Length == 0) throw new ArgumentException("A label sheet requires at least one label.", nameof(labels));
        if (_labels.Any(static label => label == null)) throw new ArgumentException("Labels cannot contain null entries.", nameof(labels));
        Columns = columns;
    }

    /// <summary>Number of labels per row.</summary>
    public int Columns { get; }

    /// <inheritdoc />
    public void Compose(PdfItemCompose content) {
        Guard.NotNull(content, nameof(content));
        var rows = new List<string[]>();
        for (int index = 0; index < _labels.Length; index += Columns) {
            var row = new string[Columns];
            for (int column = 0; column < Columns; column++) {
                int labelIndex = index + column;
                row[column] = labelIndex < _labels.Length ? _labels[labelIndex].ToDisplayText() : string.Empty;
            }
            rows.Add(row);
        }
        content.Table(rows, style: new PdfTableStyle {
            HeaderRowCount = 0,
            RowStripeFill = null,
            CellPaddingX = 10,
            CellPaddingY = 10,
            MinRowHeight = 54
        });
    }
}

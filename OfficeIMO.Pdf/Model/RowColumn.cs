namespace OfficeIMO.Pdf;

internal sealed class RowColumn {
    private readonly System.Collections.Generic.List<IPdfBlock> _blocks = new();
    private readonly System.Collections.ObjectModel.ReadOnlyCollection<IPdfBlock> _blocksView;

    public double WidthPercent { get; private set; }
    public System.Collections.Generic.IReadOnlyList<IPdfBlock> Blocks => _blocksView;

    public RowColumn(double widthPercent) {
        ValidateWidth(widthPercent, nameof(widthPercent));
        WidthPercent = widthPercent;
        _blocksView = new System.Collections.ObjectModel.ReadOnlyCollection<IPdfBlock>(_blocks);
    }

    internal void AddBlock(IPdfBlock block) {
        Guard.NotNull(block, nameof(block));
        _blocks.Add(block);
    }

    internal void SetWidthPercent(double widthPercent) {
        ValidateWidth(widthPercent, nameof(widthPercent));
        WidthPercent = widthPercent;
    }

    internal static void ValidateWidth(double widthPercent, string paramName) {
        if (double.IsNaN(widthPercent) || double.IsInfinity(widthPercent)) {
            throw new System.ArgumentOutOfRangeException(paramName, widthPercent, "Column width must be a finite percentage.");
        }

        if (widthPercent <= 0) {
            throw new System.ArgumentOutOfRangeException(paramName, widthPercent, "Column width must be greater than 0%.");
        }

        if (widthPercent > 100) {
            throw new System.ArgumentOutOfRangeException(paramName, widthPercent, "Column width cannot exceed 100%.");
        }
    }
}


using System.Collections.ObjectModel;

namespace OfficeIMO.Html;

/// <summary>Semantic roles retained independently from paint operations.</summary>
public enum HtmlRenderSemanticGroupRole {
    /// <summary>Generic semantic division.</summary>
    Division,
    /// <summary>List container.</summary>
    List,
    /// <summary>One list item.</summary>
    ListItem,
    /// <summary>List marker or label.</summary>
    ListLabel,
    /// <summary>List item body.</summary>
    ListBody,
    /// <summary>Table container.</summary>
    Table,
    /// <summary>Table row.</summary>
    TableRow,
    /// <summary>Table header cell.</summary>
    TableHeaderCell,
    /// <summary>Table data cell.</summary>
    TableCell,
    /// <summary>Table or figure caption.</summary>
    Caption
}

/// <summary>Resolved scope of a semantic HTML table header.</summary>
public enum HtmlRenderTableHeaderScope {
    /// <summary>Header applies to its row or row group.</summary>
    Row,
    /// <summary>Header applies to its column or column group.</summary>
    Column,
    /// <summary>Header applies to both axes.</summary>
    Both
}

/// <summary>Paint-neutral semantic group retained by the shared HTML render model.</summary>
public sealed class HtmlRenderSemanticGroup : HtmlRenderVisual {
    private readonly ReadOnlyCollection<HtmlRenderVisual> _visuals;

    internal HtmlRenderSemanticGroup(
        HtmlRenderSemanticGroupRole role,
        double x,
        double y,
        double width,
        double height,
        IEnumerable<HtmlRenderVisual> visuals,
        int paintOrder,
        string? source,
        int columnSpan = 1,
        int rowSpan = 1,
        HtmlRenderTableHeaderScope? headerScope = null,
        double? layoutY = null)
        : base(HtmlRenderVisualKind.SemanticGroup, x, y, width, height, paintOrder, null, source, layoutY) {
        Role = role;
        ColumnSpan = columnSpan;
        RowSpan = rowSpan;
        HeaderScope = headerScope;
        _visuals = new List<HtmlRenderVisual>(visuals ?? throw new ArgumentNullException(nameof(visuals)))
            .OrderBy(item => item.PaintOrder)
            .ToList()
            .AsReadOnly();
    }

    /// <summary>Semantic role of this group.</summary>
    public HtmlRenderSemanticGroupRole Role { get; }

    /// <summary>Table column span, or one for non-cell groups.</summary>
    public int ColumnSpan { get; }

    /// <summary>Table row span, or one for non-cell groups.</summary>
    public int RowSpan { get; }

    /// <summary>Resolved table-header scope, or null for non-header groups.</summary>
    public HtmlRenderTableHeaderScope? HeaderScope { get; }

    /// <summary>Ordered child visuals.</summary>
    public IReadOnlyList<HtmlRenderVisual> Visuals => _visuals;

    internal override HtmlRenderVisual Translate(double offsetX, double offsetY, int paintOrder) =>
        new HtmlRenderSemanticGroup(Role, X + offsetX, Y + offsetY, Width, Height, _visuals.Select((visual, index) => visual.Translate(offsetX, offsetY, index)), paintOrder, Source, ColumnSpan, RowSpan, HeaderScope, LayoutY + offsetY);

    internal override HtmlRenderVisual TranslatePaint(double offsetX, double offsetY, int paintOrder) =>
        new HtmlRenderSemanticGroup(Role, X + offsetX, Y + offsetY, Width, Height, _visuals.Select((visual, index) => visual.TranslatePaint(offsetX, offsetY, index)), paintOrder, Source, ColumnSpan, RowSpan, HeaderScope, LayoutY);
}

using System.Collections.ObjectModel;

namespace OfficeIMO.Html;

/// <summary>Semantic roles retained independently from paint operations.</summary>
public enum HtmlRenderSemanticGroupRole {
    /// <summary>Document section or landmark.</summary>
    Section,
    /// <summary>Generic semantic division.</summary>
    Division,
    /// <summary>Paragraph containing one or more text fragments.</summary>
    Paragraph,
    /// <summary>Level-one heading containing one or more text fragments.</summary>
    Heading1,
    /// <summary>Level-two heading containing one or more text fragments.</summary>
    Heading2,
    /// <summary>Level-three heading containing one or more text fragments.</summary>
    Heading3,
    /// <summary>Level-four heading containing one or more text fragments.</summary>
    Heading4,
    /// <summary>Level-five heading containing one or more text fragments.</summary>
    Heading5,
    /// <summary>Level-six heading containing one or more text fragments.</summary>
    Heading6,
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
    Caption,
    /// <summary>Decorative content intentionally excluded from tagged-PDF structure.</summary>
    Artifact,
    /// <summary>Footnote content associated with a call in the document body.</summary>
    Footnote
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
        double? layoutY = null,
        string? structureElementKey = null)
        : base(HtmlRenderVisualKind.SemanticGroup, x, y, width, height, paintOrder, null, source, layoutY) {
        Role = role;
        StructureElementKey = structureElementKey;
        ColumnSpan = columnSpan;
        RowSpan = rowSpan;
        HeaderScope = headerScope;
        _visuals = OrderVisuals(visuals);
    }

    private HtmlRenderSemanticGroup(
        HtmlRenderSemanticGroupRole role,
        double x,
        double y,
        double width,
        double height,
        List<HtmlRenderVisual> orderedVisuals,
        int paintOrder,
        string? source,
        int columnSpan,
        int rowSpan,
        HtmlRenderTableHeaderScope? headerScope,
        double layoutY,
        string? structureElementKey)
        : base(HtmlRenderVisualKind.SemanticGroup, x, y, width, height, paintOrder, null, source, layoutY) {
        Role = role;
        StructureElementKey = structureElementKey;
        ColumnSpan = columnSpan;
        RowSpan = rowSpan;
        HeaderScope = headerScope;
        _visuals = orderedVisuals.AsReadOnly();
    }

    /// <summary>Semantic role of this group.</summary>
    public HtmlRenderSemanticGroupRole Role { get; }

    internal string? StructureElementKey { get; }

    /// <summary>Table column span, or one for non-cell groups.</summary>
    public int ColumnSpan { get; }

    /// <summary>Table row span, or one for non-cell groups.</summary>
    public int RowSpan { get; }

    /// <summary>Resolved table-header scope, or null for non-header groups.</summary>
    public HtmlRenderTableHeaderScope? HeaderScope { get; }

    /// <summary>Ordered child visuals.</summary>
    public IReadOnlyList<HtmlRenderVisual> Visuals => _visuals;

    internal override HtmlRenderVisual Translate(double offsetX, double offsetY, int paintOrder) =>
        new HtmlRenderSemanticGroup(Role, X + offsetX, Y + offsetY, Width, Height, TranslateVisuals(offsetX, offsetY, translatePaint: false), paintOrder, Source, ColumnSpan, RowSpan, HeaderScope, LayoutY + offsetY, StructureElementKey);

    internal override HtmlRenderVisual TranslatePaint(double offsetX, double offsetY, int paintOrder) =>
        new HtmlRenderSemanticGroup(Role, X + offsetX, Y + offsetY, Width, Height, TranslateVisuals(offsetX, offsetY, translatePaint: true), paintOrder, Source, ColumnSpan, RowSpan, HeaderScope, LayoutY, StructureElementKey);

    private static ReadOnlyCollection<HtmlRenderVisual> OrderVisuals(IEnumerable<HtmlRenderVisual> visuals) {
        if (visuals == null) throw new ArgumentNullException(nameof(visuals));
        var materialized = new List<HtmlRenderVisual>(visuals);
        bool ordered = true;
        for (int index = 1; index < materialized.Count; index++) {
            if (materialized[index - 1].PaintOrder <= materialized[index].PaintOrder) continue;
            ordered = false;
            break;
        }
        return (ordered
            ? materialized
            : materialized.OrderBy(item => item.PaintOrder).ToList()).AsReadOnly();
    }

    private List<HtmlRenderVisual> TranslateVisuals(double offsetX, double offsetY, bool translatePaint) {
        var translated = new List<HtmlRenderVisual>(_visuals.Count);
        for (int index = 0; index < _visuals.Count; index++) {
            translated.Add(translatePaint
                ? _visuals[index].TranslatePaint(offsetX, offsetY, index)
                : _visuals[index].Translate(offsetX, offsetY, index));
        }
        return translated;
    }
}

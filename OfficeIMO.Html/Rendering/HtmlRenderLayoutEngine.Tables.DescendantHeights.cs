using AngleSharp.Dom;

namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    private readonly HashSet<(IElement Element, string Source, string Property)> _reportedTableDescendantHeights = new();
    private bool _tableCellHeightBasisActive;

    /// <summary>Finds percentage constraints whose containing-height chain reaches the cell, after the effective physical cascade.</summary>
    private TableCellPercentageContent? PrepareTableCellPercentageContent(TableFormattingCell cell, double width,
        HtmlRenderBoxStyle cellStyle, HtmlRenderBoxStyle tableStyle, int rowSpan, int depth) {
        // A genuine auto-height table and auto-height cell leave descendant
        // percentages indefinite under CSS; that is not an approximation.
        bool eligibleBasis = tableStyle.ExplicitHeight.HasValue
            || cellStyle.ExplicitHeight.HasValue && cellStyle.TablePercentageHeight.Length == 0;
        if (!eligibleBasis) return null;
        var declarations = new List<(IElement Element, string Source, string Property, string Value)>();
        string? unsupported = cell.Nodes != null ? "percentage content in an anonymous table cell"
            : rowSpan != 1 ? "percentage content in a rowspan cell"
            : IsVerticalWritingMode(cellStyle.WritingMode) ? "vertical table-cell content" : null;
        bool initialBasisCompatible = unsupported == null;
        Visit(cell.Element, cellStyle, true, depth);
        if (declarations.Count == 0) return null;
        return new TableCellPercentageContent(cell, declarations, unsupported, initialBasisCompatible, _nextLogicalTextOrder,
            ResolveContainingBlockHeight(cellStyle));

        void Visit(IElement owner, HtmlRenderBoxStyle parent, bool dependsOnCell, int currentDepth) {
            EnsureDepth(currentDepth, owner);
            foreach (HtmlPseudoElementKind kind in new[] { HtmlPseudoElementKind.Before, HtmlPseudoElementKind.After }) {
                if (cell.Nodes != null && ReferenceEquals(owner, cell.Element)) break;
                if (!_generatedContent.TryGetContent(owner, kind, out HtmlGeneratedContent generated)
                    || generated.Fragments.Count == 0
                    || !_styleResolver.TryResolvePseudo(owner, kind, width, parent, out HtmlRenderBoxStyle pseudo)
                    || pseudo.Display == "none") continue;
                if (!HasOrdinaryTableCellHeightFormatting(pseudo)
                    || generated.Fragments.Any(fragment => fragment.Kind != HtmlGeneratedContentFragmentKind.Text)) {
                    unsupported ??= "specialized generated table-cell content";
                    initialBasisCompatible = false;
                }
                if (dependsOnCell && pseudo.Display is not ("inline" or "contents")) {
                    AddDeclarations(owner, DescribePseudoSource(owner, kind), pseudo, _styleResolver.GetBoxCascadeStyle(owner, pseudo, kind));
                }
            }
            IEnumerable<IElement> children = cell.Nodes != null && ReferenceEquals(owner, cell.Element)
                ? cell.Nodes.OfType<IElement>() : owner.Children;
            foreach (IElement child in children) {
                CheckCancellation();
                ChargeLayoutOperation(HtmlRenderStyleResolver.DescribeSource(child));
                if (ShouldSkipElement(child) || IsClosedDisclosureChild(child)) continue;
                HtmlRenderBoxStyle childStyle = _styleResolver.Resolve(child, width, parent);
                if (childStyle.Display == "none") continue;
                bool ordinaryFormatting = HasOrdinaryTableCellHeightFormatting(childStyle);
                bool ordinary = ordinaryFormatting
                    && !IsReplacedImageElement(child) && !IsFormControlElement(child.LocalName)
                    && child.LocalName is not ("svg" or "math" or "iframe");
                if (!ordinary) {
                    unsupported ??= "specialized formatting in table-cell content";
                    // The existing normal-flow image renderer consumes a definite
                    // parent basis. Other specialized contexts need their own
                    // allocation qualification, even when cell heights coincide.
                    if (!ordinaryFormatting || !IsReplacedImageElement(child)) initialBasisCompatible = false;
                }
                HtmlComputedStyle? cascade = _styleResolver.GetBoxCascadeStyle(child, childStyle);
                string height = cascade?.GetValue("height") ?? string.Empty;
                bool sizedBox = childStyle.Display is not ("inline" or "contents") || IsReplacedImageElement(child);
                if (dependsOnCell && sizedBox && cascade != null) {
                    AddDeclarations(child, HtmlRenderStyleResolver.DescribeSource(child), childStyle, cascade);
                }
                // An intervening auto-height block ends the percentage-height
                // chain. A fixed height starts an independent definite basis.
                bool childDepends = dependsOnCell && height.IndexOf('%') >= 0 && sizedBox
                    && _styleResolver.ResolveTablePercentageHeight(childStyle, 1D, height).HasValue;
                if (!sizedBox) {
                    childDepends = dependsOnCell;
                }
                Visit(child, childStyle, childDepends, currentDepth + 1);
            }
        }

        void AddDeclarations(IElement element, string source, HtmlRenderBoxStyle style, HtmlComputedStyle? cascade) {
            if (cascade == null) return;
            foreach (string property in new[] { "height", "min-height", "max-height" }) {
                string value = cascade.GetValue(property);
                if (value.IndexOf('%') >= 0 && _styleResolver.ResolveTablePercentageHeight(style, 1D, value).HasValue) {
                    declarations.Add((element, source, property, value));
                }
            }
        }
    }

    private static bool HasOrdinaryTableCellHeightFormatting(HtmlRenderBoxStyle style) =>
        style.Display is "block" or "inline-block" or "inline" or "contents"
        && !IsVerticalWritingMode(style.WritingMode)
        && style.FloatSide == "none" && style.Position is "static" or "relative"
        && style.ColumnCount == null && style.ColumnWidth == null && style.ContainerType == "normal";

    /// <summary>Non-replaced inlines and contents boxes do not establish a height containing block for cell descendants.</summary>
    private HtmlRenderBoxStyle ForwardTableCellHeightBasis(IElement element, HtmlRenderBoxStyle style, HtmlRenderBoxStyle parent) {
        if (!_tableCellHeightBasisActive || style.Display is not ("inline" or "contents")
            || !HasOrdinaryTableCellHeightFormatting(style) || IsReplacedImageElement(element)
            || IsFormControlElement(element.LocalName) || element.LocalName is "svg" or "math" or "iframe") return style;
        HtmlRenderBoxStyle forwarded = style.Clone();
        double? height = ResolveContainingBlockHeight(parent);
        forwarded.ExplicitHeight = height + (forwarded.BorderBox ? forwarded.VerticalInsets : 0D);
        return forwarded;
    }

    /// <summary>Resolves eligible cell content once against used height, without another row allocation or text-order advance.</summary>
    private void ResolveTableCellPercentageContent(IReadOnlyList<TableRowLayout> rows, double spacing, int depth, bool paintSeparateBorders) {
        for (int index = 0; index < rows.Count; index++) {
            foreach (TableCellLayout cell in rows[index].Cells) {
                TableCellPercentageContent? content = cell.PercentageContent;
                if (content == null) continue;
                double height = Math.Max(0D, GetSpanningHeight(rows, index, cell.RowSpan, spacing) - cell.Style.VerticalInsets);
                if (!content.Supported) {
                    // Keep already-correct absolute-cell percentage content,
                    // including the existing image path, without a false loss.
                    if (!content.InitialBasisCompatible || !content.InitialHeight.HasValue
                        || Math.Abs(content.InitialHeight.Value - height) > 0.0001D) Report(content, content.Unsupported!);
                    continue;
                }
                HtmlRenderBoxStyle finalStyle = cell.Style.Clone();
                finalStyle.ExplicitHeight = height + (finalStyle.BorderBox ? finalStyle.VerticalInsets : 0D);
                int nextOrder = _nextLogicalTextOrder;
                _nextLogicalTextOrder = content.LogicalTextOrderStart;
                try {
                    HtmlInlineLayout resolved = LayoutTableCellContent(content.Cell, Math.Max(1D, cell.Width - cell.Style.HorizontalInsets),
                        finalStyle, depth, paintSeparateBorders);
                    if (_nextLogicalTextOrder - content.LogicalTextOrderStart == content.LogicalTextOrderCount) cell.Inline = resolved;
                    else Report(content, "cell relayout changed logical text allocation");
                } finally {
                    _nextLogicalTextOrder = nextOrder;
                }
            }
        }

        void Report(TableCellPercentageContent content, string reason) {
            foreach (var declaration in content.Declarations) {
                if (!_reportedTableDescendantHeights.Add((declaration.Element, declaration.Source, declaration.Property))) continue;
                _diagnostics.Add(ComponentName, HtmlRenderDiagnosticCodes.TableValueUnsupported,
                    "Percentage-height cell content used an allocation outside the qualified ordinary single-row subset.",
                    HtmlDiagnosticSeverity.Warning, declaration.Source,
                    declaration.Property + "=" + declaration.Value + ";" + reason, OfficeConversionLossKind.Approximation);
            }
        }
    }

    private sealed class TableCellPercentageContent {
        internal TableCellPercentageContent(TableFormattingCell cell,
            IReadOnlyList<(IElement Element, string Source, string Property, string Value)> declarations, string? unsupported, bool initialBasisCompatible,
            int logicalTextOrderStart, double? initialHeight) {
            Cell = cell;
            Declarations = declarations;
            Unsupported = unsupported;
            InitialBasisCompatible = initialBasisCompatible;
            LogicalTextOrderStart = logicalTextOrderStart;
            InitialHeight = initialHeight;
        }
        internal TableFormattingCell Cell { get; }
        internal IReadOnlyList<(IElement Element, string Source, string Property, string Value)> Declarations { get; }
        internal string? Unsupported { get; }
        internal bool Supported => Unsupported == null;
        internal bool InitialBasisCompatible { get; }
        internal int LogicalTextOrderStart { get; }
        internal int LogicalTextOrderCount { get; set; }
        internal double? InitialHeight { get; }
    }
}

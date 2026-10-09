using AngleSharp.Dom;

namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    // Formatting ownership is independent of the accessibility table role.
    private readonly HashSet<string> _tableFormattingKeys = new HashSet<string>(StringComparer.Ordinal);
    private readonly HashSet<IElement> _reportedTableFormattingFallbacks = new HashSet<IElement>();
    private readonly HashSet<(IElement Element, HtmlPseudoElementKind Kind)> _reportedTableGeneratedContentOmissions = new();
    private readonly Dictionary<INode, int> _tableFormattingSourceIndexes = new();

    private static bool IsInternalTableDisplay(string display) => display is "table-cell" or "table-row"
        or "table-row-group" or "table-header-group" or "table-footer-group";

    /// <summary>Groups improper table children without changing the source DOM or its selector and inheritance relationships.</summary>
    private IEnumerable<TableFlowEntry> EnumerateTableFlowEntries(IElement owner, IEnumerable<INode> nodes,
        double width, HtmlRenderBoxStyle parentStyle) {
        var pending = new List<INode>();
        foreach (INode node in nodes) {
            CheckCancellation();
            if (node is IElement element) {
                if (ShouldSkipElement(element)) continue;
                HtmlRenderBoxStyle childStyle = _styleResolver.Resolve(element, width, parentStyle);
                if (childStyle.Display == "none") continue;
                if (IsInternalTableDisplay(childStyle.Display) && childStyle.FloatSide == "none"
                    && !ShouldExtractOutOfFlow(childStyle)) {
                    pending.Add(node);
                    continue;
                }
            } else if (pending.Count > 0 && node is IText text && string.IsNullOrWhiteSpace(text.Data)) {
                continue;
            } else if (node is not IText) {
                continue;
            }
            if (pending.Count > 0) {
                yield return new TableFlowEntry(pending.ToArray());
                pending.Clear();
            }
            yield return new TableFlowEntry(node);
        }
        if (pending.Count > 0) yield return new TableFlowEntry(pending.ToArray());
    }

    /// <summary>Builds the CSS table, row-group, row, and cell boxes over original source nodes; anonymous boxes inherit only formatting properties.</summary>
    private TableFormattingStructure BuildTableFormattingStructure(IElement table, double width,
        HtmlRenderBoxStyle tableStyle, int depth, IReadOnlyList<INode>? anonymousNodes = null) {
        var rows = new List<TableFormattingRow>();
        IElement? caption = null;
        IElement? headerGroup = null;
        IElement? footerGroup = null;
        if (anonymousNodes == null) ReportOmittedTableGeneratedContent(table, width, tableStyle);
        AppendRows(table, anonymousNodes ?? table.ChildNodes.ToArray(), tableStyle, null, null, depth);
        return new TableFormattingStructure(table, rows, caption, anonymousNodes != null, headerGroup, footerGroup);

        void AppendRows(IElement owner, IEnumerable<INode> nodes, HtmlRenderBoxStyle parentStyle,
            IElement? group, HtmlRenderBoxStyle? groupStyle, int currentDepth) {
            EnsureDepth(currentDepth, owner);
            var cells = new List<TableFormattingCell>();
            var content = new List<INode>();
            void FlushContent() {
                if (content.Count == 0) return;
                cells.Add(new TableFormattingCell(owner, parentStyle, content.ToArray(),
                    GetAnonymousTableStructureElementKey(owner, content[0], "cell")));
                content.Clear();
            }
            void FlushRow() {
                FlushContent();
                if (cells.Count == 0) return;
                IElement anchor = cells.Select(cell => cell.ContinuationElement).FirstOrDefault(element => element != null) ?? owner;
                rows.Add(new TableFormattingRow(anchor, CreateAnonymousBoxStyle(parentStyle, "table-row", "anonymous-table-row"),
                    group, groupStyle, cells.ToArray(), anonymous: true,
                    GetAnonymousTableStructureElementKey(owner, cells[0].SourceNode, "row")));
                cells.Clear();
            }
            foreach (INode node in nodes) {
                CheckCancellation();
                if (node is IText text) {
                    if (!string.IsNullOrWhiteSpace(text.Data)) content.Add(node);
                    continue;
                }
                if (node is not IElement child || ShouldSkipElement(child)) continue;
                HtmlRenderBoxStyle childStyle = _styleResolver.Resolve(child, width, parentStyle);
                if (childStyle.Display == "none") continue;
                if (childStyle.Display is "table-row-group" or "table-header-group" or "table-footer-group") {
                    FlushRow();
                    if (childStyle.Display == "table-header-group") headerGroup ??= child;
                    if (childStyle.Display == "table-footer-group") footerGroup ??= child;
                    ReportOmittedTableGeneratedContent(child, width, childStyle);
                    _layoutStyles[child] = childStyle;
                    AppendRows(child, child.ChildNodes, childStyle, child, childStyle, currentDepth + 1);
                } else if (childStyle.Display == "table-row") {
                    FlushRow();
                    ReportOmittedTableGeneratedContent(child, width, childStyle);
                    _layoutStyles[child] = childStyle;
                    rows.Add(new TableFormattingRow(child, childStyle, group, groupStyle,
                        BuildCells(child, childStyle, currentDepth + 1), anonymous: false));
                } else if (childStyle.Display == "table-cell") {
                    FlushContent();
                    cells.Add(new TableFormattingCell(child, parentStyle));
                } else if (childStyle.Display == "table-caption" && anonymousNodes == null && ReferenceEquals(owner, table)) {
                    FlushRow();
                    if (caption == null) caption = child;
                    else ReportTableFormattingFallback(child, "multiple table captions");
                } else if (childStyle.Display is "table-column" or "table-column-group") {
                    if (child.LocalName is not "col" and not "colgroup") ReportTableFormattingFallback(child, "CSS column boxes");
                } else if (childStyle.Display == "contents") {
                    ReportTableFormattingFallback(child, "display:contents within a table formatting boundary");
                    FlushRow();
                    AppendRows(child, child.ChildNodes, childStyle, group, groupStyle, currentDepth + 1);
                } else {
                    content.Add(node);
                }
            }
            FlushRow();
        }

        IReadOnlyList<TableFormattingCell> BuildCells(IElement row, HtmlRenderBoxStyle rowStyle, int currentDepth) {
            EnsureDepth(currentDepth, row);
            var cells = new List<TableFormattingCell>();
            var content = new List<INode>();
            void FlushContent() {
                if (content.Count == 0) return;
                cells.Add(new TableFormattingCell(row, rowStyle, content.ToArray(),
                    GetAnonymousTableStructureElementKey(row, content[0], "cell")));
                content.Clear();
            }
            foreach (INode node in row.ChildNodes) {
                CheckCancellation();
                if (node is IText text && string.IsNullOrWhiteSpace(text.Data)) continue;
                if (node is IElement child) {
                    if (ShouldSkipElement(child)) continue;
                    HtmlRenderBoxStyle childStyle = _styleResolver.Resolve(child, width, rowStyle);
                    if (childStyle.Display == "none") continue;
                    if (childStyle.Display == "table-cell") {
                        FlushContent();
                        cells.Add(new TableFormattingCell(child, rowStyle));
                        continue;
                    }
                } else if (node is not IText) continue;
                content.Add(node);
            }
            FlushContent();
            return cells;
        }
    }

    private HtmlRenderBoxStyle ResolveFormattingCellStyle(TableFormattingCell cell, double width) {
        if (cell.Nodes != null) return CreateAnonymousBoxStyle(cell.ParentStyle, "table-cell", "anonymous-table-cell");
        HtmlRenderBoxStyle style = _styleResolver.Resolve(cell.Element, width, cell.ParentStyle);
        return IsFormControlElement(cell.Element.LocalName) && !UsesButtonChildLayout(cell.Element)
            && !IsInputType(cell.Element, "image") ? CreateFormControlStyle(cell.Element, style) : style;
    }

    /// <summary>Identifies an anonymous formatting box by its original owner and first source-node position, stable across reflow.</summary>
    private string GetAnonymousTableStructureElementKey(IElement owner, INode firstNode, string kind) {
        if (!_tableFormattingSourceIndexes.TryGetValue(firstNode, out int sourceIndex)) {
            int index = 0;
            foreach (INode node in firstNode.Parent?.ChildNodes ?? owner.ChildNodes) {
                ChargeLayoutOperation(HtmlRenderStyleResolver.DescribeSource(owner));
                _tableFormattingSourceIndexes[node] = index++;
            }
            _tableFormattingSourceIndexes.TryGetValue(firstNode, out sourceIndex);
        }
        return GetTableStructureElementKey(owner) + ":anonymous-" + kind + ":" + sourceIndex.ToString(System.Globalization.CultureInfo.InvariantCulture);
    }

    /// <summary>Reports effective generated content omitted by the bounded table, row, and row-group formatter.</summary>
    private void ReportOmittedTableGeneratedContent(IElement element, double width, HtmlRenderBoxStyle parentStyle) {
        Report(HtmlPseudoElementKind.Before);
        Report(HtmlPseudoElementKind.After);

        void Report(HtmlPseudoElementKind kind) {
            if (!_generatedContent.TryGetContent(element, kind, out HtmlGeneratedContent content)
                || !content.Fragments.Any(fragment => fragment.Kind != HtmlGeneratedContentFragmentKind.Text || fragment.Value.Length > 0)
                || !_styleResolver.TryResolvePseudo(element, kind, width, parentStyle, out HtmlRenderBoxStyle style)
                || style.Display == "none"
                || !_reportedTableGeneratedContentOmissions.Add((element, kind))) return;
            _diagnostics.Add(ComponentName, HtmlRenderDiagnosticCodes.TableValueUnsupported,
                "Generated content on a table, row, or row-group box was omitted.", HtmlDiagnosticSeverity.Warning,
                DescribePseudoSource(element, kind), "generated content on " + parentStyle.Display,
                OfficeConversionLossKind.Omission);
        }
    }

    private void ReportTableFormattingFallback(IElement element, string detail) {
        if (!_reportedTableFormattingFallbacks.Add(element)) return;
        _diagnostics.Add(ComponentName, HtmlRenderDiagnosticCodes.TableValueUnsupported,
            "A CSS table formatting structure used a bounded fallback.", HtmlDiagnosticSeverity.Warning,
            HtmlRenderStyleResolver.DescribeSource(element), detail,
            detail == "multiple table captions" ? OfficeConversionLossKind.Omission : OfficeConversionLossKind.Approximation);
    }

    private bool IsTableFormattingGroup(HtmlRenderSemanticGroup group) =>
        group.Role == HtmlRenderSemanticGroupRole.Table
        || group.StructureElementKey != null && _tableFormattingKeys.Contains(group.StructureElementKey);

    private static bool HasTableSemantics(IElement element) => element.LocalName == "table"
        || HtmlAccessibilitySemantics.HasRole(element, "table") || HtmlAccessibilitySemantics.HasRole(element, "grid")
        || HtmlAccessibilitySemantics.HasRole(element, "treegrid");

    private static bool HasRowSemantics(IElement element) => element.LocalName == "tr" || HtmlAccessibilitySemantics.HasRole(element, "row");

    private static bool HasHeaderCellSemantics(IElement element) => element.LocalName == "th"
        || HtmlAccessibilitySemantics.HasRole(element, "columnheader") || HtmlAccessibilitySemantics.HasRole(element, "rowheader");

    private static bool HasCellSemantics(IElement element) => IsTableCell(element) || HasHeaderCellSemantics(element)
        || HtmlAccessibilitySemantics.HasRole(element, "cell") || HtmlAccessibilitySemantics.HasRole(element, "gridcell");

    /// <summary>A source-preserving table formatting boundary; it does not create an accessibility table role.</summary>
    private sealed class TableFormattingStructure {
        internal TableFormattingStructure(IElement element, IReadOnlyList<TableFormattingRow> rows, IElement? caption, bool anonymous,
            IElement? headerGroup, IElement? footerGroup) {
            Element = element;
            Rows = rows;
            Caption = caption;
            Anonymous = anonymous;
            HeaderGroup = headerGroup;
            FooterGroup = footerGroup;
        }
        internal IElement Element { get; }
        internal IReadOnlyList<TableFormattingRow> Rows { get; }
        internal IElement? Caption { get; }
        internal bool Anonymous { get; }
        internal IElement? HeaderGroup { get; }
        internal IElement? FooterGroup { get; }
        internal bool SemanticTable => !Anonymous && HasTableSemantics(Element);
    }

    /// <summary>One authored or anonymous row with cells resolved against their actual DOM parents.</summary>
    private sealed class TableFormattingRow {
        internal TableFormattingRow(IElement element, HtmlRenderBoxStyle style, IElement? group,
            HtmlRenderBoxStyle? groupStyle, IReadOnlyList<TableFormattingCell> cells, bool anonymous, string? structureKey = null) {
            Element = element;
            Style = style;
            GroupElement = group;
            GroupStyle = groupStyle;
            Cells = cells;
            Anonymous = anonymous;
            StructureKey = structureKey;
        }
        internal IElement Element { get; }
        internal HtmlRenderBoxStyle Style { get; }
        internal IElement? GroupElement { get; }
        internal HtmlRenderBoxStyle? GroupStyle { get; }
        internal IReadOnlyList<TableFormattingCell> Cells { get; }
        internal bool Anonymous { get; }
        internal string? StructureKey { get; }
        internal IElement? ContinuationElement => Anonymous ? Cells.Select(cell => cell.ContinuationElement).FirstOrDefault(element => element != null) : Element;
        internal bool Contains(IElement target) => !Anonymous && ContainsElementOrSelf(Element, target)
            || Cells.Any(cell => cell.Nodes == null ? ContainsElementOrSelf(cell.Element, target)
                : cell.Nodes.OfType<IElement>().Any(element => ContainsElementOrSelf(element, target)));
    }

    /// <summary>A cell retains its source element or the original nodes of anonymous cell content.</summary>
    private sealed class TableFormattingCell {
        internal TableFormattingCell(IElement element, HtmlRenderBoxStyle parentStyle, IReadOnlyList<INode>? nodes = null, string? structureKey = null) {
            Element = element;
            ParentStyle = parentStyle;
            Nodes = nodes;
            StructureKey = structureKey;
        }
        internal IElement Element { get; }
        internal HtmlRenderBoxStyle ParentStyle { get; }
        internal IReadOnlyList<INode>? Nodes { get; }
        internal string? StructureKey { get; }
        internal INode SourceNode => Nodes?.FirstOrDefault() ?? Element;
        internal int ColumnSpan => Nodes == null && IsTableCell(Element) ? ReadSpan(Element.GetAttribute("colspan"), 1000) : 1;
        internal string? RowSpan => Nodes == null && IsTableCell(Element) ? Element.GetAttribute("rowspan") : null;
        internal IElement? ContinuationElement => Nodes == null ? Element : Nodes.OfType<IElement>().FirstOrDefault();
    }

    private sealed class TableFlowEntry {
        internal TableFlowEntry(INode node) { Node = node; }
        internal TableFlowEntry(IReadOnlyList<INode> nodes) { TableNodes = nodes; }
        internal INode? Node { get; }
        internal IReadOnlyList<INode>? TableNodes { get; }
    }
}

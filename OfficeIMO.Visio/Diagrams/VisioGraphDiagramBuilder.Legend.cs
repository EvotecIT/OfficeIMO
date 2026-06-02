using System;
using System.Collections.Generic;
using System.Linq;
using OfficeIMO.Visio.Stencils;
using Color = OfficeIMO.Drawing.OfficeColor;

namespace OfficeIMO.Visio.Diagrams {
    /// <summary>
    /// High-level builder for generic graph diagrams where OfficeIMO lays out arbitrary nodes and edges.
    /// </summary>
    public sealed partial class VisioGraphDiagramBuilder {
        private void AddTitle(VisioPage page) {
            if (string.IsNullOrWhiteSpace(_titleText)) {
                return;
            }

            double y = _pageHeight - _topMargin - (_titleHeight / 2D);
            VisioShape title = page.AddTextBox(_titleId, _pageWidth / 2D, y, Math.Max(1D, _pageWidth - _leftMargin - _rightMargin), _titleHeight, _titleText, _unit);
            title.TextStyle = VisioDiagramTitleStyles.Create(_theme);
            MarkDiagramAdornment(title);
        }

        private void AddLegend(VisioPage page) {
            IReadOnlyList<LegendItem> items = GetLegendItems();
            if (!_showLegend || items.Count == 0) {
                return;
            }

            int columns = GetLegendColumnCount(items.Count);
            double availableWidth = Math.Max(1D, _pageWidth - _leftMargin - _rightMargin);
            double columnWidth = availableWidth / columns;
            double legendTop = _pageHeight - _topMargin - TitleHeaderHeight;
            double titleY = legendTop - (LegendTitleHeight / 2D);
            VisioShape title = page.AddTextBox(CreateGeneratedId("legend-title"), _leftMargin + (availableWidth / 2D), titleY, availableWidth, LegendTitleHeight, _legendTitle, _unit);
            title.TextStyle = CreateLegendTextStyle();
            MarkDiagramAdornment(title);

            double firstRowY = titleY - (LegendTitleHeight / 2D) - 0.08D - (LegendRowHeight / 2D);
            for (int i = 0; i < items.Count; i++) {
                int column = i % columns;
                int row = i / columns;
                double x = _leftMargin + (column * columnWidth);
                double y = firstRowY - (row * LegendRowHeight);
                AddLegendItem(page, items[i], x, y, Math.Max(1.8D, columnWidth - 0.1D));
            }
        }

        private void AddLegendItem(VisioPage page, LegendItem item, double left, double y, double width) {
            double sampleX = left + 0.28D;
            if (item.NodeKind.HasValue) {
                VisioShape sample = new VisioShape(CreateGeneratedId("legend-" + item.IdSuffix + "-sample"), sampleX.ToInches(_unit), y.ToInches(_unit), 0.34D.ToInches(_unit), 0.18D.ToInches(_unit), string.Empty) { NameU = "Rectangle" };
                GetNodeStyle(item.NodeKind.Value).ApplyTo(sample);
                MarkDiagramAdornment(sample);
                page.Shapes.Add(sample);
            } else if (item.ConnectorKind.HasValue) {
                VisioConnectorStyle style = GetConnectorStyle(item.ConnectorKind.Value, item.Directed);
                VisioShape sample = new VisioShape(CreateGeneratedId("legend-" + item.IdSuffix + "-sample"), sampleX.ToInches(_unit), y.ToInches(_unit), 0.48D.ToInches(_unit), 0.06D.ToInches(_unit), string.Empty) { NameU = "Rectangle" };
                sample.FillPattern = 0;
                sample.LineColor = style.LineColor;
                sample.LinePattern = style.LinePattern;
                sample.LineWeight = Math.Max(0.016D, style.LineWeight);
                MarkDiagramAdornment(sample);
                page.Shapes.Add(sample);
            }

            double labelLeft = left + 0.72D;
            double labelWidth = Math.Max(0.8D, width - 0.82D);
            VisioShape label = page.AddTextBox(CreateGeneratedId("legend-" + item.IdSuffix + "-text"), labelLeft + (labelWidth / 2D), y, labelWidth, 0.22D, item.Label, _unit);
            label.TextStyle = CreateLegendTextStyle();
            MarkDiagramAdornment(label);
        }

        private IReadOnlyList<LegendItem> GetLegendItems() {
            List<LegendItem> items = new();
            if (_showLegend && _legendIncludeNodeKinds) {
                foreach (VisioGraphNodeKind kind in _nodes.Select(node => node.Kind).Distinct().OrderBy(kind => kind.ToString(), StringComparer.Ordinal)) {
                    items.Add(new LegendItem("node-" + SlugId(kind.ToString()), GetNodeKindLabel(kind), kind, null, true));
                }
            }

            if (_showLegend && _legendIncludeConnectorKinds) {
                foreach (IGrouping<string, EdgeItem> group in _edges.GroupBy(CreateLegendConnectorKey).OrderBy(group => group.Key, StringComparer.Ordinal)) {
                    EdgeItem edge = group.First();
                    items.Add(new LegendItem("edge-" + SlugId(group.Key), GetConnectorLegendLabel(edge.Kind, edge.Directed), null, edge.Kind, edge.Directed));
                }
            }

            return items.AsReadOnly();
        }

        private static string CreateLegendConnectorKey(EdgeItem edge) {
            return edge.Kind.ToString() + "-" + (edge.Directed ? "directed" : "relationship");
        }

        private int GetLegendColumnCount(int itemCount) {
            double availableWidth = Math.Max(1D, _pageWidth - _leftMargin - _rightMargin);
            if (itemCount < 2 || availableWidth < 5.6D) {
                return 1;
            }

            return 2;
        }

        private double LegendHeaderHeight {
            get {
                IReadOnlyList<LegendItem> items = GetLegendItems();
                if (!_showLegend || items.Count == 0) {
                    return 0D;
                }

                int columns = GetLegendColumnCount(items.Count);
                int rows = (int)Math.Ceiling(items.Count / (double)columns);
                return LegendTitleHeight + 0.08D + (rows * LegendRowHeight) + LegendGap;
            }
        }

        private double TitleHeaderHeight => string.IsNullOrWhiteSpace(_titleText) ? 0D : _titleHeight + _titleGap;

        private VisioTextStyle CreateLegendTextStyle() {
            VisioTextStyle style = _theme.Connector.TextStyle?.Clone() ?? new VisioTextStyle();
            style.FontFamily = string.IsNullOrWhiteSpace(style.FontFamily) ? "Aptos" : style.FontFamily;
            style.Size = Math.Max(style.Size ?? 0D, 8.5D);
            if (!style.Color.HasValue) {
                style.Color = _theme.Container.TextStyle?.Color ?? _theme.Connector.LineColor;
            }
            style.BackgroundTransparency = 100;
            style.HorizontalAlignment = VisioTextHorizontalAlignment.Left;
            style.VerticalAlignment = VisioTextVerticalAlignment.Middle;
            return style;
        }

        private static void MarkDiagramAdornment(VisioShape shape) {
            VisioSemanticUserCells.MarkGeneratedAdornment(shape);
        }
    }
}

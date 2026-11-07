using OfficeIMO.Drawing;
using OfficeIMO.Html;
using PptCore = OfficeIMO.PowerPoint;

namespace OfficeIMO.PowerPoint.Html;

public static partial class PowerPointHtmlConverterExtensions {
    private static void AppendTable(StringBuilder body, PptCore.PowerPointTable table, bool includeShapeMetadata = false, IList<HtmlDiagnostic>? diagnostics = null) {
        DocumentFormat.OpenXml.Drawing.TableStyleEntry? tableStyle = PptCore.PowerPointSlideImageRenderer.ResolveTableStyle(table);
        body.Append("<table class=\"officeimo-table\"");
        if (includeShapeMetadata) {
            AppendSemanticShapeAttributes(body, table, "table");
        }

        body.Append(" style=\"border-collapse:collapse;table-layout:fixed;width:")
            .Append(FormatNumber(Enumerable.Range(0, table.ColumnCount).Sum(column => table.GetColumnWidth(column).ToPoints())))
            .Append("pt\"><colgroup>");
        for (int column = 0; column < table.ColumnCount; column++) {
            body.Append("<col style=\"width:").Append(FormatNumber(table.GetColumnWidth(column).ToPoints())).Append("pt\">");
        }
        body.Append("</colgroup><tbody>");
        IReadOnlyList<PptCore.PowerPointTableRow> rowItems = table.Rows;
        for (int rowIndex = 0; rowIndex < rowItems.Count; rowIndex++) {
            PptCore.PowerPointTableRow row = rowItems[rowIndex];
            body.Append("<tr style=\"height:").Append(FormatNumber(table.GetRowHeight(rowIndex).ToPoints())).Append("pt\">");
            IReadOnlyList<PptCore.PowerPointTableCell> cells = row.Cells;
            for (int columnIndex = 0; columnIndex < cells.Count; columnIndex++) {
                PptCore.PowerPointTableCell cell = cells[columnIndex];
                if (cell.IsMergedCell) {
                    continue;
                }

                body.Append("<td");
                (int rows, int columns) = cell.Merge;
                if (rows > 1) {
                    body.Append(" rowspan=\"").Append(rows.ToString(CultureInfo.InvariantCulture)).Append('"');
                }

                if (columns > 1) {
                    body.Append(" colspan=\"").Append(columns.ToString(CultureInfo.InvariantCulture)).Append('"');
                }

                AppendSourceTableCellStyle(body, table, cell, rowIndex, columnIndex, tableStyle, diagnostics);
                body.Append('>');
                if (includeShapeMetadata) {
                    IReadOnlyList<DocumentFormat.OpenXml.Drawing.TableCellTextStyle> tableTextStyles =
                        PptCore.PowerPointSlideImageRenderer.ResolveTableCellTextStylesForExport(
                            table, rowIndex, columnIndex, tableStyle);
                    if (RequiresSemanticTableCellContent(cell, tableTextStyles)) {
                        DocumentFormat.OpenXml.Drawing.TextBody? textBody = cell.Cell.TextBody;
                        DocumentFormat.OpenXml.OpenXmlCompositeElement? masterTextStyle = cell.SlidePart?.SlideLayoutPart?.SlideMasterPart?
                            .SlideMaster?.TextStyles?.OtherStyle;
                        AppendSemanticParagraphContent(
                            body, cell.Paragraphs, textBody?.ListStyle, masterTextStyle, tableTextStyles);
                    } else {
                        body.Append(OfficeHtmlText.Escape(NormalizeText(cell.Text)));
                    }
                } else {
                    body.Append(OfficeHtmlText.Escape(NormalizeText(cell.Text)));
                }
                body.Append("</td>");
            }

            body.Append("</tr>");
        }

        body.Append("</tbody></table>");
    }

    private static void AppendSourceTableCellStyle(StringBuilder body, PptCore.PowerPointTable table,
        PptCore.PowerPointTableCell cell, int row, int column, DocumentFormat.OpenXml.Drawing.TableStyleEntry? tableStyle,
        IList<HtmlDiagnostic>? diagnostics) {
        var css = new StringBuilder();
        AppendCss(css, "background-color", PptCore.PowerPointSlideImageRenderer.ResolveTableCellFillColorForExport(table, row, column, tableStyle).ToCssColor());
        AppendCss(css, "padding-left", FormatNumber(cell.PaddingLeft?.ToPoints() ?? 7.2D) + "pt");
        AppendCss(css, "padding-right", FormatNumber(cell.PaddingRight?.ToPoints() ?? 7.2D) + "pt");
        AppendCss(css, "padding-top", FormatNumber(cell.PaddingTop?.ToPoints() ?? 3.6D) + "pt");
        AppendCss(css, "padding-bottom", FormatNumber(cell.PaddingBottom?.ToPoints() ?? 3.6D) + "pt");
        AppendCss(css, "text-align", cell.HorizontalAlignment switch {
            PptCore.PowerPointTextAlignment.Center => "center",
            PptCore.PowerPointTextAlignment.Right => "right",
            PptCore.PowerPointTextAlignment.Justified or PptCore.PowerPointTextAlignment.JustifiedLow
                or PptCore.PowerPointTextAlignment.Distributed or PptCore.PowerPointTextAlignment.ThaiDistributed => "justify",
            _ => "left"
        });
        AppendCss(css, "vertical-align", cell.VerticalAlignment switch {
            PptCore.PowerPointTextVerticalAlignment.Center => "middle",
            PptCore.PowerPointTextVerticalAlignment.Bottom => "bottom",
            _ => "top"
        });
        OfficeBorderBox borders = PptCore.PowerPointSlideImageRenderer.ResolveTableCellBordersForExport(table, row, column, tableStyle);
        AppendBorder("left", borders.Left);
        AppendBorder("top", borders.Top);
        AppendBorder("right", borders.Right);
        AppendBorder("bottom", borders.Bottom);
        if (borders.DiagonalDown?.IsVisible == true || borders.DiagonalUp?.IsVisible == true) {
            ReportBorderLoss("Diagonal table borders have no CSS table-border equivalent.");
        }
        body.Append(" style=\"").Append(OfficeHtmlText.EscapeAttribute(css.ToString())).Append('"');

        void AppendBorder(string edge, OfficeBorderSide? side) {
            if (!side.HasValue || !side.Value.IsVisible) {
                AppendCss(css, "border-" + edge, "none");
                return;
            }
            OfficeBorderSide border = side.Value;
            string style = border.LineKind == OfficeBorderLineKind.Double ? "double" : border.DashStyle switch {
                OfficeStrokeDashStyle.Dot => "dotted",
                OfficeStrokeDashStyle.Dash or OfficeStrokeDashStyle.DashDot or OfficeStrokeDashStyle.DashDotDot => "dashed",
                _ => "solid"
            };
            double width = border.LineKind == OfficeBorderLineKind.Double
                ? border.Width + border.DoubleLineSeparation : border.Width;
            AppendCss(css, "border-" + edge, FormatNumber(width) + "pt " + style + " " + border.Color.ToCssColor());
            if (border.DashStyle is OfficeStrokeDashStyle.DashDot or OfficeStrokeDashStyle.DashDotDot) {
                ReportBorderLoss("Dash-dot table borders are approximated by CSS dashed borders.");
            }
            if (border.LineKind == OfficeBorderLineKind.Double && Math.Abs(border.DoubleLineSeparation - 2D * border.Width) > .001D) {
                ReportBorderLoss("Double table-border spacing is approximated by CSS double borders.");
            }
        }

        void ReportBorderLoss(string message) => diagnostics?.Add(new HtmlDiagnostic(
            "OfficeIMO.PowerPoint.Html", "powerpoint-table-border-approximated", message,
            HtmlDiagnosticSeverity.Warning, "powerpoint:table:" + table.DrawingOrder.ToString(CultureInfo.InvariantCulture),
            lossKind: OfficeConversionLossKind.Approximation));
    }
}

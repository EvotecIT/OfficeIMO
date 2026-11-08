using System.Globalization;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word;

public partial class WordDocument {
    internal static bool IsThemeHeaderCell(TableCell cell) {
        TableRow? row = cell.Ancestors<TableRow>().FirstOrDefault();
        Table? table = row?.Ancestors<Table>().FirstOrDefault();
        if (row == null || table == null) return false;
        TableHeader? declared = row.TableRowProperties?.GetFirstChild<TableHeader>();
        return declared != null && WordTableRow.ReadOnOff(declared) ||
            ReferenceEquals(WordTableContent.Rows(table).FirstOrDefault(), row) && WordTable.HasFirstRowFormatting(table);
    }

    private static void ApplyThemeTable(Table table, OfficeTheme theme) {
        Style appearance = WordThemeDefaults.CreateTableStyle(theme);
        StyleTableProperties tableVisuals = appearance.StyleTableProperties!;
        TableProperties properties = table.GetFirstChild<TableProperties>() ?? table.PrependChild(new TableProperties());
        WordThemeDefaults.SetVisualChild(properties, tableVisuals.TableBorders!);
        TableCellMarginDefault padding = tableVisuals.TableCellMarginDefault!;
        WordThemeDefaults.SetVisualChild(properties, padding);
        properties.Shading = new Shading { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "FFFFFF" };
        TableStyleProperties header = appearance.Elements<TableStyleProperties>().Single(part => part.Type?.Value == TableStyleOverrideValues.FirstRow);
        int bodyIndex = 0;
        foreach (TableRow row in WordTableContent.Rows(table)) {
            TableCell[] cells = WordTableContent.Cells(row).ToArray();
            bool isHeader = cells.Length > 0 && IsThemeHeaderCell(cells[0]);
            string fill = isHeader ? theme.TableStyle.HeaderFill.Resolve(theme).ToRgbHex() :
                theme.TableStyle.Banding && bodyIndex % 2 == 1
                    ? (theme.IsDark ? OfficeTheme.Modern.TableStyle.BandFill : theme.TableStyle.BandFill).Resolve(theme).ToRgbHex() : "FFFFFF";
            foreach (TableCell cell in cells) {
                TableCellProperties cellProperties = cell.TableCellProperties ??= new TableCellProperties();
                cellProperties.Shading = new Shading { Val = ShadingPatternValues.Clear, Color = "auto", Fill = fill };
                cellProperties.TableCellBorders?.Remove();
                if (isHeader)
                    cellProperties.TableCellBorders = (TableCellBorders)header.TableStyleConditionalFormattingTableCellProperties!.TableCellBorders!.CloneNode(true);
                cellProperties.TableCellMargin = new TableCellMargin(
                    padding.TopMargin!.CloneNode(true),
                    new LeftMargin { Width = padding.GetFirstChild<TableCellLeftMargin>()!.Width!.Value.ToString(CultureInfo.InvariantCulture), Type = TableWidthUnitValues.Dxa },
                    padding.BottomMargin!.CloneNode(true),
                    new RightMargin { Width = padding.GetFirstChild<TableCellRightMargin>()!.Width!.Value.ToString(CultureInfo.InvariantCulture), Type = TableWidthUnitValues.Dxa });
            }
            if (!isHeader) bodyIndex++;
        }
    }
}

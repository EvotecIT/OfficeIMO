using System.Globalization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word;

public partial class WordDocument {
    /// <summary>Uses declared headers or an active first-row region in the existing table style.</summary>
    internal static bool IsThemeHeaderCell(TableCell cell) {
        TableRow? row = cell.Ancestors<TableRow>().FirstOrDefault();
        Table? table = row?.Ancestors<Table>().FirstOrDefault();
        if (row == null || table == null) return false;
        TableHeader? declared = row.TableRowProperties?.GetFirstChild<TableHeader>();
        if (declared != null && WordTableRow.ReadOnOff(declared)) return true;
        if (!ReferenceEquals(WordTableContent.Rows(table).FirstOrDefault(), row) || !WordTable.HasFirstRowFormatting(table)) return false;
        // Native first-row look defaults select style regions; they do not declare a header by themselves.
        OpenXmlPartRootElement? root = table.Ancestors<OpenXmlPartRootElement>().LastOrDefault();
        Styles? styles = (root?.OpenXmlPart?.OpenXmlPackage as WordprocessingDocument)?.MainDocumentPart?.StyleDefinitionsPart?.Styles;
        string? styleId = table.GetFirstChild<TableProperties>()?.TableStyle?.Val?.Value;
        return WordRunFormatting.EnumerateStyleChain(styles, styleId, StyleValues.Table)
            .Any(style => style.Elements<TableStyleProperties>().Any(part => part.Type?.Value == TableStyleOverrideValues.FirstRow));
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

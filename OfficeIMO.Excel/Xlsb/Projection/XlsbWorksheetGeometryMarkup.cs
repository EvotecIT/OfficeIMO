using System.Xml;
using OfficeIMO.Excel.Xlsb.Model;

namespace OfficeIMO.Excel.Xlsb.Projection {
    /// <summary>Writes the worksheet layout recorded in the native worksheet, apart from the ordered cell grid.</summary>
    internal static class XlsbWorksheetGeometryMarkup {
        internal static void WriteBeforeGrid(OfficeXmlNode worksheet, XlsbSheet source, bool? tabSelected = null) {
            if (source.UsedRange != null) {
                XlsbWorksheetMarkup.Append(worksheet, "dimension").SetAttribute("ref", source.UsedRange.ToA1Reference());
            }
            OfficeXmlNode sheetViews = XlsbWorksheetMarkup.Append(worksheet, "sheetViews");
            WriteView(sheetViews, source, tabSelected ?? source.TabSelected == true);
            if (source.FormatInfo != null) WriteFormat(XlsbWorksheetMarkup.Append(worksheet, "sheetFormatPr"), source.FormatInfo);
            if (source.Columns.Count != 0) WriteColumns(XlsbWorksheetMarkup.Append(worksheet, "cols"), source.Columns);
        }

        internal static void WriteMerges(OfficeXmlNode worksheet, IReadOnlyList<XlsbCellRange> ranges) {
            if (ranges.Count == 0) return;
            OfficeXmlNode merges = XlsbWorksheetMarkup.Append(worksheet, "mergeCells");
            merges.SetUInt32Attribute("count", (uint)ranges.Count);
            foreach (XlsbCellRange range in ranges) {
                XlsbWorksheetMarkup.Append(merges, "mergeCell").SetAttribute("ref", range.ToA1Reference());
            }
        }

        internal static void WriteRowAttributes(XmlWriter writer, XlsbRowInfo source) {
            if (source.CustomHeight) {
                writer.WriteAttributeString("ht", InvariantNumberText.Get(source.HeightTwips / 20D));
                writer.WriteAttributeString("customHeight", "1");
            }
            if (source.Hidden) writer.WriteAttributeString("hidden", "1");
            if (source.OutlineLevel != 0) writer.WriteAttributeString("outlineLevel", InvariantNumberText.Get((uint)source.OutlineLevel));
            if (source.Collapsed) writer.WriteAttributeString("collapsed", "1");
            if (source.CustomFormat) {
                writer.WriteAttributeString("s", InvariantNumberText.Get(source.StyleIndex));
                writer.WriteAttributeString("customFormat", "1");
            }
        }

        private static void WriteView(OfficeXmlNode views, XlsbSheet source, bool tabSelected) {
            OfficeXmlNode view = XlsbWorksheetMarkup.Append(views, "sheetView");
            view.SetUInt32Attribute("workbookViewId", 0U);
            if (source.ShowGridLines == false) view.SetBooleanAttribute("showGridLines", false);
            if (source.View != null) WriteViewSettings(view, source.View);
            if (tabSelected) view.SetBooleanAttribute("tabSelected", true);

            if (source.Pane is { } pane) {
                OfficeXmlNode paneElement = XlsbWorksheetMarkup.Append(view, "pane");
                paneElement.SetDoubleAttribute("xSplit", pane.HorizontalSplit);
                paneElement.SetDoubleAttribute("ySplit", pane.VerticalSplit);
                paneElement.SetAttribute("topLeftCell", ExcelA1.CellReference(pane.TopRow + 1, pane.LeftColumn + 1));
                paneElement.SetAttribute("activePane", PaneName(pane.ActivePane));
                paneElement.SetAttribute("state", pane.Frozen ? pane.FrozenNoSplit ? "frozenSplit" : "frozen" : "split");
            }

            IReadOnlyList<XlsbSelection> selections = source.View?.Selections ?? Array.Empty<XlsbSelection>();
            if (selections.Count == 1 && IsDefaultSelection(selections[0])) return;
            foreach (XlsbSelection selection in selections) {
                OfficeXmlNode node = XlsbWorksheetMarkup.Append(view, "selection");
                if (selection.Pane != 3U) node.SetAttribute("pane", PaneName(selection.Pane));
                node.SetAttribute("activeCell", ExcelA1.CellReference(selection.ActiveRow, selection.ActiveColumn));
                if (selection.ActiveCellId != 0U) node.SetUInt32Attribute("activeCellId", selection.ActiveCellId);
                node.SetAttribute("sqref", string.Join(" ", selection.Ranges.Select(range => range.ToA1Reference())));
            }
        }

        // Only fields that disagree with the SpreadsheetML defaults become attributes.
        private static void WriteViewSettings(OfficeXmlNode node, XlsbSheetView source) {
            ushort flags = source.Flags;
            if ((flags & 0x0001) != 0) node.SetBooleanAttribute("windowProtection", true);
            if ((flags & 0x0002) != 0) node.SetBooleanAttribute("showFormulas", true);
            if ((flags & 0x0008) == 0) node.SetBooleanAttribute("showRowColHeaders", false);
            if ((flags & 0x0010) == 0) node.SetBooleanAttribute("showZeros", false);
            if ((flags & 0x0020) != 0) node.SetBooleanAttribute("rightToLeft", true);
            if ((flags & 0x0080) == 0) node.SetBooleanAttribute("showRuler", false);
            if ((flags & 0x0100) == 0) node.SetBooleanAttribute("showOutlineSymbols", false);
            if ((flags & 0x0200) == 0) node.SetBooleanAttribute("defaultGridColor", false);
            if ((flags & 0x0400) != 0) node.SetBooleanAttribute("showWhiteSpace", false);
            if (source.View == 1U) node.SetAttribute("view", "pageBreakPreview");
            else if (source.View == 2U) node.SetAttribute("view", "pageLayout");
            if (source.TopRow != 0 || source.LeftColumn != 0) node.SetAttribute("topLeftCell", ExcelA1.CellReference(source.TopRow + 1, source.LeftColumn + 1));
            if (source.GridColor != 64) node.SetUInt32Attribute("colorId", source.GridColor);
            if (source.Zoom != 100) node.SetUInt32Attribute("zoomScale", source.Zoom);
            if (source.ZoomNormal != 0) node.SetUInt32Attribute("zoomScaleNormal", source.ZoomNormal);
            if (source.ZoomPageBreak != 0) node.SetUInt32Attribute("zoomScaleSheetLayoutView", source.ZoomPageBreak);
            if (source.ZoomPageLayout != 0) node.SetUInt32Attribute("zoomScalePageLayoutView", source.ZoomPageLayout);
        }

        private static bool IsDefaultSelection(XlsbSelection selection) =>
            selection.Pane == 3U && selection.ActiveRow == 1 && selection.ActiveColumn == 1 && selection.ActiveCellId == 0U
            && selection.Ranges.Count == 1 && selection.Ranges[0].FirstRow == 1 && selection.Ranges[0].LastRow == 1
            && selection.Ranges[0].FirstColumn == 1 && selection.Ranges[0].LastColumn == 1;

        private static string PaneName(uint pane) => pane switch {
            0U => "bottomRight",
            1U => "topRight",
            2U => "bottomLeft",
            3U => "topLeft",
            _ => throw new InvalidDataException($"Unsupported XLSB pane index {pane}.")
        };

        private static void WriteFormat(OfficeXmlNode node, XlsbWorksheetFormatInfo source) {
            if (source.DefaultColumnWidth.HasValue) node.SetDoubleAttribute("defaultColWidth", source.DefaultColumnWidth.Value);
            if (source.BaseColumnWidth.HasValue) node.SetUInt32Attribute("baseColWidth", source.BaseColumnWidth.Value);
            if (source.DefaultRowHeight > 0D) node.SetDoubleAttribute("defaultRowHeight", source.DefaultRowHeight);
            if (source.CustomDefaultRowHeight) node.SetBooleanAttribute("customHeight", true);
            if (source.DefaultRowsHidden) node.SetBooleanAttribute("zeroHeight", true);
            if (source.MaximumRowOutlineLevel > 0) node.SetUInt32Attribute("outlineLevelRow", source.MaximumRowOutlineLevel);
            if (source.MaximumColumnOutlineLevel > 0) node.SetUInt32Attribute("outlineLevelCol", source.MaximumColumnOutlineLevel);
        }

        private static void WriteColumns(OfficeXmlNode node, IReadOnlyList<XlsbColumnInfo> source) {
            foreach (XlsbColumnInfo item in source) {
                OfficeXmlNode column = XlsbWorksheetMarkup.Append(node, "col");
                column.SetUInt32Attribute("min", checked((uint)item.FirstColumn));
                column.SetUInt32Attribute("max", checked((uint)item.LastColumn));
                column.SetDoubleAttribute("width", item.Width);
                if (item.UserSet) column.SetBooleanAttribute("customWidth", true);
                if (item.BestFit) column.SetBooleanAttribute("bestFit", true);
                if (item.Hidden) column.SetBooleanAttribute("hidden", true);
                if (item.StyleIndex != 0) column.SetUInt32Attribute("style", item.StyleIndex);
                if (item.OutlineLevel != 0) column.SetUInt32Attribute("outlineLevel", item.OutlineLevel);
                if (item.Collapsed) column.SetBooleanAttribute("collapsed", true);
                if (item.Phonetic) column.SetBooleanAttribute("phonetic", true);
            }
        }
    }
}

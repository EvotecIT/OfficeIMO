using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.Excel.Xlsb.Model;
using OfficeIMO.Excel.Xlsb.Package;

namespace OfficeIMO.Excel.Xlsb.Projection {
    /// <summary>Projects worksheet metadata into an owned SpreadsheetML part around its streamed grid.</summary>
    internal static class XlsbOwnedWorksheetMarkup {
        private const int Spreadsheet = OfficeXmlNamespaces.SpreadsheetML;
        private const int Relationships = OfficeXmlNamespaces.Relationships;
        private const string HyperlinkRelationship = OfficeXmlNamespaces.TransitionalRelationshipTypePrefix + "hyperlink";

        internal static OfficeXmlNode Create(OfficeOpcEditablePackage package, string partName, XlsbSheet source,
            bool tabSelected, XlsbOwnedRelatedParts related) {
            var tree = OfficeXmlTree.Create(package.Namespaces, Spreadsheet, "worksheet");
            var worksheet = new OfficeXmlNode(tree, tree.Root);
            if (source.Properties != null) XlsbWorksheetPropertiesProjector.Apply(worksheet, source.Properties);
            XlsbWorksheetGeometryMarkup.WriteBeforeGrid(worksheet, source, tabSelected);
            XlsbWorksheetMarkup.Append(worksheet, "sheetData");
            if (source.Protection != null) XlsbWorksheetProtectionProjector.Apply(worksheet, source.Protection);
            if (source.AutoFilter != null) WriteFilter(worksheet, source.AutoFilter);
            XlsbWorksheetGeometryMarkup.WriteMerges(worksheet, source.MergedRanges);

            WriteFeatures(worksheet, source, "conditionalFormatting", "dataValidations");
            WriteHyperlinks(package, partName, worksheet, source.Hyperlinks);
            XlsbWorksheetPrintSettingsProjector.Apply(worksheet, source.PrintOptions, source.PageMargins, source.PageSetup, source.HeaderFooter);
            WriteFeatures(worksheet, source, "rowBreaks", "colBreaks", "ignoredErrors");
            if (source.DrawingRelationshipId is string id && source.DrawingPartName is string drawing) {
                related.Link(partName, drawing, source.Relationships[id].Type, id);
                XlsbWorksheetMarkup.Append(worksheet, "drawing").SetAttribute(Relationships, "id", id);
            }
            return worksheet;
        }

        internal static void WriteChartSheet(OfficeXmlNode chartsheet, XlsbSheet source, bool tabSelected,
            string partName, XlsbOwnedRelatedParts related) {
            if (!string.IsNullOrEmpty(source.ChartCodeName))
                XlsbWorksheetMarkup.Append(chartsheet, "sheetPr").SetAttribute("codeName", source.ChartCodeName);
            OfficeXmlNode view = XlsbWorksheetMarkup.Append(XlsbWorksheetMarkup.Append(chartsheet, "sheetViews"), "sheetView");
            view.SetUInt32Attribute("workbookViewId", 0U);
            if (tabSelected) view.SetBooleanAttribute("tabSelected", true);
            string drawingId = source.DrawingRelationshipId!;
            XlsbPackageRelationship drawing = source.Relationships[drawingId];
            related.Link(partName, XlsbPackagePartReader.ResolveTarget(source.PartName!, drawing.Target), drawing.Type, drawingId);
            XlsbWorksheetMarkup.Append(chartsheet, "drawing").SetAttribute(Relationships, "id", drawingId);
        }

        private static void WriteFilter(OfficeXmlNode worksheet, XlsbAutoFilter source) {
            OfficeXmlNode filter = XlsbWorksheetMarkup.Append(worksheet, "autoFilter");
            filter.SetAttribute("ref", source.Range.ToA1Reference());
            foreach (XlsbAutoFilterColumn sourceColumn in source.Columns) {
                if (sourceColumn.HasUnsupportedContent) continue;
                OfficeXmlNode column = XlsbWorksheetMarkup.Append(filter, "filterColumn");
                column.SetUInt32Attribute("colId", sourceColumn.ColumnId);
                if (sourceColumn.CriterionMarkup is { Exists: true } criterion) {
                    column.Append(criterion);
                } else {
                    OfficeXmlNode values = XlsbWorksheetMarkup.Append(column, "filters");
                    values.SetBooleanAttribute("blank", sourceColumn.IncludeBlank);
                    foreach (string value in sourceColumn.Values)
                        XlsbWorksheetMarkup.Append(values, "filter").SetAttribute("val", value);
                }
            }
        }

        private static void WriteHyperlinks(OfficeOpcEditablePackage package, string partName, OfficeXmlNode worksheet,
            IReadOnlyList<XlsbHyperlink> links) {
            if (links.Count == 0) return;
            OfficeXmlNode container = XlsbWorksheetMarkup.Append(worksheet, "hyperlinks");
            foreach (XlsbHyperlink source in links) {
                OfficeXmlNode link = XlsbWorksheetMarkup.Append(container, "hyperlink");
                link.SetAttribute("ref", source.Range.ToA1Reference());
                if (source.IsExternal) {
                    var target = new Uri(source.ExternalTarget!, UriKind.RelativeOrAbsolute);
                    string id = package.AddRelationship(partName, HyperlinkRelationship, target.OriginalString,
                        external: true, id: source.RelationshipId);
                    link.SetAttribute(Relationships, "id", id);
                }
                if (!string.IsNullOrEmpty(source.Location)) link.SetAttribute("location", source.Location);
                if (!string.IsNullOrEmpty(source.Tooltip)) link.SetAttribute("tooltip", source.Tooltip);
                if (!string.IsNullOrEmpty(source.Display)) link.SetAttribute("display", source.Display);
            }
        }

        // Binary feature readers already supply their SpreadsheetML fragments. Match only recognised direct children
        // and group them in worksheet schema order, as the model's worksheet save does.
        private static void WriteFeatures(OfficeXmlNode worksheet, XlsbSheet source, params string[] localNames) {
            foreach (string name in localNames) {
                foreach (OfficeXmlNode element in source.ReadFeatureMarkup()) {
                    if (element.Is(Spreadsheet, name)) worksheet.Append(element);
                }
            }
        }
    }
}

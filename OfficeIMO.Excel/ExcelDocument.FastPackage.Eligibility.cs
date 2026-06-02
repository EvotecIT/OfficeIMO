using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Diagnostics;
using System.Globalization;
using System.IO.Compression;
using System.Threading;
using System.Xml;

namespace OfficeIMO.Excel {
    public partial class ExcelDocument {

        private static bool CanWriteSimpleWorksheet(WorksheetPart worksheetPart, Worksheet worksheet, out string? skipReason, bool allowDrawings = false, bool allowPivotTables = false) {
            skipReason = null;

            if (worksheetPart.WorksheetCommentsPart != null) {
                skipReason = "Worksheet contains comments.";
                return false;
            }

            if (!allowDrawings && worksheetPart.DrawingsPart != null) {
                skipReason = "Worksheet contains drawings.";
                return false;
            }

            if (!allowPivotTables && worksheetPart.PivotTableParts.Any()) {
                skipReason = "Worksheet contains pivot tables.";
                return false;
            }

            if (worksheetPart.ExternalRelationships.Any()) {
                skipReason = "Worksheet contains external relationships.";
                return false;
            }

            foreach (var child in worksheet.ChildElements) {
                if (child is not SheetProperties
                    && child is not SheetDimension
                    && child is not SheetViews
                    && child is not SheetFormatProperties
                    && child is not Columns
                    && child is not SheetData
                    && child is not SheetCalculationProperties
                    && child is not SheetProtection
                    && child is not DocumentFormat.OpenXml.Spreadsheet.ProtectedRanges
                    && child is not Scenarios
                    && child is not AutoFilter
                    && child is not SortState
                    && child is not MergeCells
                    && child is not PhoneticProperties
                    && child is not Hyperlinks
                    && child is not DocumentFormat.OpenXml.Spreadsheet.ConditionalFormatting
                    && child is not DataValidations
                    && child is not PrintOptions
                    && child is not PageMargins
                    && child is not PageSetup
                    && child is not HeaderFooter
                    && child is not RowBreaks
                    && child is not ColumnBreaks
                    && child is not CellWatches
                    && child is not DocumentFormat.OpenXml.Spreadsheet.IgnoredErrors
                    && (!allowDrawings || child is not DocumentFormat.OpenXml.Spreadsheet.Drawing)
                    && child is not TableParts) {
                    skipReason = "Worksheet contains unsupported element '" + child.LocalName + "'.";
                    return false;
                }

                if (child is not SheetData
                    && child.Descendants<DocumentFormat.OpenXml.OpenXmlUnknownElement>().Any()) {
                    skipReason = "Worksheet contains unknown Open XML elements.";
                    return false;
                }
            }

            var tableParts = worksheet.GetFirstChild<TableParts>();
            bool hasTableDefinitionParts = worksheetPart.TableDefinitionParts.Any();
            if (tableParts != null || hasTableDefinitionParts) {
                if (tableParts != null && worksheet.Elements<TableParts>().Skip(1).Any()) {
                    skipReason = "Worksheet contains multiple tableParts elements.";
                    return false;
                }

                var tableDefinitionParts = worksheetPart.TableDefinitionParts.ToList();
                var relationshipIds = new HashSet<string>(tableDefinitionParts.Select(worksheetPart.GetIdOfPart), StringComparer.Ordinal);
                var worksheetTablePartIds = tableParts == null
                    ? new List<string>()
                    : tableParts.Elements<TablePart>()
                        .Select(part => part.Id?.Value)
                        .Where(id => !string.IsNullOrEmpty(id))
                        .Select(id => id!)
                        .ToList();

                if (worksheetTablePartIds.Count != tableDefinitionParts.Count
                    || worksheetTablePartIds.Any(id => !relationshipIds.Contains(id))) {
                    skipReason = "Worksheet table relationships do not match tableParts entries.";
                    return false;
                }

                foreach (var tableDefinitionPart in tableDefinitionParts) {
                    var table = tableDefinitionPart.Table;
                    if (table == null
                        || table.Reference == null
                        || table.TableColumns == null
                        || table.Descendants<DocumentFormat.OpenXml.OpenXmlUnknownElement>().Any()) {
                        skipReason = "Worksheet contains unsupported table metadata.";
                        return false;
                    }
                }
            }

            var hyperlinks = worksheet.GetFirstChild<Hyperlinks>();
            bool hasHyperlinkRelationships = worksheetPart.HyperlinkRelationships.Any();
            if (hyperlinks != null || hasHyperlinkRelationships) {
                var hyperlinkRelationships = worksheetPart.HyperlinkRelationships.ToList();
                var hyperlinkIds = new HashSet<string>(hyperlinkRelationships.Select(relationship => relationship.Id), StringComparer.Ordinal);
                if (hyperlinks != null) {
                    foreach (var hyperlink in hyperlinks.Elements<Hyperlink>()) {
                        string? relationshipId = hyperlink.Id?.Value;
                        if (!string.IsNullOrEmpty(relationshipId) && !hyperlinkIds.Contains(relationshipId!)) {
                            skipReason = "Worksheet hyperlink relationships do not match hyperlink entries.";
                            return false;
                        }
                    }
                }
            }

            var sheetData = worksheet.GetFirstChild<SheetData>();
            if (sheetData == null) {
                return true;
            }

            foreach (var sheetDataChild in sheetData.ChildElements) {
                if (sheetDataChild is not Row) {
                    skipReason = sheetDataChild is DocumentFormat.OpenXml.OpenXmlUnknownElement
                        ? "Worksheet contains unknown Open XML elements."
                        : "Worksheet contains sheetData children outside the simple writer surface.";
                    return false;
                }
            }

            foreach (var row in sheetData.Elements<Row>()) {
                if (!IsSimpleRow(row)) {
                    skipReason = "Worksheet contains row formatting outside the simple writer surface.";
                    return false;
                }

                foreach (var rowChild in row.ChildElements) {
                    if (rowChild is DocumentFormat.OpenXml.OpenXmlUnknownElement) {
                        skipReason = "Worksheet contains unknown Open XML elements.";
                        return false;
                    }

                    if (rowChild is not Cell cell) {
                        skipReason = "Worksheet contains row children outside the simple writer surface.";
                        return false;
                    }

                    foreach (var cellChild in cell.ChildElements) {
                        if (cellChild is DocumentFormat.OpenXml.OpenXmlUnknownElement) {
                            skipReason = "Worksheet contains unknown Open XML elements.";
                            return false;
                        }
                    }

                    if (cell.InlineString != null) {
                        if (cell.InlineString.Descendants<DocumentFormat.OpenXml.OpenXmlUnknownElement>().Any()) {
                            skipReason = "Worksheet inline strings contain unknown Open XML elements.";
                            return false;
                        }
                    }

                    if (cell.CellFormula != null
                        && cell.CellFormula.Descendants<DocumentFormat.OpenXml.OpenXmlUnknownElement>().Any()) {
                        skipReason = "Worksheet contains formula metadata outside the simple writer surface.";
                        return false;
                    }

                    var dataType = cell.DataType?.Value;
                    if (dataType != null
                        && dataType != CellValues.Number
                        && dataType != CellValues.SharedString
                        && dataType != CellValues.InlineString
                        && dataType != CellValues.String
                        && dataType != CellValues.Boolean) {
                        skipReason = "Worksheet contains unsupported cell data type '" + dataType.Value.ToString() + "'.";
                        return false;
                    }
                }
            }

            return true;
        }

        private static bool IsSimpleRow(Row row) {
            foreach (var attribute in row.GetAttributes()) {
                if (!string.Equals(attribute.LocalName, "r", StringComparison.Ordinal)
                    && !string.Equals(attribute.LocalName, "hidden", StringComparison.Ordinal)
                    && !string.Equals(attribute.LocalName, "ht", StringComparison.Ordinal)
                    && !string.Equals(attribute.LocalName, "customHeight", StringComparison.Ordinal)
                    && !string.Equals(attribute.LocalName, "outlineLevel", StringComparison.Ordinal)
                    && !string.Equals(attribute.LocalName, "collapsed", StringComparison.Ordinal)) {
                    return false;
                }
            }

            return row.CustomFormat?.Value != true && row.StyleIndex == null;
        }
    }
}

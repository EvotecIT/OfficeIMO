using System.Globalization;
using System.Xml;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Core.Internal;
using static OfficeIMO.Excel.SpreadsheetXml.SpreadsheetXmlContext;

namespace OfficeIMO.Excel.SpreadsheetXml;

internal static class SpreadsheetXmlWriter {
    internal static OfficeConversionResult<byte[], ExcelSpreadsheetXmlReport> Write(byte[] input, ExcelSpreadsheetXmlOptions options) {
        if (input.Length > options.MaxInputBytes) throw new InvalidDataException("Workbook snapshot exceeds MaxInputBytes.");
        using var source = new MemoryStream(input, writable: false);
        using var package = SpreadsheetDocument.Open(source, false);
        WorkbookPart workbook = package.WorkbookPart!;
        var context = new SpreadsheetXmlContext(options);
        var styles = SpreadsheetXmlStyles.Export(workbook, context);
        Sheet[] sheets = workbook.Workbook!.Sheets?.Elements<Sheet>().ToArray() ?? Array.Empty<Sheet>();
        if (sheets.Length == 0 || sheets.Length > options.MaxWorksheets) throw new InvalidDataException("Workbook requires worksheets within MaxWorksheets.");
        foreach (OpenXmlElement child in workbook.Workbook.ChildElements) {
            if (child is not Sheets && child is not WorkbookProperties && child is not BookViews && child is not CalculationProperties && child is not DefinedNames)
                context.Omitted("workbook " + child.LocalName);
        }
        if (workbook.Workbook.WorkbookProperties is WorkbookProperties properties)
            InspectAttributes(properties, context, "date1904");
        if (workbook.Workbook.CalculationProperties?.HasAttributes == true) context.Omitted("workbook calculation settings");
        if (workbook.Workbook.BookViews?.ChildElements.Any(view => view.HasAttributes) == true) context.Omitted("workbook views");
        if (package.PackageProperties.Title != null || package.PackageProperties.Creator != null
            || package.PackageProperties.Subject != null || package.PackageProperties.Keywords != null
            || package.PackageProperties.Description != null) context.Omitted("core document properties");
        foreach (IdPartPair part in workbook.Parts) {
            if (part.OpenXmlPart is not WorksheetPart && part.OpenXmlPart is not WorkbookStylesPart
                && part.OpenXmlPart is not SharedStringTablePart && part.OpenXmlPart is not ThemePart && part.OpenXmlPart is not CalculationChainPart)
                context.Omitted("package part " + part.OpenXmlPart.ContentType);
        }
        if (package.ExtendedFilePropertiesPart != null || package.CustomFilePropertiesPart != null)
            context.Omitted("extended or custom document properties");
        using var output = new OfficeBoundedMemoryStream(options.MaxOutputBytes);
        using (var writer = XmlWriter.Create(output, new XmlWriterSettings {
            Encoding = new UTF8Encoding(false), CloseOutput = false, NewLineHandling = NewLineHandling.Entitize
        })) {
            writer.WriteStartDocument();
            writer.WriteStartElement("Workbook", Ss.NamespaceName);
            writer.WriteAttributeString("xmlns", "ss", null, Ss.NamespaceName);
            writer.WriteAttributeString("xmlns", "x", null, X.NamespaceName);
            if (workbook.Workbook.WorkbookProperties?.Date1904?.Value == true) {
                writer.WriteStartElement("x", "ExcelWorkbook", X.NamespaceName);
                writer.WriteElementString("x", "Date1904", X.NamespaceName, string.Empty);
                writer.WriteEndElement();
            }
            writer.WriteStartElement("Styles", Ss.NamespaceName);
            for (int index = 0; index < styles.Count; index++) {
                styles[index].WriteTo(writer);
            }
            writer.WriteEndElement();
            SpreadsheetXmlMetadata.WriteNames(writer, workbook, null, context);
            var strings = workbook.SharedStringTablePart?.SharedStringTable?.Elements<SharedStringItem>().ToArray() ?? Array.Empty<SharedStringItem>();
            foreach (Sheet sheet in sheets) {
                options.CancellationToken.ThrowIfCancellationRequested();
                writer.WriteStartElement("Worksheet", Ss.NamespaceName); Attribute(writer, "Name", sheet.Name?.Value ?? string.Empty);
                if (sheet.Id != null && workbook.GetPartById(sheet.Id.Value!) is WorksheetPart statePart) {
                    if (statePart.Worksheet?.GetFirstChild<SheetProtection>()?.Sheet?.Value == true) Attribute(writer, "Protected", "1");
                    if (statePart.Worksheet?.GetFirstChild<SheetViews>()?.GetFirstChild<SheetView>()?.RightToLeft?.Value == true) Attribute(writer, "RightToLeft", "1");
                }
                SpreadsheetXmlMetadata.WriteNames(writer, workbook, (uint)Array.IndexOf(sheets, sheet), context);
                if (sheet.Id == null || workbook.GetPartById(sheet.Id.Value!) is not WorksheetPart part) {
                    context.Omitted("non-worksheet sheet " + sheet.Name?.Value);
                    writer.WriteStartElement("Table", Ss.NamespaceName); writer.WriteEndElement();
                } else WriteTable(writer, part, strings, styles.Count - 1, context);
                SpreadsheetXmlOptionsCodec.Write(writer, sheet, sheet.Id != null ? workbook.GetPartById(sheet.Id.Value!) as WorksheetPart : null, context);
                writer.WriteEndElement();
            }
            writer.WriteEndElement(); writer.WriteEndDocument();
        }
        return new(output.ToArray(), context.Report());
    }

    private static void WriteTable(XmlWriter writer, WorksheetPart part, SharedStringItem[] strings, int styleCount, SpreadsheetXmlContext context) {
        SpreadsheetXmlOptionsCodec.InspectNested(part.Worksheet!, context);
        foreach (OpenXmlElement child in part.Worksheet!.ChildElements) {
            if (child is not SheetData && child is not SheetDimension && child is not SheetViews && child is not SheetFormatProperties && child is not Columns && child is not MergeCells && child is not Hyperlinks && child is not LegacyDrawing && child is not PageMargins && child is not PageSetup && child is not PrintOptions && child is not HeaderFooter && child is not SheetProperties && child is not SheetProtection)
                context.Omitted("worksheet " + child.LocalName);

            if (child is SheetViews sheetViews && sheetViews.Descendants<SheetView>().Any(view =>
                    view.GetAttributes().Any(attribute => attribute.LocalName is not "workbookViewId" and not "showGridLines" and not "showRowColHeaders" and not "zoomScale" and not "tabSelected" and not "rightToLeft")))
                context.Omitted("worksheet view settings");
        }
        foreach (IdPartPair child in part.Parts)
            if (child.OpenXmlPart is not WorksheetCommentsPart && child.OpenXmlPart is not VmlDrawingPart) context.Omitted("worksheet part " + child.OpenXmlPart.ContentType);
        if (part.ExternalRelationships.Any()) context.Omitted("worksheet external relationships");
        writer.WriteStartElement("Table", Ss.NamespaceName);
        SpreadsheetXmlGeometry.WriteColumns(writer, part, styleCount, context);
        var metadata = SpreadsheetXmlMetadata.WriteCellMetadata(part, context);
        var merges = SpreadsheetXmlMerges.Build(part, metadata.Keys, context, out var rows);
        var columnStyles = new uint?[257];
        foreach (Column definition in part.Worksheet.GetFirstChild<Columns>()?.Elements<Column>() ?? Enumerable.Empty<Column>())
            if (definition.Style != null) for (uint index = definition.Min?.Value ?? 1; index <= Math.Min(256, definition.Max?.Value ?? 256); index++) columnStyles[index] = definition.Style.Value;
        int lastRow = 0;
        foreach (Row row in rows) {
            context.Options.CancellationToken.ThrowIfCancellationRequested();
            int rowIndex = checked((int)(row.RowIndex?.Value ?? (uint)(lastRow + 1)));
            Coordinate(rowIndex, 1);
            if (rowIndex <= lastRow) throw new InvalidDataException("Worksheet row indexes must increase.");
            lastRow = rowIndex;
            writer.WriteStartElement("Row", Ss.NamespaceName); Attribute(writer, "Index", rowIndex.ToString(CultureInfo.InvariantCulture));
            InspectAttributes(row, context, "r", "ht", "customHeight", "hidden", "spans", "s", "customFormat");
            if (row.Hidden?.Value == true) Attribute(writer, "Hidden", "1");
            if (row.Height != null) Attribute(writer, "Height", row.Height.Value.ToString("R", CultureInfo.InvariantCulture));
            if (row.StyleIndex != null) {
                if (row.StyleIndex.Value >= styleCount) throw new InvalidDataException("Invalid row style.");
                Attribute(writer, "StyleID", SpreadsheetXmlStyles.StyleId(row.StyleIndex.Value));
            }
            if (row.OutlineLevel?.Value > 0) context.Omitted("row outlining");
            int lastColumn = 0;
            foreach (Cell cell in row.Elements<Cell>()) {
                InspectAttributes(cell, context, "r", "s", "t");
                int column = lastColumn + 1;
                if (cell.CellReference != null) {
                    ExcelReference reference = ExcelReference.Parse(cell.CellReference.Value!);
                    if (reference.Kind != ExcelReferenceKind.Cell || reference.Start.Row != rowIndex) throw new InvalidDataException("Cell address disagrees with its row.");
                    column = reference.Start.Column;
                }
                Coordinate(rowIndex, column);
                if (column <= lastColumn) throw new InvalidDataException("Worksheet cell indexes must increase.");
                lastColumn = column;
                writer.WriteStartElement("Cell", Ss.NamespaceName); Attribute(writer, "Index", column.ToString(CultureInfo.InvariantCulture));
                uint style = cell.StyleIndex?.Value ?? 0;
                if (style >= styleCount) throw new InvalidDataException("Cell refers to a missing style.");
                if (cell.StyleIndex != null || row.StyleIndex == null && columnStyles[column] == null)
                    Attribute(writer, "StyleID", SpreadsheetXmlStyles.StyleId(style));
                if (cell.CellFormula != null) {
                    InspectAttributes(cell.CellFormula, context, "t");
                    if (cell.CellFormula.FormulaType != null && cell.CellFormula.FormulaType.Value != CellFormulaValues.Normal) {
                        context.Omitted("shared, array or data-table formula structures");
                    } else {
                        string formula = cell.CellFormula.Text;
                        if (formula.Length > 8192) throw new InvalidDataException("Formula exceeds XML Spreadsheet's supported length.");
                        var syntax = ExcelFormulaSyntaxTree.Parse(formula);
                        if (syntax.Nodes.OfType<ExcelFormulaStructuredReferenceSyntax>().Any()) context.Omitted("table-structured formula expressions");
                        else Attribute(writer, "Formula", "=" + syntax.ConvertReferences(ExcelReferenceStyle.R1C1, rowIndex, column));
                    }
                }
                if (merges.TryGetValue(cell.CellReference?.Value ?? ExcelA1.CellReference(rowIndex, column), out var merge)) {
                    if (merge.across > 0) Attribute(writer, "MergeAcross", merge.across.ToString(CultureInfo.InvariantCulture));
                    if (merge.down > 0) Attribute(writer, "MergeDown", merge.down.ToString(CultureInfo.InvariantCulture));
                }
                if (metadata.TryGetValue(cell.CellReference?.Value ?? ExcelA1.CellReference(rowIndex, column), out var extra))
                    foreach (var attribute in extra.Attributes()) writer.WriteAttributeString(attribute.Name.Namespace == X ? "x" : "ss", attribute.Name.LocalName, attribute.Name.NamespaceName, attribute.Value);
                WriteValue(writer, cell, strings, part.GetParentParts().OfType<WorkbookPart>().Single(), context);
                if (extra != null) foreach (var child in extra.Elements()) child.WriteTo(writer);
                writer.WriteEndElement();
            }
            writer.WriteEndElement();
        }
        writer.WriteEndElement();
    }

    private static void WriteValue(XmlWriter writer, Cell cell, SharedStringItem[] strings, WorkbookPart workbook, SpreadsheetXmlContext context) {
        string? value = cell.CellValue?.Text;
        string type = "Number";
        OpenXmlCompositeElement? rich = null;
        if (cell.DataType?.Value == CellValues.SharedString) {
            if (!int.TryParse(value, NumberStyles.None, CultureInfo.InvariantCulture, out int index) || index < 0 || index >= strings.Length) throw new InvalidDataException("Invalid shared string index.");
            value = VisibleText(strings[index], context); rich = strings[index]; type = "String";
        } else if (cell.DataType?.Value == CellValues.InlineString) {
            value = cell.InlineString == null ? string.Empty : VisibleText(cell.InlineString, context); rich = cell.InlineString; type = "String";
        } else if (cell.DataType?.Value == CellValues.String) type = "String";
        else if (cell.DataType?.Value == CellValues.Boolean) type = "Boolean";
        else if (cell.DataType?.Value == CellValues.Error) type = "Error";
        else if (cell.DataType?.Value == CellValues.Date) type = "DateTime";
        if (value == null) return;
        if (value.Length > 32767) throw new InvalidDataException("Cell value exceeds Excel's text limit.");
        if (type == "Number" && (!double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out double number)
                || double.IsNaN(number) || double.IsInfinity(number))) throw new InvalidDataException("Invalid numeric cell value.");
        if (type == "Boolean" && value != "0" && value != "1") throw new InvalidDataException("Invalid Boolean cell value.");
        writer.WriteStartElement("Data", Ss.NamespaceName); Attribute(writer, "Type", type);
        if (type == "String") writer.WriteAttributeString("xml", "space", "http://www.w3.org/XML/1998/namespace", "preserve");
        if (rich?.Elements<Run>().Any() == true) SpreadsheetXmlRichText.Write(writer, rich, workbook, context);
        else writer.WriteString(value);
        writer.WriteEndElement();
    }

    private static string VisibleText(OpenXmlCompositeElement source, SpreadsheetXmlContext context) {
        if (source.Elements<PhoneticRun>().Any()) context.Omitted("phonetic guidance");
        return string.Concat(source.ChildElements.Select(item => item is Text text ? text.Text : item is Run run ? string.Concat(run.Elements<Text>().Select(text => text.Text)) : string.Empty));
    }

    private static void Attribute(XmlWriter writer, string name, string value) => writer.WriteAttributeString("ss", name, Ss.NamespaceName, value);

    private static void InspectAttributes(OpenXmlElement element, SpreadsheetXmlContext context, params string[] supported) {
        foreach (OpenXmlAttribute attribute in element.GetAttributes())
            if (!string.IsNullOrEmpty(attribute.NamespaceUri) || !supported.Contains(attribute.LocalName))
                context.Omitted(element.LocalName + " attribute " + attribute.LocalName);
    }
}

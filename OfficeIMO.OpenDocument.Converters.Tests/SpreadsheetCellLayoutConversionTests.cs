using System;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using OfficeIMO.Excel;
using OfficeIMO.Excel.OpenDocument;
using OfficeIMO.OpenDocument;
using Xunit;

namespace OfficeIMO.OpenDocument.Converters.Tests;

public sealed class SpreadsheetCellLayoutConversionTests {
    [Fact]
    public void ExcelAlignmentAndWrapRoundTripThroughOdsCellStyle() {
        using ExcelDocument source = ExcelDocument.Create();
        ExcelSheet sheet = source.AddWorksheet("Layout");
        sheet.CellAt(1, 1).SetValue("Wrapped text");
        sheet.CellAlign(1, 1, ExcelHorizontalAlignment.Right);
        sheet.CellVerticalAlign(1, 1, ExcelVerticalAlignment.Center);
        sheet.CellWrapText(1, 1);

        OdfConversionResult<OdsDocument> toOds = source.ToOpenDocumentResult();
        Assert.DoesNotContain(toOds.Report.Mappings, mapping => mapping.Feature == "cell-format-details" &&
            mapping.Status == OdfConversionMappingStatus.Unsupported);
        OdsDocument reopened = OdsDocument.Load(new MemoryStream(toOds.Value.ToBytes()));
        Assert.True(reopened.Validate().IsValid);
        OdsCell cell = reopened.GetSheet("Layout")!.Cell(0, 0);
        Assert.Equal("right", cell.TextAlign);
        Assert.Equal("fix", cell.TextAlignSource);
        Assert.Equal("middle", cell.VerticalAlign);
        Assert.Equal("wrap", cell.WrapOption);

        OdfConversionResult<ExcelDocument> toExcel = reopened.ToExcelDocumentResult();
        using ExcelDocument roundTrip = toExcel.Value;
        ExcelCellStyleSnapshot style = roundTrip.CreateInspectionSnapshot().Worksheets.Single()
            .Cells.Single().Style!;
        Assert.Equal("right", style.HorizontalAlignment);
        Assert.Equal("center", style.VerticalAlignment);
        Assert.True(style.WrapText);
        Assert.DoesNotContain(toExcel.Report.Mappings, mapping => mapping.Feature == "cell-layout" &&
            mapping.Status == OdfConversionMappingStatus.Unsupported);
    }

    [Fact]
    public void InheritedOdsCellLayoutProjectsToExcel() {
        OdsDocument source = OdsDocument.Create();
        OdfStyle parent = source.Styles.CreateNamed("LayoutParent", OdfStyleFamily.TableCell);
        parent.TextAlign = "center";
        parent.CellVerticalAlign = "bottom";
        parent.CellWrapOption = "no-wrap";
        OdfStyle child = source.Styles.CreateAutomatic(OdfStyleFamily.TableCell);
        child.ParentStyleName = parent.Name;
        OdsCell cell = source.AddSheet("Layout").Cell(0, 0);
        cell.SetString("Inherited");
        cell.StyleName = child.Name;

        OdsDocument reopened = OdsDocument.Load(new MemoryStream(source.ToBytes()));
        Assert.True(reopened.Validate().IsValid);
        Assert.Equal("center", reopened.GetSheet("Layout")!.Cell(0, 0).TextAlign);
        OdfConversionResult<ExcelDocument> conversion = reopened.ToExcelDocumentResult();
        using ExcelDocument target = conversion.Value;
        ExcelCellStyleSnapshot style = target.CreateInspectionSnapshot().Worksheets.Single()
            .Cells.Single().Style!;
        Assert.Equal("center", style.HorizontalAlignment);
        Assert.Equal("bottom", style.VerticalAlignment);
        Assert.False(style.WrapText);
    }

    [Fact]
    public void UnsupportedAlignmentTokensRemainExplicitLoss() {
        using ExcelDocument excel = ExcelDocument.Create();
        ExcelSheet sheet = excel.AddWorksheet("Layout");
        sheet.CellAt(1, 1).SetValue("Across selection");
        sheet.CellAlign(1, 1, ExcelHorizontalAlignment.CenterContinuous);
        OdfConversionResult<OdsDocument> toOds = excel.ToOpenDocumentResult();
        Assert.Contains(toOds.Report.Mappings, mapping => mapping.Feature == "cell-format-details" &&
            mapping.Status == OdfConversionMappingStatus.Unsupported);

        OdsDocument ods = OdsDocument.Create();
        OdsCell cell = ods.AddSheet("Layout").Cell(0, 0);
        cell.SetString("Direction-sensitive");
        cell.TextAlign = "start";
        OdfConversionResult<ExcelDocument> toExcel = ods.ToExcelDocumentResult();
        using ExcelDocument target = toExcel.Value;
        Assert.Contains(toExcel.Report.Mappings, mapping => mapping.Feature == "cell-layout" &&
            mapping.Status == OdfConversionMappingStatus.Unsupported);
        Assert.Throws<OdfConversionLossException>(() => ods.ToExcelDocumentResult(
            new ExcelOpenDocumentConversionOptions {
                LossPolicy = OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported
            }));
    }

    [Fact]
    public void RowAndColumnDefaultCellStylesProjectWithRowPrecedence() {
        OdsDocument source = OdsDocument.Create();
        OdfStyle columnStyle = source.Styles.CreateNamed("ColumnLayout", OdfStyleFamily.TableCell);
        columnStyle.TextAlign = "right";
        columnStyle.CellWrapOption = "wrap";
        OdfStyle rowStyle = source.Styles.CreateNamed("RowLayout", OdfStyleFamily.TableCell);
        rowStyle.TextAlign = "center";
        rowStyle.CellVerticalAlign = "middle";

        OdsSheet sheet = source.AddSheet("Layout");
        sheet.Column(0).DefaultCellStyleName = columnStyle.Name;
        sheet.Row(0).DefaultCellStyleName = rowStyle.Name;
        sheet.Cell(0, 0).SetString("Row wins");
        sheet.Cell(1, 0).SetString("Column applies");
        Assert.Null(sheet.Cell(0, 0).StyleName);
        Assert.Null(sheet.Cell(1, 0).StyleName);

        OdsDocument reopened = OdsDocument.Load(new MemoryStream(source.ToBytes()));
        OdsSheet loaded = reopened.GetSheet("Layout")!;
        Assert.True(reopened.Validate().IsValid);
        Assert.Equal("center", loaded.Cell(0, 0).TextAlign);
        Assert.Equal("right", loaded.Cell(1, 0).TextAlign);
        Assert.Equal("wrap", loaded.Cell(1, 0).WrapOption);
        OdfConversionResult<ExcelDocument> conversion = reopened.ToExcelDocumentResult();
        using ExcelDocument target = conversion.Value;
        var cells = target.CreateInspectionSnapshot().Worksheets.Single().Cells;
        Assert.Equal("center", cells.Single(cell => cell.Row == 1).Style!.HorizontalAlignment);
        Assert.Equal("center", cells.Single(cell => cell.Row == 1).Style!.VerticalAlignment);
        Assert.Equal("right", cells.Single(cell => cell.Row == 2).Style!.HorizontalAlignment);
        Assert.True(cells.Single(cell => cell.Row == 2).Style!.WrapText);
        Assert.DoesNotContain(conversion.Report.Mappings, mapping => mapping.Feature == "cell-layout" &&
            mapping.Status == OdfConversionMappingStatus.Unsupported);
    }

    [Fact]
    public void ValueTypeAlignmentDoesNotBecomeFixedExcelAlignment() {
        OdsDocument source = OdsDocument.Create();
        OdfStyle style = source.Styles.CreateNamed("ValueAlignment", OdfStyleFamily.TableCell);
        style.TextAlign = "center";
        style.CellTextAlignSource = "value-type";
        OdsCell cell = source.AddSheet("Layout").Cell(0, 0);
        cell.SetString("Text");
        cell.StyleName = style.Name;

        OdfConversionResult<ExcelDocument> conversion = source.ToExcelDocumentResult();
        using ExcelDocument target = conversion.Value;
        ExcelCellStyleSnapshot? projected = target.CreateInspectionSnapshot().Worksheets.Single().Cells.Single().Style;
        Assert.NotEqual("center", projected?.HorizontalAlignment);
        Assert.Contains(conversion.Report.Mappings, mapping => mapping.Feature == "cell-layout" &&
            mapping.Status == OdfConversionMappingStatus.Unsupported);
    }

    [Fact]
    public void OmittedAlignmentSourceDefaultsToValueTypeAndReportsLoss() {
        OdsDocument source = OdsDocument.Create();
        OdfStyle style = source.Styles.CreateNamed("OmittedSource", OdfStyleFamily.TableCell);
        style.TextAlign = "center";
        style.CellTextAlignSource = null;
        OdsCell cell = source.AddSheet("Layout").Cell(0, 0);
        cell.SetNumber(42);
        cell.StyleName = style.Name;

        OdfConversionResult<ExcelDocument> conversion = source.ToExcelDocumentResult();
        using ExcelDocument target = conversion.Value;
        ExcelCellStyleSnapshot? projected = target.CreateInspectionSnapshot().Worksheets.Single().Cells.Single().Style;
        Assert.NotEqual("center", projected?.HorizontalAlignment);
        Assert.Contains(conversion.Report.Mappings, mapping => mapping.Feature == "cell-layout" &&
            mapping.Status == OdfConversionMappingStatus.Unsupported);
        Assert.Throws<OdfConversionLossException>(() => source.ToExcelDocumentResult(
            new ExcelOpenDocumentConversionOptions {
                LossPolicy = OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported
            }));
    }

    [Fact]
    public void RepeatedCellRunUsesEachColumnsDefaultStyle() {
        OdsDocument source = OdsDocument.Create();
        OdfStyle left = source.Styles.CreateNamed("LeftColumn", OdfStyleFamily.TableCell);
        left.TextAlign = "left";
        OdfStyle right = source.Styles.CreateNamed("RightColumn", OdfStyleFamily.TableCell);
        right.TextAlign = "right";
        OdsSheet sheet = source.AddSheet("Layout");
        sheet.Column(0).DefaultCellStyleName = left.Name;
        sheet.Column(1).DefaultCellStyleName = right.Name;
        OdsCell cell = sheet.Cell(0, 0);
        cell.SetString("Repeated");
        cell.Element.SetAttributeValue(OdfNamespaces.Table + "number-columns-repeated", "2");
        source.MarkPartDirty("content.xml");

        OdsDocument reopened = OdsDocument.Load(new MemoryStream(source.ToBytes()));
        OdfConversionResult<ExcelDocument> conversion = reopened.ToExcelDocumentResult();
        using ExcelDocument target = conversion.Value;
        var cells = target.CreateInspectionSnapshot().Worksheets.Single().Cells;
        Assert.Equal("left", cells.Single(item => item.Column == 1).Style!.HorizontalAlignment);
        Assert.Equal("right", cells.Single(item => item.Column == 2).Style!.HorizontalAlignment);
    }

    [Fact]
    public void DefaultTableCellStyleProjectsWhenCellRowAndColumnHaveNoStyle() {
        OdsDocument source = OdsDocument.Create();
        XElement styles = source.GetXml("styles.xml").Root!.Element(OdfNamespaces.Office + "styles")!;
        styles.Add(new XElement(OdfNamespaces.Style + "default-style",
            new XAttribute(OdfNamespaces.Style + "family", "table-cell"),
            new XElement(OdfNamespaces.Style + "paragraph-properties",
                new XAttribute(OdfNamespaces.Fo + "text-align", "right")),
            new XElement(OdfNamespaces.Style + "table-cell-properties",
                new XAttribute(OdfNamespaces.Style + "text-align-source", "fix"),
                new XAttribute(OdfNamespaces.Fo + "wrap-option", "wrap"))));
        source.MarkPartDirty("styles.xml");
        source.AddSheet("Layout").Cell(0, 0).SetString("Default");

        OdsDocument reopened = OdsDocument.Load(new MemoryStream(source.ToBytes()));
        Assert.True(reopened.Validate().IsValid);
        OdfConversionResult<ExcelDocument> conversion = reopened.ToExcelDocumentResult();
        using ExcelDocument target = conversion.Value;
        ExcelCellStyleSnapshot style = target.CreateInspectionSnapshot().Worksheets.Single().Cells.Single().Style!;
        Assert.Equal("right", style.HorizontalAlignment);
        Assert.True(style.WrapText);
    }

    [Fact]
    public void BlankCellsWithInheritedStylesReportExplicitLoss() {
        OdsDocument source = OdsDocument.Create();
        OdfStyle style = source.Styles.CreateNamed("BlankLayout", OdfStyleFamily.TableCell);
        style.BackgroundColor = OdfColor.Parse("#FFCC00");
        OdsSheet sheet = source.AddSheet("Layout");
        sheet.Row(0).DefaultCellStyleName = style.Name;
        sheet.Cell(0, 0);
        sheet.Column(1).DefaultCellStyleName = style.Name;
        sheet.Cell(1, 0).SetString("Value");
        sheet.Cell(1, 1);

        OdfConversionResult<ExcelDocument> conversion = source.ToExcelDocumentResult();
        using ExcelDocument target = conversion.Value;
        Assert.Contains(conversion.Report.Mappings, mapping => mapping.Feature == "blank-cell-styles" &&
            mapping.Status == OdfConversionMappingStatus.Unsupported);
        Assert.Throws<OdfConversionLossException>(() => source.ToExcelDocumentResult(
            new ExcelOpenDocumentConversionOptions {
                LossPolicy = OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported
            }));
    }

    [Fact]
    public void RepeatedRowsRetainInheritedBlankStyleLossCount() {
        OdsDocument source = OdsDocument.Create();
        OdfStyle style = source.Styles.CreateNamed("BlankLayout", OdfStyleFamily.TableCell);
        style.BackgroundColor = OdfColor.Parse("#FFCC00");
        OdsSheet sheet = source.AddSheet("Layout");
        sheet.Column(1).DefaultCellStyleName = style.Name;
        sheet.Cell(0, 0).SetString("Value");
        sheet.Cell(0, 1);
        sheet.Element.Elements(OdfNamespaces.Table + "table-row").Single()
            .SetAttributeValue(OdfNamespaces.Table + "number-rows-repeated", "128");

        OdfConversionResult<ExcelDocument> conversion = source.ToExcelDocumentResult();
        using ExcelDocument target = conversion.Value;
        Assert.Contains(conversion.Report.ForFeature("blank-cell-styles"), mapping =>
            mapping.Status == OdfConversionMappingStatus.Unsupported && mapping.Count == 128);
    }

    [Fact]
    public void EmptyFamilyDefaultDoesNotReportBlankCellStyleLoss() {
        OdsDocument source = OdsDocument.Create();
        XElement styles = source.Package.GetXml("styles.xml").Root!
            .Element(OdfNamespaces.Office + "styles")!;
        styles.Add(new XElement(OdfNamespaces.Style + "default-style",
            new XAttribute(OdfNamespaces.Style + "family", "table-cell")));
        source.Package.MarkXmlDirty("styles.xml");
        source.AddSheet("Layout").Cell(0, 0).SetString("Value");

        OdfConversionResult<ExcelDocument> conversion = source.ToExcelDocumentResult();
        using ExcelDocument target = conversion.Value;
        Assert.DoesNotContain(conversion.Report.ForFeature("blank-cell-styles"), mapping =>
            mapping.Status == OdfConversionMappingStatus.Unsupported);
    }

    [Fact]
    public void EmptyNamedRowAndColumnDefaultsDoNotReportBlankCellStyleLoss() {
        OdsDocument source = OdsDocument.Create();
        OdfStyle empty = source.Styles.CreateNamed("EmptyCell", OdfStyleFamily.TableCell);
        OdsSheet sheet = source.AddSheet("Layout");
        sheet.Row(0).DefaultCellStyleName = empty.Name;
        sheet.Column(1).DefaultCellStyleName = empty.Name;
        sheet.Cell(0, 0).SetString("Value");

        OdfConversionResult<ExcelDocument> conversion = source.ToExcelDocumentResult();
        using ExcelDocument target = conversion.Value;
        Assert.DoesNotContain(conversion.Report.ForFeature("blank-cell-styles"), mapping =>
            mapping.Status == OdfConversionMappingStatus.Unsupported);
    }

    [Fact]
    public void TextOnlyFamilyDefaultStylesPopulatedCellsWithoutBlankCellLoss() {
        OdsDocument source = OdsDocument.Create();
        XElement styles = source.Package.GetXml("styles.xml").Root!
            .Element(OdfNamespaces.Office + "styles")!;
        styles.Add(new XElement(OdfNamespaces.Style + "default-style",
            new XAttribute(OdfNamespaces.Style + "family", "table-cell"),
            new XElement(OdfNamespaces.Style + "text-properties",
                new XAttribute(OdfNamespaces.Fo + "color", "#336699"))));
        source.Package.MarkXmlDirty("styles.xml");
        source.AddSheet("Layout").Cell(0, 0).SetString("Styled");

        OdfConversionResult<ExcelDocument> conversion = source.ToExcelDocumentResult();
        using ExcelDocument target = conversion.Value;
        ExcelCellStyleSnapshot style = target.CreateInspectionSnapshot().Worksheets.Single().Cells.Single().Style!;
        Assert.Equal("336699", style.FontColorHex);
        Assert.DoesNotContain(conversion.Report.ForFeature("blank-cell-styles"), mapping =>
            mapping.Status == OdfConversionMappingStatus.Unsupported);
    }

    [Fact]
    public void ColumnDefaultBeyondLastSerializedCellReportsBlankStyleLoss() {
        OdsDocument source = OdsDocument.Create();
        OdfStyle style = source.Styles.CreateNamed("SparseColumn", OdfStyleFamily.TableCell);
        style.BackgroundColor = OdfColor.Parse("#FFCC00");
        OdsSheet sheet = source.AddSheet("Layout");
        sheet.Cell(0, 0).SetString("Value");
        sheet.Column(4).DefaultCellStyleName = style.Name;

        OdfConversionResult<ExcelDocument> conversion = source.ToExcelDocumentResult();
        using ExcelDocument target = conversion.Value;
        Assert.Contains(conversion.Report.Mappings, mapping => mapping.Feature == "blank-cell-styles" &&
            mapping.Status == OdfConversionMappingStatus.Unsupported);
        Assert.Throws<OdfConversionLossException>(() => source.ToExcelDocumentResult(
            new ExcelOpenDocumentConversionOptions {
                LossPolicy = OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported
            }));
    }

    [Fact]
    public void CollapsedOdsRowAndColumnGroupsStayHiddenAndReportStructureLoss() {
        OdsDocument source = OdsDocument.Create();
        OdsSheet sheet = source.AddSheet("Grouped");
        sheet.Column(0);
        sheet.Cell(0, 0).SetNumber(42);
        XElement column = sheet.Element.Elements(OdfNamespaces.Table + "table-column").Single();
        column.ReplaceWith(new XElement(OdfNamespaces.Table + "table-column-group",
            new XAttribute(OdfNamespaces.Table + "display", "false"), new XElement(column)));
        XElement row = sheet.Element.Elements(OdfNamespaces.Table + "table-row").Single();
        row.ReplaceWith(new XElement(OdfNamespaces.Table + "table-row-group",
            new XAttribute(OdfNamespaces.Table + "display", "false"), new XElement(row)));
        source.MarkPartDirty("content.xml");

        OdsDocument reopened = OdsDocument.Load(new MemoryStream(source.ToBytes()));
        OdsSheet grouped = reopened.GetSheet("Grouped")!;
        Assert.True(grouped.RowRuns.Single().Hidden);
        Assert.True(grouped.ColumnRuns.Single().Hidden);
        Assert.True(grouped.Row(0).Hidden);
        Assert.True(grouped.Column(0).Hidden);
        Assert.Throws<InvalidOperationException>(() => grouped.Row(0).Hidden = false);
        Assert.Throws<InvalidOperationException>(() => grouped.Column(0).Hidden = false);
        OdfConversionResult<ExcelDocument> conversion = reopened.ToExcelDocumentResult();
        using ExcelDocument target = conversion.Value;
        ExcelWorksheetSnapshot snapshot = target.CreateInspectionSnapshot().Worksheets.Single();
        Assert.True(snapshot.Rows.Single(rowSnapshot => rowSnapshot.Index == 1).Hidden);
        Assert.True(snapshot.Columns.Single(columnSnapshot => columnSnapshot.StartIndex == 1).Hidden);
        Assert.Contains(conversion.Report.Mappings, mapping => mapping.Feature == "row-groups" &&
            mapping.Status == OdfConversionMappingStatus.Unsupported);
        Assert.Contains(conversion.Report.Mappings, mapping => mapping.Feature == "column-groups" &&
            mapping.Status == OdfConversionMappingStatus.Unsupported);
        Assert.Throws<OdfConversionLossException>(() => reopened.ToExcelDocumentResult(
            new ExcelOpenDocumentConversionOptions {
                LossPolicy = OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported
            }));
    }

    [Fact]
    public void FixedAlignmentWithoutTextAlignReportsDefaultStartLoss() {
        OdsDocument source = OdsDocument.Create();
        OdfStyle style = source.Styles.CreateNamed("FixedStart", OdfStyleFamily.TableCell);
        style.CellTextAlignSource = "fix";
        OdsCell cell = source.AddSheet("Layout").Cell(0, 0);
        cell.SetNumber(42);
        cell.StyleName = style.Name;

        OdfConversionResult<ExcelDocument> conversion = source.ToExcelDocumentResult();
        using ExcelDocument target = conversion.Value;
        Assert.Contains(conversion.Report.Mappings, mapping => mapping.Feature == "cell-layout" &&
            mapping.Status == OdfConversionMappingStatus.Unsupported);
        Assert.Throws<OdfConversionLossException>(() => source.ToExcelDocumentResult(
            new ExcelOpenDocumentConversionOptions {
                LossPolicy = OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported
            }));
    }

    [Fact]
    public void FilterHiddenOdsRowsAndColumnsStayHiddenWithExplicitLoss() {
        OdsDocument source = OdsDocument.Create();
        OdsSheet sheet = source.AddSheet("Filtered");
        sheet.Column(0);
        sheet.Cell(0, 0).SetNumber(42);
        sheet.Element.Elements(OdfNamespaces.Table + "table-column").Single()
            .SetAttributeValue(OdfNamespaces.Table + "visibility", "filter");
        sheet.Element.Elements(OdfNamespaces.Table + "table-row").Single()
            .SetAttributeValue(OdfNamespaces.Table + "visibility", "filter");
        source.MarkPartDirty("content.xml");

        OdsDocument reopened = OdsDocument.Load(new MemoryStream(source.ToBytes()));
        OdsSheet filtered = reopened.GetSheet("Filtered")!;
        Assert.True(filtered.RowRuns.Single().Hidden);
        Assert.True(filtered.ColumnRuns.Single().Hidden);
        Assert.True(filtered.Row(0).Hidden);
        Assert.True(filtered.Column(0).Hidden);
        OdfConversionResult<ExcelDocument> conversion = reopened.ToExcelDocumentResult();
        using ExcelDocument target = conversion.Value;
        ExcelWorksheetSnapshot snapshot = target.CreateInspectionSnapshot().Worksheets.Single();
        Assert.True(snapshot.Rows.Single(rowSnapshot => rowSnapshot.Index == 1).Hidden);
        Assert.True(snapshot.Columns.Single(columnSnapshot => columnSnapshot.StartIndex == 1).Hidden);
        Assert.Contains(conversion.Report.Mappings, mapping => mapping.Feature == "filtered-visibility" &&
            mapping.Status == OdfConversionMappingStatus.Unsupported);
        Assert.Throws<OdfConversionLossException>(() => reopened.ToExcelDocumentResult(
            new ExcelOpenDocumentConversionOptions {
                LossPolicy = OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported
            }));
    }
}

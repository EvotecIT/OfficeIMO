using System;
using System.Linq;
using System.Xml.Linq;
using Xunit;

namespace OfficeIMO.OpenDocument.Tests;

public sealed class OdsCellStyleInheritanceTests {
    [Theory]
    [InlineData("table-column-group")]
    [InlineData("table-header-columns")]
    public void EditingNestedRepeatedColumnUsesItsLogicalPosition(string wrapperName) {
        OdsDocument document = OdsDocument.Create();
        OdsSheet sheet = document.AddSheet("Data");
        sheet.Cell(0, 0).SetString("Value");
        XElement table = document.Package.GetXml("content.xml")
            .Descendants(OdfNamespaces.Table + "table").Single();
        table.Elements(OdfNamespaces.Table + "table-column").Remove();
        table.AddFirst(new XElement(OdfNamespaces.Table + wrapperName,
            new XElement(OdfNamespaces.Table + "table-column",
                new XAttribute(OdfNamespaces.Table + "number-columns-repeated", 3))));
        document.MarkPartDirty("content.xml");

        Assert.Single(sheet.ColumnRuns);
        Assert.Equal(3, sheet.ColumnRuns[0].RepeatCount);
        sheet.Column(1).Hidden = true;

        Assert.Empty(table.Elements(OdfNamespaces.Table + "table-column"));
        Assert.Equal(new long[] { 0, 1, 2 }, sheet.ColumnRuns.Select(run => run.StartColumn));
        Assert.Equal(new bool[] { false, true, false }, sheet.ColumnRuns.Select(run => run.Hidden));
        OdsSheet reopened = OdsDocument.Load(new System.IO.MemoryStream(document.ToBytes())).Sheets.Single();
        Assert.Equal(new bool[] { false, true, false }, reopened.ColumnRuns.Select(run => run.Hidden));
    }

    [Fact]
    public void RetainedCellReadsCurrentRowAndColumnDefaultStyles() {
        OdsDocument document = OdsDocument.Create();
        OdfStyle columnStyle = document.Styles.CreateNamed("ColumnDefault", OdfStyleFamily.TableCell);
        columnStyle.Bold = true;
        OdfStyle rowStyle = document.Styles.CreateNamed("RowDefault", OdfStyleFamily.TableCell);
        rowStyle.Bold = false;
        OdsSheet sheet = document.AddSheet("Data");
        OdsCell cell = sheet.Cell(0, 0);
        cell.SetString("Value");

        sheet.Column(0).DefaultCellStyleName = columnStyle.Name;
        Assert.True(cell.Bold);
        sheet.Row(0).DefaultCellStyleName = rowStyle.Name;
        Assert.False(cell.Bold);
        sheet.Row(0).DefaultCellStyleName = null;
        Assert.True(cell.Bold);
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void EditingInheritedAutomaticStyleDoesNotChangeSiblingCells(bool rowDefault) {
        OdsDocument document = OdsDocument.Create();
        OdfStyle shared = document.Styles.CreateAutomatic(OdfStyleFamily.TableCell);
        shared.Bold = true;
        OdsSheet sheet = document.AddSheet("Data");
        if (rowDefault) sheet.Row(0).DefaultCellStyleName = shared.Name;
        else sheet.Column(0).DefaultCellStyleName = shared.Name;
        OdsCell edited = sheet.Cell(0, 0);
        edited.SetString("Edited");
        OdsCell sibling = rowDefault ? sheet.Cell(0, 1) : sheet.Cell(1, 0);
        sibling.SetString("Sibling");

        edited.Italic = true;

        Assert.NotEqual(shared.Name, edited.StyleName);
        Assert.True(edited.Bold);
        Assert.True(edited.Italic);
        Assert.True(sibling.Bold);
        Assert.Null(sibling.Italic);
        Assert.Null(shared.Italic);
    }

    [Fact]
    public void FamilyDefaultPropertiesFillGapsInNamedCellStyle() {
        OdsDocument document = OdsDocument.Create();
        XElement styles = document.GetXml("styles.xml").Root!.Element(OdfNamespaces.Office + "styles")!;
        styles.Add(new XElement(OdfNamespaces.Style + "default-style",
            new XAttribute(OdfNamespaces.Style + "family", "table-cell"),
            new XElement(OdfNamespaces.Style + "text-properties",
                new XAttribute(OdfNamespaces.Fo + "font-weight", "bold"),
                new XAttribute(OdfNamespaces.Fo + "color", "#123456"),
                new XAttribute(OdfNamespaces.Fo + "font-size", "13pt")),
            new XElement(OdfNamespaces.Style + "table-cell-properties",
                new XAttribute(OdfNamespaces.Fo + "background-color", "#FFCC00"))));
        document.MarkPartDirty("styles.xml");
        OdfStyle named = document.Styles.CreateNamed("AlignmentOnly", OdfStyleFamily.TableCell);
        named.TextAlign = "right";
        OdsCell cell = document.AddSheet("Data").Cell(0, 0);
        cell.SetString("Value");
        cell.StyleName = named.Name;

        Assert.True(cell.Bold);
        Assert.Equal("#123456", cell.Color?.ToString());
        Assert.Equal(OdfLength.Points(13), cell.FontSize);
        Assert.Equal("#FFCC00", cell.BackgroundColor?.ToString());
    }

    [Fact]
    public void PublicSparseCellRunsExposeCurrentInheritedStyle() {
        OdsDocument document = OdsDocument.Create();
        OdfStyle rowStyle = document.Styles.CreateNamed("RowStyle", OdfStyleFamily.TableCell);
        OdfStyle columnStyle = document.Styles.CreateNamed("ColumnStyle", OdfStyleFamily.TableCell);
        OdsSheet sheet = document.AddSheet("Data");
        sheet.Cell(0, 0).SetString("First");
        sheet.Cell(1, 0).SetString("Second");
        OdsRowRun firstRow = sheet.RowRuns.Single(run => run.StartRow == 0);
        OdsRowRun secondRow = sheet.RowRuns.Single(run => run.StartRow == 1);
        OdsCellRun firstCell = firstRow.CellRuns.Single();
        OdsCellRun secondCell = secondRow.CellRuns.Single();

        sheet.Row(0).DefaultCellStyleName = rowStyle.Name;
        sheet.Column(0).DefaultCellStyleName = columnStyle.Name;

        Assert.Equal(rowStyle.Name, firstCell.EffectiveStyleName);
        Assert.Equal(columnStyle.Name, secondCell.EffectiveStyleName);
    }

    [Fact]
    public void RepeatedPublicCellRunSplitsAtColumnStyleBoundaries() {
        OdsDocument document = OdsDocument.Create();
        OdfStyle left = document.Styles.CreateNamed("Left", OdfStyleFamily.TableCell);
        left.TextAlign = "left";
        OdfStyle right = document.Styles.CreateNamed("Right", OdfStyleFamily.TableCell);
        right.TextAlign = "right";
        OdsSheet sheet = document.AddSheet("Data");
        sheet.Column(0).DefaultCellStyleName = left.Name;
        sheet.Column(1).DefaultCellStyleName = right.Name;
        sheet.Cell(0, 0).SetString("Same");
        XElement cell = document.Package.GetXml("content.xml")
            .Descendants(OdfNamespaces.Table + "table-cell").Single();
        cell.SetAttributeValue(OdfNamespaces.Table + "number-columns-repeated", 2);
        document.MarkPartDirty("content.xml");

        OdsCellRun[] runs = sheet.RowRuns.Single().CellRuns.ToArray();
        Assert.Equal(new long[] { 0, 1 }, runs.Select(run => run.StartColumn));
        Assert.All(runs, run => Assert.Equal(1, run.RepeatCount));
        Assert.Equal(new[] { left.Name, right.Name }, runs.Select(run => run.EffectiveStyleName));
        Assert.Equal(new[] { "left", "right" }, runs.Select(run => run.TextAlign));
        Assert.Single(document.Package.GetXml("content.xml").Descendants(OdfNamespaces.Table + "table-cell"));
    }

    [Fact]
    public void ExplicitValueTypeAlignmentSurvivesEitherSetterOrder() {
        OdsDocument document = OdsDocument.Create();
        OdfStyle first = document.Styles.CreateNamed("First", OdfStyleFamily.TableCell);
        first.CellTextAlignSource = "value-type";
        first.TextAlign = "right";
        OdfStyle second = document.Styles.CreateNamed("Second", OdfStyleFamily.TableCell);
        second.TextAlign = "right";
        second.CellTextAlignSource = "value-type";
        OdfStyle fixedStyle = document.Styles.CreateNamed("Fixed", OdfStyleFamily.TableCell);
        fixedStyle.TextAlign = "center";

        Assert.Equal("value-type", first.CellTextAlignSource);
        Assert.Equal("value-type", second.CellTextAlignSource);
        Assert.Equal("fix", fixedStyle.CellTextAlignSource);
    }
}

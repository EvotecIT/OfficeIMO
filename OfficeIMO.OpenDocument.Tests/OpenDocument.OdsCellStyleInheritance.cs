using System;
using System.Linq;
using System.Xml.Linq;
using Xunit;

namespace OfficeIMO.OpenDocument.Tests;

public sealed class OdsCellStyleInheritanceTests {
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
}

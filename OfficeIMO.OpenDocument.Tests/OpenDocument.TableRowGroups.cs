using System;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using Xunit;

namespace OfficeIMO.OpenDocument.Tests;

public sealed class OdfTableRowGroupTests {
    [Theory]
    [InlineData("table-row-group")]
    [InlineData("table-rows")]
    public void GroupedRowsKeepLogicalPositionsAcrossSparseEdits(string containerName) {
        OdsDocument document = OdsDocument.Create();
        OdsSheet sheet = document.AddSheet("Data");
        sheet.Cell(0, 0).SetString("Grouped");
        XElement table = document.Package.GetXml("content.xml")
            .Descendants(OdfNamespaces.Table + "table").Single();
        XElement original = table.Elements(OdfNamespaces.Table + "table-row").Single();
        XElement grouped = new XElement(original);
        grouped.SetAttributeValue(OdfNamespaces.Table + "number-rows-repeated", 2);
        XElement tail = new XElement(original);
        tail.Descendants(OdfNamespaces.Text + "p").Single().Value = "Tail";
        original.ReplaceWith(new XElement(OdfNamespaces.Table + containerName, grouped), tail);
        document.MarkPartDirty("content.xml");

        Assert.Equal(3, sheet.RowCount);
        Assert.Equal("Grouped", sheet.GetValue(0, 0).DisplayText);
        Assert.Equal("Tail", sheet.GetValue(2, 0).DisplayText);
        sheet.Cell(1, 0).SetString("Edited");
        sheet.Cell(2, 0).SetString("Tail2");

        Assert.Equal(new[] { "Grouped", "Edited", "Tail2" },
            Enumerable.Range(0, 3).Select(row => sheet.GetValue(row, 0).DisplayText));
        OdsSheet reopened = OdsDocument.Load(new MemoryStream(document.ToBytes())).Sheets.Single();
        Assert.Equal(new[] { "Grouped", "Edited", "Tail2" },
            Enumerable.Range(0, 3).Select(row => reopened.GetValue(row, 0).DisplayText));
    }

    [Theory]
    [InlineData("table-row-group", true)]
    [InlineData("table-rows", true)]
    [InlineData("table-row-group", false)]
    public void NewColumnsPrecedeRootRowContainers(string containerName, bool hasRow) {
        OdsDocument document = OdsDocument.Create();
        OdsSheet sheet = document.AddSheet("Data");
        XElement table = document.Package.GetXml("content.xml")
            .Descendants(OdfNamespaces.Table + "table").Single();
        XElement original = table.Elements(OdfNamespaces.Table + "table-row").Single();
        original.ReplaceWith(new XElement(OdfNamespaces.Table + containerName,
            hasRow ? new XElement(original) : null));
        document.MarkPartDirty("content.xml");

        sheet.Column(1).Hidden = true;

        Assert.Equal(new[] { "table-column", "table-column", containerName },
            table.Elements().Select(element => element.Name.LocalName));
    }

    [Fact]
    public void NumericFalseGroupDisplayHidesRowsAndColumns() {
        OdsDocument document = OdsDocument.Create();
        OdsSheet sheet = document.AddSheet("Data");
        XElement table = document.Package.GetXml("content.xml")
            .Descendants(OdfNamespaces.Table + "table").Single();
        XElement column = table.Elements(OdfNamespaces.Table + "table-column").Single();
        column.ReplaceWith(new XElement(OdfNamespaces.Table + "table-column-group",
            new XAttribute(OdfNamespaces.Table + "display", "0"), new XElement(column)));
        XElement row = table.Elements(OdfNamespaces.Table + "table-row").Single();
        row.ReplaceWith(new XElement(OdfNamespaces.Table + "table-row-group",
            new XAttribute(OdfNamespaces.Table + "display", "0"), new XElement(row)));
        document.MarkPartDirty("content.xml");

        Assert.True(sheet.Row(0).Hidden);
        Assert.True(sheet.Column(0).Hidden);
        Assert.True(sheet.RowRuns.Single().Hidden);
        Assert.True(sheet.ColumnRuns.Single().Hidden);
        Assert.Throws<InvalidOperationException>(() => sheet.Row(0).Hidden = false);
        Assert.Throws<InvalidOperationException>(() => sheet.Column(0).Hidden = false);
    }
}

using System;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using OfficeIMO.OpenDocument;
using Xunit;

namespace OfficeIMO.OpenDocument.Tests;

public sealed class OpenDocumentOdsDataPilotTests {
    [Fact]
    public void ReadsExcelProducedPivotAndPreservesItsPackage() {
        string path = Path.Combine(AppContext.BaseDirectory, "Fixtures", "microsoft-excel-pivot.ods");
        OdsDocument document = OdsDocument.Load(path);
        byte[] content = document.GetPackageEntryBytes("content.xml");

        OdsDataPilotTable pivot = Assert.Single(document.DataPilotTables);
        Assert.Equal("SalesPivot", pivot.Name);
        Assert.Equal("Data.A1:Data.C5", pivot.SourceRangeAddress);
        Assert.Equal("Data.E1:Data.H5", pivot.TargetRangeAddress);
        Assert.Equal(new[] { "Month", "Region", "Sales" }, pivot.Fields.Select(item => item.SourceFieldName));
        Assert.Equal(new[] { "column", "row", "data" }, pivot.Fields.Select(item => item.Orientation));
        Assert.Equal("sum", pivot.Fields[2].Function);
        Assert.Contains(document.InspectFeatures().Findings,
            finding => finding.Name == "spreadsheet-data-pilot-tables" && finding.Count == 1);
        Assert.True(document.Validate().IsValid);

        OdsDocument reopened = OdsDocument.Load(new MemoryStream(document.ToBytes()));
        Assert.Single(reopened.DataPilotTables);
        Assert.Equal(content, reopened.GetPackageEntryBytes("content.xml"));
    }

    [Fact]
    public void AuthorsLocalRangePivotAndReopensWithFields() {
        OdsDocument document = OdsDocument.Create();
        OdsSheet sheet = document.AddSheet("Data");
        sheet.Cell(0, 0).SetString("Region");
        sheet.Cell(0, 1).SetString("Sales");
        sheet.Cell(1, 0).SetString("North");
        sheet.Cell(1, 1).SetNumber(10);
        OdsDataPilotTable pivot = document.AddDataPilotTable("SalesPivot", "Data.A1:Data.B2", "Data.D1:Data.E3");
        pivot.AddField("Region", "row");
        pivot.AddField("Sales", "data", "sum");

        Assert.Contains(document.InspectFeatures().Findings,
            finding => finding.Name == "spreadsheet-data-pilot-tables"
                && finding.Support == OdfFeatureSupport.Editable && finding.Count == 1);

        OdsDocument reopened = OdsDocument.Load(new MemoryStream(document.ToBytes()));
        OdsDataPilotTable actual = Assert.Single(reopened.DataPilotTables);
        Assert.Equal("SalesPivot", actual.Name);
        Assert.Equal("Data.A1:Data.B2", actual.SourceRangeAddress);
        Assert.Equal("Data.D1:Data.E3", actual.TargetRangeAddress);
        Assert.Equal("sum", actual.Fields[1].Function);
        Assert.True(reopened.Validate().IsValid);

        actual.Element.Add(new System.Xml.Linq.XElement(OdfNamespaces.Table + "data-pilot-source"));
        reopened.MarkPartDirty("content.xml");
        Assert.Contains(reopened.InspectFeatures().Findings,
            finding => finding.Name == "spreadsheet-data-pilot-tables"
                && finding.Support == OdfFeatureSupport.Inspected && finding.Count == 1);
    }

    [Fact]
    public void RejectsInvalidAddressesAndFieldSemanticsBeforeMutation() {
        OdsDocument document = OdsDocument.Create();
        document.AddSheet("Data");
        Assert.Throws<ArgumentException>(() => document.AddDataPilotTable("Bad", "Missing.A1:Missing.B2", "Data.D1:Data.E3"));
        Assert.Throws<ArgumentException>(() => document.AddDataPilotTable("Bad", "Data.A1:Data.B2", "Data.D1:Other.E3"));
        Assert.Empty(document.DataPilotTables);
        OdsDataPilotTable pivot = document.AddDataPilotTable("Good", "Data.A1:Data.B2", "Data.D1:Data.E3");
        Assert.Throws<ArgumentException>(() => pivot.AddField("Sales", "data", "auto"));
        Assert.Throws<ArgumentException>(() => pivot.AddField("Sales", "unknown"));
        Assert.Empty(pivot.Fields);
    }

    [Fact]
    public void PivotAndNamedExpressionsPrecedeExistingSpreadsheetTail() {
        OdsDocument document = OdsDocument.Create();
        OdsSheet sheet = document.AddSheet("Data");
        sheet.Cell(0, 0).SetString("Region");
        sheet.Cell(1, 0).SetString("North");
        XElement spreadsheet = document.Package.GetXml("content.xml")
            .Descendants(OdfNamespaces.Office + "spreadsheet").Single();
        spreadsheet.Add(new XElement(OdfNamespaces.Table + "consolidation"),
            new XElement(OdfNamespaces.Table + "dde-links"));
        document.Package.MarkXmlDirty("content.xml");

        OdsDataPilotTable pivot = document.AddDataPilotTable("Pivot", "Data.A1:Data.A2", "Data.C1:Data.C2");
        pivot.AddField("Region", "row");
        document.AddNamedRange("Source", "Data.A1:Data.A2");
        document.AddSheet("Later");
        document.MoveSheet("Data", 1);

        Assert.Equal(new[] { "table", "table", "named-expressions", "data-pilot-tables", "consolidation", "dde-links" },
            spreadsheet.Elements().Select(element => element.Name.LocalName));
    }

    [Fact]
    public void ImportedHiddenPivotFieldIsInspectedAndNotClaimedEditable() {
        OdsDocument document = OdsDocument.Create();
        OdsSheet sheet = document.AddSheet("Data");
        sheet.Cell(0, 0).SetString("Region");
        sheet.Cell(1, 0).SetString("North");
        OdsDataPilotTable pivot = document.AddDataPilotTable("Pivot", "Data.A1:Data.A2", "Data.C1:Data.C2");
        OdsDataPilotField field = pivot.AddField("Region", "row");
        field.Element.SetAttributeValue(OdfNamespaces.Table + "orientation", "hidden");
        document.Package.MarkXmlDirty("content.xml");

        Assert.Contains(document.InspectFeatures().Findings, finding =>
            finding.Name == "spreadsheet-data-pilot-tables" && finding.Count == 1 &&
            finding.Support == OdfFeatureSupport.Inspected);
    }

    [Fact]
    public void SourceHeadersAreCheckedBeforeMutationAndDistinctDataAggregationsAreAllowed() {
        OdsDocument document = OdsDocument.Create();
        OdsSheet sheet = document.AddSheet("Data");
        sheet.Cell(0, 0).SetString("Region");
        sheet.Cell(0, 1).SetString("Sales");
        sheet.Cell(1, 0).SetString("North");
        sheet.Cell(1, 1).SetNumber(10);
        OdsDataPilotTable pivot = document.AddDataPilotTable("Pivot", "Data.A1:Data.B2", "Data.D1:Data.E3");

        Assert.Throws<ArgumentException>(() => pivot.AddField("Missing", "row"));
        Assert.Empty(pivot.Fields);
        pivot.AddField("Region", "row");
        pivot.AddField("Sales", "data", "sum");
        pivot.AddField("Sales", "data", "average");
        Assert.Throws<InvalidOperationException>(() => pivot.AddField("Sales", "data", "sum"));
        Assert.Equal(new[] { "sum", "average" }, pivot.Fields.Skip(1).Select(field => field.Function));
    }
}

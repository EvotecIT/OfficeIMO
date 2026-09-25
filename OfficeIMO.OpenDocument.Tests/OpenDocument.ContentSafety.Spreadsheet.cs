using System;
using System.Linq;
using System.Xml.Linq;
using OfficeIMO.ContentSafety;
using Xunit;

namespace OfficeIMO.OpenDocument.Tests;

public sealed class OdsContentSafetyInheritanceTests {
    [Fact]
    public void NestedHiddenColumnGroupConcealsCellText() {
        OdsDocument document = OdsDocument.Create();
        OdsSheet sheet = document.AddSheet("Data");
        sheet.Cell(0, 0).SetString("group hidden payload");
        XElement table = document.Package.GetXml("content.xml")
            .Descendants(OdfNamespaces.Table + "table").Single();
        table.Elements(OdfNamespaces.Table + "table-column").Remove();
        table.AddFirst(new XElement(OdfNamespaces.Table + "table-column-group",
            new XAttribute(OdfNamespaces.Table + "display", "false"),
            new XElement(OdfNamespaces.Table + "table-column")));
        document.MarkPartDirty("content.xml");

        OfficeContentSafetyReport report = OdfDocument.InspectContentSafety(document.ToBytes());
        Assert.Contains(report.Findings, finding =>
            finding.Kind == OfficeContentConcealmentKind.HiddenContainer &&
            finding.TextPreview.IndexOf("group hidden payload", StringComparison.Ordinal) >= 0);
    }

    [Fact]
    public void NestedCollapsedColumnConcealsCellText() {
        OdsDocument document = OdsDocument.Create();
        OdsSheet sheet = document.AddSheet("Data");
        sheet.Cell(0, 0).SetString("collapsed column payload");
        XElement table = document.Package.GetXml("content.xml")
            .Descendants(OdfNamespaces.Table + "table").Single();
        table.Elements(OdfNamespaces.Table + "table-column").Remove();
        table.AddFirst(new XElement(OdfNamespaces.Table + "table-header-columns",
            new XElement(OdfNamespaces.Table + "table-column",
                new XAttribute(OdfNamespaces.Table + "visibility", "collapse"))));
        document.MarkPartDirty("content.xml");

        OfficeContentSafetyReport report = OdfDocument.InspectContentSafety(document.ToBytes());
        Assert.Contains(report.Findings, finding =>
            finding.Kind == OfficeContentConcealmentKind.HiddenContainer &&
            finding.TextPreview.IndexOf("collapsed column payload", StringComparison.Ordinal) >= 0);
    }

    [Fact]
    public void RowDefaultCellStyleConcealsCellText() {
        OdsDocument document = OdsDocument.Create();
        OdfStyle style = document.Styles.CreateNamed("HiddenCell", OdfStyleFamily.TableCell);
        style.Element.Add(new XElement(OdfNamespaces.Style + "text-properties",
            new XAttribute(OdfNamespaces.Text + "display", "none")));
        document.Package.MarkXmlDirty("styles.xml");
        OdsSheet sheet = document.AddSheet("Data");
        sheet.Row(0).DefaultCellStyleName = style.Name;
        sheet.Cell(0, 0).SetString("row style hidden payload");

        OfficeContentSafetyReport report = OdfDocument.InspectContentSafety(document.ToBytes());
        Assert.Contains(report.Findings, finding =>
            finding.Kind == OfficeContentConcealmentKind.HiddenByProperty &&
            finding.TextPreview.IndexOf("row style hidden payload", StringComparison.Ordinal) >= 0);
    }

    [Fact]
    public void NestedColumnDefaultCellStyleConcealsCellText() {
        OdsDocument document = OdsDocument.Create();
        OdfStyle style = document.Styles.CreateNamed("HiddenColumnCell", OdfStyleFamily.TableCell);
        style.Element.Add(new XElement(OdfNamespaces.Style + "text-properties",
            new XAttribute(OdfNamespaces.Text + "display", "none")));
        document.Package.MarkXmlDirty("styles.xml");
        OdsSheet sheet = document.AddSheet("Data");
        sheet.Cell(0, 0).SetString("column style hidden payload");
        XElement table = document.Package.GetXml("content.xml")
            .Descendants(OdfNamespaces.Table + "table").Single();
        table.Elements(OdfNamespaces.Table + "table-column").Remove();
        table.AddFirst(new XElement(OdfNamespaces.Table + "table-header-columns",
            new XElement(OdfNamespaces.Table + "table-column",
                new XAttribute(OdfNamespaces.Table + "default-cell-style-name", style.Name))));
        document.MarkPartDirty("content.xml");

        OfficeContentSafetyReport report = OdfDocument.InspectContentSafety(document.ToBytes());
        Assert.Contains(report.Findings, finding =>
            finding.Kind == OfficeContentConcealmentKind.HiddenByProperty &&
            finding.TextPreview.IndexOf("column style hidden payload", StringComparison.Ordinal) >= 0);
    }

    [Fact]
    public void FamilyDefaultCellStyleConcealsCellText() {
        OdsDocument document = OdsDocument.Create();
        XElement styles = document.Package.GetXml("styles.xml").Root!
            .Element(OdfNamespaces.Office + "styles")!;
        styles.Add(new XElement(OdfNamespaces.Style + "default-style",
            new XAttribute(OdfNamespaces.Style + "family", "table-cell"),
            new XElement(OdfNamespaces.Style + "text-properties",
                new XAttribute(OdfNamespaces.Text + "display", "none"))));
        document.Package.MarkXmlDirty("styles.xml");
        document.AddSheet("Data").Cell(0, 0).SetString("family style hidden payload");

        OfficeContentSafetyReport report = OdfDocument.InspectContentSafety(document.ToBytes());
        Assert.Contains(report.Findings, finding =>
            finding.Kind == OfficeContentConcealmentKind.HiddenByProperty &&
            finding.TextPreview.IndexOf("family style hidden payload", StringComparison.Ordinal) >= 0);
    }

    [Fact]
    public void TransparentCellBackgroundOverridesDarkFamilyDefaultForContrast() {
        OdsDocument document = OdsDocument.Create();
        XElement styles = document.Package.GetXml("styles.xml").Root!
            .Element(OdfNamespaces.Office + "styles")!;
        styles.Add(new XElement(OdfNamespaces.Style + "default-style",
            new XAttribute(OdfNamespaces.Style + "family", "table-cell"),
            new XElement(OdfNamespaces.Style + "table-cell-properties",
                new XAttribute(OdfNamespaces.Fo + "background-color", "#000000"))));
        document.Package.MarkXmlDirty("styles.xml");
        OdfStyle overrideStyle = document.Styles.CreateNamed("TransparentCell", OdfStyleFamily.TableCell);
        overrideStyle.Element.Add(new XElement(OdfNamespaces.Style + "table-cell-properties",
            new XAttribute(OdfNamespaces.Fo + "background-color", "transparent")));
        document.Package.MarkXmlDirty("styles.xml");
        OdsCell cell = document.AddSheet("Data").Cell(0, 0);
        cell.SetString("Visible on white");
        cell.StyleName = overrideStyle.Name;

        OfficeContentSafetyReport report = OdfDocument.InspectContentSafety(document.ToBytes());
        Assert.DoesNotContain(report.Findings, finding =>
            finding.Kind == OfficeContentConcealmentKind.LowContrastText &&
            finding.TextPreview.Contains("Visible on white"));
    }
}

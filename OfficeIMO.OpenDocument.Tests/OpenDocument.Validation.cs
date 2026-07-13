using System;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using Xunit;

namespace OfficeIMO.OpenDocument.Tests;

public sealed class OpenDocumentValidationContractTests {
    [Theory]
    [InlineData("libreoffice-writer-basic.odt")]
    [InlineData("microsoft-word-basic.odt")]
    [InlineData("libreoffice-calc-basic.ods")]
    [InlineData("microsoft-excel-basic.ods")]
    [InlineData("libreoffice-impress-basic.odp")]
    [InlineData("microsoft-powerpoint-basic.odp")]
    public void AuthoredCompatibilityFixturesPassProductionValidation(string fixtureName) {
        string path = Path.Combine(AppContext.BaseDirectory, "Fixtures", fixtureName);
        OdfDocument document = OdfDocument.Load(path);

        OdfValidationResult result = document.Validate();

        Assert.True(result.IsValid, string.Join(Environment.NewLine, result.Diagnostics.Select(item => item.Id + ": " + item.Message)));
    }

    [Fact]
    public void ReportsMissingManifestTargetsAndIncorrectMediaTypes() {
        OdtDocument document = OdtDocument.Create();
        XDocument manifest = document.Package.GetXml("META-INF/manifest.xml");
        XNamespace ns = "urn:oasis:names:tc:opendocument:xmlns:manifest:1.0";
        manifest.Root!.Add(new XElement(ns + "file-entry",
            new XAttribute(ns + "full-path", "Pictures/missing.png"),
            new XAttribute(ns + "media-type", "text/plain")));
        document.Package.MarkXmlDirty("META-INF/manifest.xml");

        OdfValidationResult result = document.Validate();

        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Id == "ODF107");
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Id == "ODF108");
    }

    [Fact]
    public void ReportsMissingStyleParentsAndPackageReferencesAcrossDocumentKinds() {
        OdtDocument document = OdtDocument.Create();
        document.Styles.CreateNamed("Child", OdfStyleFamily.Paragraph, "MissingParent");
        XDocument content = document.Package.GetXml("content.xml");
        XNamespace draw = "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0";
        XNamespace xlink = "http://www.w3.org/1999/xlink";
        XElement body = content.Descendants().First(element => element.Name.LocalName == "text");
        body.Add(new XElement(XName.Get("p", "urn:oasis:names:tc:opendocument:xmlns:text:1.0"),
            new XElement(draw + "frame", new XElement(draw + "image", new XAttribute(xlink + "href", "Pictures/missing.png")))));
        document.Package.MarkXmlDirty("content.xml");

        OdfValidationResult result = document.Validate();

        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Id == "ODF202");
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Id == "ODF300");
    }

    [Fact]
    public void ReportsInvalidSpreadsheetValuesFormulasAndMergeCoverage() {
        OdsDocument document = OdsDocument.Create();
        OdsSheet sheet = document.AddSheet("Data");
        sheet.Merge(0, 0, 1, 2);
        XDocument content = document.Package.GetXml("content.xml");
        XNamespace table = "urn:oasis:names:tc:opendocument:xmlns:table:1.0";
        XNamespace office = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
        XElement anchor = content.Descendants(table + "table-cell").First(element => element.Attribute(table + "number-columns-spanned") != null);
        anchor.SetAttributeValue(office + "value-type", "float");
        anchor.SetAttributeValue(office + "value", "not-a-number");
        anchor.SetAttributeValue(table + "formula", "invalid-formula");
        content.Descendants(table + "covered-table-cell").First().ReplaceWith(new XElement(table + "table-cell"));
        document.Package.MarkXmlDirty("content.xml");

        OdfValidationResult result = document.Validate();

        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Id == "ODS102");
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Id == "ODS103");
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Id == "ODS104");
    }

    [Fact]
    public void FeatureInspectionReportsExternalLinksAndEditableTrackedChanges() {
        OdtDocument document = OdtDocument.Create();
        document.AddTrackedParagraphInsertion("Inserted", "Author");
        document.AddParagraph().AddHyperlink("External", "https://example.com");

        OdfFeatureReport report = document.InspectFeatures();

        Assert.Contains(report.Findings, finding => finding.Name == "tracked-changes" && finding.Support == OdfFeatureSupport.Editable);
        Assert.Contains(report.Findings, finding => finding.Name == "external-links" && finding.Support == OdfFeatureSupport.Preserved);
    }
}

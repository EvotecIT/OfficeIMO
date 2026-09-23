using System;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Xunit;

namespace OfficeIMO.OpenDocument.Tests;

public sealed class OpenDocumentConditionalStyleTests {
    [Fact]
    public void ConditionalCellStyleMapSurvivesReopenAndEdit() {
        OdsDocument document = OdsDocument.Create();
        OdsSheet sheet = document.AddSheet("Data");
        OdfStyle highlight = document.Styles.CreateNamed("Highlight", OdfStyleFamily.TableCell);
        highlight.BackgroundColor = OdfColor.Parse("#FFE699");
        OdfStyle ordinary = document.Styles.CreateAutomatic(OdfStyleFamily.TableCell);
        ordinary.AddConditionalMap("cell-content()>0", highlight.Name, "$'Data'.$A$1");
        ordinary.Bold = true;
        ordinary.BackgroundColor = OdfColor.Parse("#FFFFFF");
        ordinary.TextAlign = "center";
        sheet.Cell(0, 0).StyleName = ordinary.Name;

        OdsDocument reopened = OdsDocument.Load(new MemoryStream(document.ToBytes()));
        Assert.True(reopened.Validate().IsValid);
        OdfStyle mappedStyle = reopened.Styles.Find(OdfStyleFamily.TableCell, ordinary.Name)!;
        OdfStyleMap map = Assert.Single(mappedStyle.ConditionalMaps);
        Assert.Equal("cell-content()>0", map.Condition);
        Assert.Equal("Highlight", map.ApplyStyleName);
        Assert.Equal("$'Data'.$A$1", map.BaseCellAddress);
        XNamespace styleNamespace = "urn:oasis:names:tc:opendocument:xmlns:style:1.0";
        XElement xmlStyle = reopened.Package.GetXml("content.xml")
            .Descendants(styleNamespace + "style").Single(element =>
                (string?)element.Attribute(styleNamespace + "name") == ordinary.Name);
        Assert.Equal(new[] { "table-cell-properties", "paragraph-properties", "text-properties", "map" },
            xmlStyle.Elements().Select(element => element.Name.LocalName));
        Assert.Equal(ordinary.Name, reopened.Sheets.Single().Cell(0, 0).StyleName);
        Assert.Contains(reopened.InspectFeatures().Findings,
            finding => finding.Name == "conditional-style-maps" && finding.Count == 1);

        reopened.Styles.CreateNamed("Changed", OdfStyleFamily.TableCell).BackgroundColor = OdfColor.Parse("#A9D18E");
        map.Condition = "cell-content()<0";
        map.ApplyStyleName = "Changed";
        map.BaseCellAddress = "$'Data'.$B$2";
        OdsDocument edited = OdsDocument.Load(new MemoryStream(reopened.ToBytes()));
        Assert.True(edited.Validate().IsValid);
        OdfStyleMap editedMap = Assert.Single(edited.Styles.Find(OdfStyleFamily.TableCell, ordinary.Name)!.ConditionalMaps);
        Assert.Equal("cell-content()<0", editedMap.Condition);
        Assert.Equal("Changed", editedMap.ApplyStyleName);
        Assert.Equal("$'Data'.$B$2", editedMap.BaseCellAddress);
    }

    [Fact]
    public void ConditionalStyleMapTargetMustBeCommonStyleOfSameFamily() {
        OdsDocument document = OdsDocument.Create();
        document.AddSheet("Data");
        OdfStyle baseStyle = document.Styles.CreateAutomatic(OdfStyleFamily.TableCell);
        OdfStyleMap map = baseStyle.AddConditionalMap("cell-content()>0", "Later", "$'Data'.$A$1");

        Assert.Contains(document.Validate().Diagnostics, diagnostic => diagnostic.Id == "ODF204");
        document.Styles.CreateNamed("Later", OdfStyleFamily.TableCell);
        Assert.True(document.Validate().IsValid);

        map.ApplyStyleName = document.Styles.CreateAutomatic(OdfStyleFamily.TableCell).Name;
        Assert.Contains(document.Validate().Diagnostics, diagnostic => diagnostic.Id == "ODF204");

        document.Styles.CreateNamed("ParagraphOnly", OdfStyleFamily.Paragraph);
        map.ApplyStyleName = "ParagraphOnly";
        Assert.Contains(document.Validate().Diagnostics, diagnostic => diagnostic.Id == "ODF204");
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData(" ")]
    public void ImportedConditionalStyleMapRequiresCondition(string? condition) {
        OdsDocument document = OdsDocument.Create();
        document.AddSheet("Data");
        OdfStyle highlight = document.Styles.CreateNamed("Highlight", OdfStyleFamily.TableCell);
        document.Styles.CreateAutomatic(OdfStyleFamily.TableCell)
            .AddConditionalMap("cell-content()>0", highlight.Name, "$'Data'.$A$1");
        XNamespace styleNamespace = "urn:oasis:names:tc:opendocument:xmlns:style:1.0";
        document.Package.GetXml("content.xml").Descendants(styleNamespace + "map").Single()
            .SetAttributeValue(styleNamespace + "condition", condition);
        document.Package.MarkXmlDirty("content.xml");

        OdfValidationResult result = document.Validate();

        Assert.False(result.IsValid);
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Id == "ODF205" &&
            diagnostic.PartPath == "content.xml");
    }

    [Fact]
    public void InspectorKeepsUnmodeledDataStyleMapsSeparate() {
        OdsDocument document = OdsDocument.Create();
        XNamespace numberNamespace = "urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0";
        XNamespace styleNamespace = "urn:oasis:names:tc:opendocument:xmlns:style:1.0";
        XElement styles = document.Package.GetXml("styles.xml").Descendants(numberNamespace + "number-style").FirstOrDefault()
            ?? new XElement(numberNamespace + "number-style", new XAttribute(styleNamespace + "name", "ConditionalNumber"));
        if (styles.Parent == null) document.Package.GetXml("styles.xml").Root!.Element(
            XName.Get("styles", "urn:oasis:names:tc:opendocument:xmlns:office:1.0"))!.Add(styles);
        styles.Add(new XElement(numberNamespace + "number", new XAttribute(numberNamespace + "min-integer-digits", "1")),
            new XElement(styleNamespace + "map", new XAttribute(styleNamespace + "condition", "value()>0"),
                new XAttribute(styleNamespace + "apply-style-name", "ConditionalNumber")));
        document.Package.MarkXmlDirty("styles.xml");

        OdfFeatureReport report = document.InspectFeatures();

        Assert.Contains(report.Findings, finding => finding.Name == "unmodeled-style-maps" &&
            finding.Support == OdfFeatureSupport.Preserved && finding.Count == 1);
        Assert.DoesNotContain(report.Findings, finding => finding.Name == "conditional-style-maps");
    }

    [Fact]
    public void InspectorKeepsEmbeddedObjectStyleMapsPreservationOnly() {
        OdsDocument document = OdsDocument.Create();
        XNamespace officeNamespace = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
        XNamespace styleNamespace = "urn:oasis:names:tc:opendocument:xmlns:style:1.0";
        var embedded = new XElement(officeNamespace + "document-content",
            new XElement(officeNamespace + "automatic-styles",
                new XElement(styleNamespace + "style",
                    new XAttribute(styleNamespace + "name", "EmbeddedCell"),
                    new XAttribute(styleNamespace + "family", "table-cell"),
                    new XElement(styleNamespace + "map",
                        new XAttribute(styleNamespace + "condition", "cell-content()>0"),
                        new XAttribute(styleNamespace + "apply-style-name", "Highlight")))));
        document.Package.AddOrReplaceEntry("Object 1/content.xml", Encoding.UTF8.GetBytes(embedded.ToString()), "text/xml");

        OdfFeatureReport report = document.InspectFeatures();

        Assert.Contains(report.Findings, finding => finding.Name == "unmodeled-style-maps" &&
            finding.PartPath == "Object 1/content.xml" && finding.Support == OdfFeatureSupport.Preserved);
        Assert.DoesNotContain(report.Findings, finding => finding.Name == "conditional-style-maps" &&
            finding.PartPath == "Object 1/content.xml");
    }
}

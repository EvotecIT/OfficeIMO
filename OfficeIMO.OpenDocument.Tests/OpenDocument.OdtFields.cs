using System;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using Xunit;

namespace OfficeIMO.OpenDocument.Tests;

public sealed class OpenDocumentOdtFieldTests {
    [Fact]
    public void NativeFieldsKeepOrderAndCachedTextAfterReopen() {
        OdtDocument document = OdtDocument.Create();
        OdtParagraph paragraph = document.AddParagraph("Page ");
        paragraph.AddField(OdtFieldKind.PageNumber, "3");
        paragraph.AddText(" of ");
        paragraph.AddField(OdtFieldKind.PageCount, "12");
        paragraph.AddText(" on ");
        paragraph.AddField(OdtFieldKind.Date, "2026-09-25");
        paragraph.AddText(" at ");
        paragraph.AddField(OdtFieldKind.Time, "09:30");

        OdtDocument reopened = OdtDocument.Load(new MemoryStream(document.ToBytes()));
        OdtParagraph result = Assert.Single(reopened.Paragraphs);
        Assert.Equal("Page 3 of 12 on 2026-09-25 at 09:30", result.Text);
        Assert.Equal(new[] { OdtFieldKind.PageNumber, OdtFieldKind.PageCount,
            OdtFieldKind.Date, OdtFieldKind.Time }, result.Fields.Select(field => field.Kind));
        Assert.Equal(4, result.InlineNodes.Count(node => node.Kind == OdtInlineNodeKind.Field));
        Assert.Equal("12", result.Fields[1].DisplayText);
        Assert.Contains(reopened.InspectFeatures().Findings, finding => finding.Name == "text-fields" &&
            finding.Support == OdfFeatureSupport.Editable && finding.Count == 4);
    }

    [Fact]
    public void FieldDisplayAndFixedStateRemainEditable() {
        OdtDocument document = OdtDocument.Create();
        OdtField field = document.AddParagraph().AddField(OdtFieldKind.Date, "Old");
        field.DisplayText = "New";
        field.IsFixed = true;

        OdtDocument reopened = OdtDocument.Load(new MemoryStream(document.ToBytes()));
        OdtField result = Assert.Single(Assert.Single(reopened.Paragraphs).Fields);
        Assert.Equal("New", result.DisplayText);
        Assert.True(result.IsFixed);
        var text = (System.Xml.Linq.XNamespace)"urn:oasis:names:tc:opendocument:xmlns:text:1.0";
        reopened.Package.GetXml("content.xml").Descendants(text + "date").Single()
            .SetAttributeValue(text + "date-value", "2026-09-25");
        Assert.Contains(reopened.InspectFeatures().Findings, finding => finding.Name == "text-fields" &&
            finding.Support == OdfFeatureSupport.Inspected && finding.Count == 1);
        Assert.Throws<NotSupportedException>(() =>
            reopened.Paragraphs.Single().AddField(OdtFieldKind.PageCount, "2").IsFixed = true);
    }

    [Fact]
    public void FieldInspectorLimitsEditableClaimsToOdtStories() {
        OdtDocument textDocument = OdtDocument.Create();
        textDocument.AddParagraph().AddField(OdtFieldKind.Date, "Today");
        XDocument content = textDocument.Package.GetXml("content.xml");
        XElement body = content.Descendants(OdfNamespaces.Office + "text").Single();
        body.Add(new XElement(OdfNamespaces.Text + "tracked-changes",
            new XElement(OdfNamespaces.Text + "changed-region",
                new XElement(OdfNamespaces.Text + "deletion",
                    new XElement(OdfNamespaces.Text + "p",
                        new XElement(OdfNamespaces.Text + "date", "Old"))))));
        body.Add(new XElement(OdfNamespaces.Text + "table-of-content",
            new XElement(OdfNamespaces.Text + "index-body",
                new XElement(OdfNamespaces.Text + "p",
                    new XElement(OdfNamespaces.Text + "page-number", "4")))));
        body.Add(new XElement(OdfNamespaces.Draw + "frame",
            new XElement(OdfNamespaces.Draw + "text-box",
                new XElement(OdfNamespaces.Text + "p",
                    new XElement(OdfNamespaces.Text + "date", "Inside drawing")))));
        body.Add(new XElement(OdfNamespaces.Table + "table",
            new XElement(OdfNamespaces.Table + "table-row",
                new XElement(OdfNamespaces.Table + "table-cell",
                    new XElement(OdfNamespaces.Table + "table",
                        new XElement(OdfNamespaces.Table + "table-row",
                            new XElement(OdfNamespaces.Table + "table-cell",
                                new XElement(OdfNamespaces.Text + "p",
                                    new XElement(OdfNamespaces.Text + "page-number", "5")))))))));
        textDocument.Package.MarkXmlDirty("content.xml");
        OdfFeatureFinding[] odtFindings = textDocument.InspectFeatures().Findings
            .Where(finding => finding.Name == "text-fields").ToArray();
        Assert.Contains(odtFindings, finding => finding.Support == OdfFeatureSupport.Editable && finding.Count == 1);
        Assert.Contains(odtFindings, finding => finding.Support == OdfFeatureSupport.Inspected && finding.Count == 4);

        OdpPresentation presentation = OdpPresentation.Create();
        presentation.AddSlide().AddTextBox(OdfRect.FromCentimeters(1, 1, 4, 2), "Today");
        XElement paragraph = presentation.Package.GetXml("content.xml")
            .Descendants(OdfNamespaces.Text + "p").Single();
        paragraph.Add(new XElement(OdfNamespaces.Text + "date", "Today"));
        presentation.Package.MarkXmlDirty("content.xml");
        OdfFeatureFinding finding = Assert.Single(presentation.InspectFeatures().Findings,
            item => item.Name == "text-fields");
        Assert.Equal(OdfFeatureSupport.Inspected, finding.Support);
    }

    [Fact]
    public void FirstMasterHeaderAndFooterVariantsAreEditable() {
        OdtDocument document = OdtDocument.Create();
        document.PageLayout.Header.AddParagraph().AddField(OdtFieldKind.Date, "First");
        XDocument styles = document.Package.GetXml("styles.xml");
        XElement first = styles.Descendants(OdfNamespaces.Style + "master-page").First();
        first.Add(new XElement(OdfNamespaces.Style + "header-left",
            new XElement(OdfNamespaces.Text + "p",
                new XElement(OdfNamespaces.Text + "date", "Left"))));
        first.Parent!.Add(new XElement(OdfNamespaces.Style + "master-page",
            new XAttribute(OdfNamespaces.Style + "name", "Second"),
            new XElement(OdfNamespaces.Style + "header",
                new XElement(OdfNamespaces.Text + "p",
                    new XElement(OdfNamespaces.Text + "date", "Second")))));
        document.Package.MarkXmlDirty("styles.xml");

        OdfFeatureFinding[] fields = document.InspectFeatures().Findings
            .Where(finding => finding.Name == "text-fields" && finding.PartPath == "styles.xml")
            .ToArray();
        Assert.Contains(fields, finding => finding.Support == OdfFeatureSupport.Editable && finding.Count == 2);
        Assert.Contains(fields, finding => finding.Support == OdfFeatureSupport.Inspected && finding.Count == 1);
    }
}

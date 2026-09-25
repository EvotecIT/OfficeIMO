using System;
using System.IO;
using System.Linq;
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
}

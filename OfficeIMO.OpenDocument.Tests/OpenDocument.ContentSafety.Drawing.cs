using System;
using System.Linq;
using System.Xml.Linq;
using OfficeIMO.ContentSafety;
using Xunit;

namespace OfficeIMO.OpenDocument.Tests;

public sealed class OdfDrawingContentSafetyTests {
    [Fact]
    public void NestedFrameFillOverridesAncestorParagraphBackgroundForContrast() {
        OdtDocument document = OdtDocument.Create();
        OdfStyle outerStyle = document.Styles.CreateNamed("OuterWhite", OdfStyleFamily.Paragraph);
        outerStyle.Element.Add(new XElement(OdfNamespaces.Style + "paragraph-properties",
            new XAttribute(OdfNamespaces.Fo + "background-color", "#ffffff")),
            new XElement(OdfNamespaces.Style + "text-properties",
                new XAttribute(OdfNamespaces.Fo + "color", "#000000")));
        OdfStyle frameStyle = document.Styles.CreateNamed("InnerBlack", OdfStyleFamily.Graphic);
        frameStyle.Element.Add(new XElement(OdfNamespaces.Style + "graphic-properties",
            new XAttribute(OdfNamespaces.Draw + "fill", "solid"),
            new XAttribute(OdfNamespaces.Draw + "fill-color", "#000000")));
        OdfStyle innerWhiteStyle = document.Styles.CreateNamed("InnerWhite", OdfStyleFamily.Paragraph);
        innerWhiteStyle.Element.Add(new XElement(OdfNamespaces.Style + "paragraph-properties",
            new XAttribute(OdfNamespaces.Fo + "background-color", "#ffffff")));
        document.Package.MarkXmlDirty("styles.xml");

        OdtParagraph paragraph = document.AddParagraph("Outer visible text");
        paragraph.StyleName = outerStyle.Name;
        paragraph.Element.Add(new XElement(OdfNamespaces.Draw + "frame",
            new XAttribute(OdfNamespaces.Draw + "style-name", frameStyle.Name),
            new XElement(OdfNamespaces.Draw + "text-box",
                new XElement(OdfNamespaces.Text + "p", "Black on black frame"),
                new XElement(OdfNamespaces.Text + "p",
                    new XAttribute(OdfNamespaces.Text + "style-name", innerWhiteStyle.Name),
                    "Black on white frame"))));
        document.MarkPartDirty("content.xml");

        OfficeContentSafetyReport report = OdfDocument.InspectContentSafety(document.ToBytes());
        Assert.Contains(report.Findings, finding =>
            finding.Kind == OfficeContentConcealmentKind.LowContrastText &&
            finding.TextPreview.IndexOf("Black on black frame", StringComparison.Ordinal) >= 0);
        Assert.DoesNotContain(report.Findings, finding =>
            finding.Kind == OfficeContentConcealmentKind.LowContrastText &&
            finding.TextPreview.IndexOf("Black on white frame", StringComparison.Ordinal) >= 0);
    }
}

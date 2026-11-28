using System.Text;
using System.Xml.Linq;
using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.PowerPoint.OpenXml;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;
using OfficeIMO.OpenXml.Internal;

namespace OfficeIMO.Tests;

/// <summary>
/// Element edits through <see cref="OfficeXmlNode"/> give the same markup on OfficeIMO's tree and on the SDK's
/// elements, and elements added on the SDK side are the SDK's typed classes, so code holding the SDK package keeps
/// reading them through its typed members.
/// </summary>
public class OfficeXmlNodeTests {
    private const string Slide =
        "<p:sld xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\">"
        + "<p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id=\"1\" name=\"\"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>"
        + "<p:sp><p:nvSpPr><p:cNvPr id=\"2\" name=\"Title\"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr/>"
        + "<p:txBody><a:bodyPr/><a:p><!-- note --><a:r><a:t>Hello</a:t></a:r><a:endParaRPr lang=\"en-US\"/></a:p></p:txBody></p:sp>"
        + "</p:spTree></p:cSld></p:sld>";

    [Fact]
    public void TypedSdkElementsKeepTheirClassWhereTheSdkCannotTypeThemOutOfContext() {
        var choice = new DocumentFormat.OpenXml.AlternateContentChoice { Requires = "wps" };
        var alternate = new DocumentFormat.OpenXml.AlternateContent(choice);
        var paragraph = new DocumentFormat.OpenXml.Wordprocessing.Paragraph(new DocumentFormat.OpenXml.Wordprocessing.Run(alternate));

        OfficeXmlNode.FromSdk(choice).Append(OfficeXmlNode.FromSdk(new DocumentFormat.OpenXml.Wordprocessing.Run(
            new DocumentFormat.OpenXml.Wordprocessing.Text("inside"))));

        DocumentFormat.OpenXml.Wordprocessing.Run? run = choice.GetFirstChild<DocumentFormat.OpenXml.Wordprocessing.Run>();
        Assert.NotNull(run);
        Assert.Equal("inside", run!.InnerText);
        Assert.Same(choice, paragraph.Descendants<DocumentFormat.OpenXml.AlternateContentChoice>().Single());
    }

    [Fact]
    public void EditsGiveTheSameMarkupOnTheTreeAndOnTheSdkElements() {
        (OfficeXmlNode tree, OfficeXmlNode sdk, P.Slide slide) = Both();

        foreach (OfficeXmlNode root in new[] { tree, sdk }) {
            OfficeXmlNode paragraph = Paragraph(root);
            OfficeXmlNode run = paragraph.Child(OfficeXmlNamespaces.DrawingML, "r");
            Assert.Equal("Hello", run.InnerText);

            // Properties go before the text, in schema order.
            OfficeXmlNode properties = run.GetOrAddChild(OfficeXmlNamespaces.DrawingML, "rPr", ARegularTextRunView.Order);
            properties.SetBooleanAttribute("b", true);
            properties.SetInt32Attribute("sz", 2400);
            properties.AddOrdered(OfficeXmlNamespaces.DrawingML, "latin", ATextCharacterPropertiesOrder).SetAttribute("typeface", "Aptos");
            run.Child(OfficeXmlNamespaces.DrawingML, "t").SetInnerText("Hi & bye");

            // A run built detached, then inserted before the end paragraph properties.
            OfficeXmlNode added = root.CreateElement(OfficeXmlNamespaces.DrawingML, "r");
            added = paragraph.InsertBefore(BuildRun(added, "second"), paragraph.Child(OfficeXmlNamespaces.DrawingML, "endParaRPr"));
            Assert.Equal("second", added.InnerText);

            OfficeXmlNode drawing = root.Child(OfficeXmlNamespaces.PresentationML, "cSld").Child(OfficeXmlNamespaces.PresentationML, "spTree")
                .Child(OfficeXmlNamespaces.PresentationML, "sp").Child(OfficeXmlNamespaces.PresentationML, "nvSpPr").Child(OfficeXmlNamespaces.PresentationML, "cNvPr");
            drawing.SetAttribute("descr", "A title");
            drawing.SetAttribute("name", null);
            Assert.Null(drawing.Attribute("name"));
            Assert.Equal(true, properties.BooleanAttribute("b"));
            Assert.Equal(2400, properties.Int32Attribute("sz"));
        }

        Assert.Equal(Normalize(tree.OuterXml()), Normalize(sdk.OuterXml()));

        // The SDK side reads everything through its typed members.
        A.Paragraph sdkParagraph = slide.Descendants<A.Paragraph>().Single();
        A.Run[] runs = sdkParagraph.Elements<A.Run>().ToArray();
        Assert.Equal(2, runs.Length);
        Assert.True(runs[0].RunProperties!.Bold!.Value);
        Assert.Equal(2400, runs[0].RunProperties!.FontSize!.Value);
        Assert.Equal("Aptos", runs[0].RunProperties!.GetFirstChild<A.LatinFont>()!.Typeface!.Value);
        Assert.Equal("Hi & bye", runs[0].Text!.Text);
        Assert.Equal("second", runs[1].Text!.Text);
        Assert.Equal("A title", slide.Descendants<P.NonVisualDrawingProperties>().First(item => item.Id!.Value == 2).Description!.Value);
        // The inserted run declares nothing its parent already declares.
        Assert.Empty(runs[1].NamespaceDeclarations);
    }

    [Fact]
    public void ChildrenAreElementsInOrderOnBothSides() {
        (OfficeXmlNode tree, OfficeXmlNode sdk, _) = Both();
        foreach (OfficeXmlNode root in new[] { tree, sdk }) {
            OfficeXmlNode paragraph = Paragraph(root);
            // The comment before the run is not a child element.
            Assert.Equal(new[] { "r", "endParaRPr" }, paragraph.Children().ToList().Select(child => child.LocalName));
            Assert.Equal(paragraph.Child(OfficeXmlNamespaces.DrawingML, "endParaRPr"), paragraph.LastChild);
            Assert.Equal(paragraph, paragraph.FirstChild.Parent);
            Assert.True(paragraph.FirstChild.IsAttached);
            Assert.True(paragraph.FirstChild.IsInside(root));
            OfficeXmlNode copy = paragraph.Clone();
            Assert.False(copy.IsAttached);
            Assert.Equal("Hello", copy.InnerText);
        }
    }

    [Fact]
    public void RemovesAndMovesElementsOnBothSides() {
        (OfficeXmlNode tree, OfficeXmlNode sdk, P.Slide slide) = Both();
        foreach (OfficeXmlNode root in new[] { tree, sdk }) {
            OfficeXmlNode paragraph = Paragraph(root);
            OfficeXmlNode run = paragraph.Child(OfficeXmlNamespaces.DrawingML, "r");
            OfficeXmlNode copy = run.Clone();
            run.InsertAfterSelf(copy);
            Assert.Equal(2, paragraph.Children(OfficeXmlNamespaces.DrawingML, "r").Count);
            paragraph.RemoveChildren(OfficeXmlNamespaces.DrawingML, "r");
            Assert.True(paragraph.Child(OfficeXmlNamespaces.DrawingML, "r").IsNone);
        }
        Assert.Equal(Normalize(tree.OuterXml()), Normalize(sdk.OuterXml()));
        Assert.Empty(slide.Descendants<A.Run>());
    }

    [Fact]
    public void SdkElementsMoveIntoTheTreeAsMarkup() {
        (OfficeXmlNode tree, OfficeXmlNode sdk, _) = Both();
        OfficeXmlNode sdkRun = Paragraph(sdk).Child(OfficeXmlNamespaces.DrawingML, "r");
        OfficeXmlNode treeParagraph = Paragraph(tree);
        OfficeXmlNode inserted = treeParagraph.Append(sdkRun.Clone());
        Assert.Same(tree.Tree, inserted.Tree);
        Assert.Equal("Hello", inserted.InnerText);
    }

    private static readonly OfficeXmlChildOrder ATextCharacterPropertiesOrder = ATextCharacterPropertiesView.Order;

    private static OfficeXmlNode BuildRun(OfficeXmlNode run, string text) {
        OfficeXmlNode properties = run.Append(run.CreateElement(OfficeXmlNamespaces.DrawingML, "rPr"));
        properties.SetAttribute("lang", "en-US");
        run.Append(run.CreateElement(OfficeXmlNamespaces.DrawingML, "t")).SetInnerText(text);
        return run;
    }

    private static OfficeXmlNode Paragraph(OfficeXmlNode root) =>
        root.Child(OfficeXmlNamespaces.PresentationML, "cSld").Child(OfficeXmlNamespaces.PresentationML, "spTree")
            .Child(OfficeXmlNamespaces.PresentationML, "sp").Child(OfficeXmlNamespaces.PresentationML, "txBody")
            .Child(OfficeXmlNamespaces.DrawingML, "p");

    private static (OfficeXmlNode Tree, OfficeXmlNode Sdk, P.Slide Slide) Both() {
        byte[] bytes = Encoding.UTF8.GetBytes(Slide);
        OfficeXmlTree tree = OfficeXmlTree.Parse(bytes, bytes.Length, new OfficeXmlNamespaces.Table());
        var slide = new P.Slide(Slide);
        return (new OfficeXmlNode(tree, tree.Root), OfficeXmlNode.FromSdk(slide), slide);
    }

    // Comments, namespace declaration placement and attribute order (the SDK writes attributes in schema order) are not
    // part of the content compared here.
    private static string Normalize(byte[] markup) {
        XElement element = XElement.Parse(Encoding.UTF8.GetString(markup));
        foreach (XComment comment in element.DescendantNodes().OfType<XComment>().ToList()) comment.Remove();
        foreach (XElement item in element.DescendantsAndSelf()) {
            XAttribute[] attributes = item.Attributes().Where(attribute => !attribute.IsNamespaceDeclaration)
                .OrderBy(attribute => attribute.Name.ToString(), StringComparer.Ordinal).ToArray();
            item.RemoveAttributes();
            item.Add(attributes);
        }
        return element.ToString(SaveOptions.DisableFormatting);
    }
}

using System.Text;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO;
using OfficeIMO.PowerPoint;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class PowerPointViewLengthsAndCompositionReturns {
    [Fact]
    public void ViewLengthsAndReturnedGuidesRoundTripThroughTheNativePackage() {
        using var document = PowerPointDocument.Create();
        document.AddSlide();
        document.SetGridSpacing(OfficeLength.Emus(12345), OfficeLength.Points(3));
        var guide = document.AddGuide(PowerPointGuideOrientation.Vertical, OfficeLength.Emus(-12345));
        Assert.Equal(OfficeLength.Emus(-12345), guide.Position);
        var columns = document.AddColumnGuides(2, OfficeLength.Points(10), OfficeLength.Points(5));
        Assert.Equal(4, columns.Count);
        Assert.All(columns, item => Assert.Equal(PowerPointGuideOrientation.Vertical, item.Orientation));
        var rows = document.AddRowGuides(2, OfficeLength.Points(10), OfficeLength.Points(5), includeOuterEdges: false);
        Assert.Equal(2, rows.Count);
        Assert.Throws<ArgumentOutOfRangeException>(() => document.SetGridSpacing(OfficeLength.Points(7), OfficeLength.Emus(long.MaxValue)));
        Assert.Equal(OfficeLength.Emus(12345), document.GridSpacingX);
        Assert.Throws<ArgumentOutOfRangeException>(() => document.SetGuides(new[] {
            guide, new PowerPointGuideInfo(PowerPointGuideOrientation.Horizontal, OfficeLength.Emus(long.MaxValue))
        }));
        Assert.Equal(1 + columns.Count + rows.Count, document.GetGuides().Count);
        using var output = new MemoryStream();
        document.Save(output);
        Assert.False(document.IsSdkLoaded);
        output.Position = 0;
        using (var package = PresentationDocument.Open(output, false)) {
            var view = package.PresentationPart!.ViewPropertiesPart!.ViewProperties!;
            Assert.Equal(12345, view.GridSpacing!.Cx!.Value);
            Assert.Equal(38100, view.GridSpacing.Cy!.Value);
        }
        output.Position = 0;
        using var loaded = PowerPointDocument.Load(output);
        Assert.Equal(document.GridSpacingY, loaded.GridSpacingY);
        Assert.Equal(guide.Position, loaded.GetGuides()[0].Position);
        Assert.Equal(1 + columns.Count + rows.Count, loaded.GetGuides().Count);
        Assert.False(loaded.IsSdkLoaded);
    }

    [Fact]
    public void ViewPropertiesIgnoreSettingsInsideExtensions() {
        using var source = PowerPointDocument.Create();
        source.AddSlide();
        source.SetGridSpacing(OfficeLength.Emus(321), OfficeLength.Emus(654));
        using var input = new MemoryStream(source.ToBytes());
        using OfficeOpcEditablePackage package = OfficeOpcEditablePackage.Open(input, leaveOpen: true);
        const string viewPart = "/ppt/viewProps.xml";
        XDocument view = XDocument.Parse(Encoding.UTF8.GetString(package.ReadPartBytes(viewPart)));
        XNamespace p = "http://schemas.openxmlformats.org/presentationml/2006/main";
        XNamespace unknown = "urn:officeimo:test:unselected-view";
        XElement root = view.Root!;
        root.AddFirst(new XElement(p + "extLst", new XElement(p + "ext", new XAttribute("uri", "{F1A55DE2-61FD-492C-8EAD-BF0717977356}"),
            new XElement(unknown + "settings", new XElement(p + "gridSpacing", new XAttribute("cx", 999)),
                new XElement(p + "slideViewPr", new XElement(p + "cSldViewPr", new XAttribute("snapToGrid", true)))))));
        byte[] xml = Encoding.UTF8.GetBytes(view.ToString(SaveOptions.DisableFormatting));
        package.ReplacePart(viewPart, xml, xml.Length);
        using var deck = new MemoryStream();
        package.Save(deck);
        deck.Position = 0;
        using var loaded = PowerPointDocument.Load(deck);

        Assert.Equal(OfficeLength.Emus(321), loaded.GridSpacingX);
        Assert.False(loaded.SnapToGrid);
        Assert.False(loaded.IsSdkLoaded);
    }

    [Fact]
    public void ViewSettingsCreateTheMissingViewPartWithoutOpeningTheSdk() {
        using var source = PowerPointDocument.Create();
        source.AddSlide();
        byte[] bytes = source.ToBytes();
        using var input = new MemoryStream();
        input.Write(bytes, 0, bytes.Length);
        input.Position = 0;
        using (PresentationDocument original = PresentationDocument.Open(input, true)) {
            var presentation = original.PresentationPart!;
            presentation.DeletePart(presentation.ViewPropertiesPart!);
        }

        input.Position = 0;
        using var document = PowerPointDocument.Load(input);
        document.SnapToGrid = true;
        document.SetGridSpacing(OfficeLength.Emus(444), OfficeLength.Emus(666));
        document.AddGuide(PowerPointGuideOrientation.Vertical, OfficeLength.Emus(-222));
        byte[] saved = document.ToBytes();

        Assert.False(document.IsSdkLoaded);
        using PresentationDocument native = PresentationDocument.Open(new MemoryStream(saved), false);
        var view = native.PresentationPart!.ViewPropertiesPart!.ViewProperties!;
        Assert.Equal(444L, view.GridSpacing!.Cx!.Value);
        using var reopened = PowerPointDocument.Load(new MemoryStream(saved));
        Assert.True(reopened.SnapToGrid);
        Assert.Equal(OfficeLength.Emus(-222), Assert.Single(reopened.GetGuides()).Position);
    }

    [Fact]
    public void InvalidGuidePositionKeepsTheSdkParseOutcome() {
        using var source = PowerPointDocument.Create();
        source.AddSlide();
        source.AddGuide(PowerPointGuideOrientation.Horizontal, OfficeLength.Emus(222));
        byte[] damaged = RewriteViewPackage(source.ToBytes(), package => {
            const int p = OfficeXmlNamespaces.PresentationML;
            OfficeXmlTree view = package.GetXml("/ppt/viewProps.xml");
            int slide = view.FindChild(view.Root, p, "slideViewPr");
            int common = view.FindChild(slide, p, "cSldViewPr");
            int guides = view.FindChild(common, p, "guideLst");
            int guide = view.FindChild(guides, p, "guide");
            view.SetAttribute(guide, OfficeXmlNamespaces.None, "pos", "bogus");
        });

        using var document = PowerPointDocument.Load(new MemoryStream(damaged));
        using PresentationDocument sdk = PresentationDocument.Open(new MemoryStream(damaged), false);
        var view = sdk.PresentationPart!.ViewPropertiesPart!.ViewProperties!;
        var guide = Assert.Single(view.Descendants<DocumentFormat.OpenXml.Presentation.Guide>());

        Assert.Throws<FormatException>(() => _ = guide.Position!.Value);
        Assert.Throws<FormatException>(() => document.GetGuides());
        Assert.False(document.IsSdkLoaded);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void InvalidViewRelationshipKeepsTheSdkOutcome(bool wrongContentType) {
        using var source = PowerPointDocument.Create();
        source.AddSlide();
        const string viewPart = "/ppt/viewProps.xml";
        byte[] damaged = RewriteViewPackage(source.ToBytes(), package => {
            if (!wrongContentType) {
                package.DeletePart(viewPart);
                return;
            }
            OfficeXmlTree types = package.ContentTypes;
            for (int entry = types.FindChild(types.Root, OfficeXmlNamespaces.ContentTypes, "Override"); entry != OfficeXmlTree.NoNode;
                entry = types.FindNextSibling(entry, OfficeXmlNamespaces.ContentTypes, "Override")) {
                if (types.GetAttribute(entry, OfficeXmlNamespaces.None, "PartName") != viewPart) continue;
                types.SetAttribute(entry, OfficeXmlNamespaces.None, "ContentType", "application/xml");
                return;
            }
            throw new InvalidOperationException("The test package has no view content type.");
        });

        Exception? expected = Record.Exception(() => {
            using PresentationDocument sdk = PresentationDocument.Open(new MemoryStream(damaged), false);
            _ = sdk.PresentationPart!.ViewPropertiesPart!.ViewProperties;
        });
        Exception? actual = Record.Exception(() => {
            using var document = PowerPointDocument.Load(new MemoryStream(damaged));
            document.GetGuides();
        });

        Assert.NotNull(expected);
        Assert.Equal(expected.GetType(), actual?.GetType());
    }

    private static byte[] RewriteViewPackage(byte[] input, Action<OfficeOpcEditablePackage> edit) {
        using var source = new MemoryStream(input, writable: false);
        using OfficeOpcEditablePackage package = OfficeOpcEditablePackage.Open(source, leaveOpen: true);
        edit(package);
        using var output = new MemoryStream();
        package.Save(output);
        return output.ToArray();
    }

    [Fact]
    public void CompositeAddsReturnOnlyTheirNewAttachedEditableShapes() {
        using var document = PowerPointDocument.Create();
        var plan = new PowerPointDeckPlan().AddCustom("Editable composition", context => {
            var beforeTitle = context.Slide.Shapes.ToArray();
            var title = context.AddTitle("Original", "Subtitle");
            Assert.Equal(context.Slide.Shapes.Except(beforeTitle), title);
            Assert.All(title, shape => Assert.Contains(shape, context.Slide.Shapes));
            Assert.Single(title.OfType<PowerPointTextBox>(), box => box.Text == "Original").Text = "Edited title";
            var beforeCallout = context.Slide.Shapes.ToArray();
            var callout = context.AddCalloutBand("Original callout", context.ContentArea());
            Assert.Equal(context.Slide.Shapes.Except(beforeCallout), callout);
            Assert.All(callout, shape => Assert.DoesNotContain(shape, title));
            Assert.Single(callout.OfType<PowerPointTextBox>(), box => box.Text == "Original callout").Text = "Edited callout";
            var explicitText = context.AddTextBlock("Physical font", context.ContentArea(), OfficeLength.Inches(0.1875));
            Assert.Equal(OfficeLength.Points(13.5), explicitText.FontSize);
            var defaultText = context.AddTextBlock("Default font", context.ContentArea());
            Assert.Equal(OfficeLength.Points(12), defaultText.FontSize);
            int count = context.Slide.Shapes.Count;
            Assert.Throws<ArgumentOutOfRangeException>(() => context.AddTextBlock("Invalid font", context.ContentArea(), OfficeLength.Zero));
            Assert.Equal(count, context.Slide.Shapes.Count);
            Assert.Empty(context.AddMetricStrip(Array.Empty<PowerPointMetric>(), context.ContentArea()));
        });
        var options = PowerPointCompositionOptions.FromBrief(PowerPointDesignBrief.FromBrand("#008C95", "composition-return-proof"));
        options.SelectBestAlternative = false;
        document.Compose(plan, options);
        using var output = new MemoryStream();
        document.Save(output);
        output.Position = 0;
        using var loaded = PowerPointDocument.Load(output);
        var text = loaded.Slides.Single().TextBoxes.Select(box => box.Text).ToArray();
        Assert.Contains("Edited title", text);
        Assert.Contains("Edited callout", text);
        var textBoxes = loaded.Slides.Single().TextBoxes;
        Assert.Equal(OfficeLength.Points(13.5), textBoxes.Single(box => box.Text == "Physical font").FontSize);
        Assert.Equal(OfficeLength.Points(12), textBoxes.Single(box => box.Text == "Default font").FontSize);
    }
}

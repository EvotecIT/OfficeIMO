using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing.Diagrams;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using OfficeIMO.Word;
using Dsp = DocumentFormat.OpenXml.Office.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class WordSmartArtStoryOwnershipTests {
    public static IEnumerable<object[]> StoryLayouts() {
        foreach (WordSmartArtType type in Enum.GetValues(typeof(WordSmartArtType))) {
            yield return new object[] { false, type };
            yield return new object[] { true, type };
        }
    }

    [Theory]
    [MemberData(nameof(StoryLayouts))]
    public void HeaderAndFooterDiagramsOwnTheirPartsAndRoundTrip(bool footer, WordSmartArtType type) {
        using WordDocument document = WordDocument.Create();
        WordHeaderFooter story = footer ? document.FooterDefaultOrCreate : document.HeaderDefaultOrCreate;
        WordSmartArt created = story.AddParagraph().AddSmartArt(type);
        created.SetNodeText(0, "Original story node");
        AssertStoryParts(document, created, footer, type);
        byte[] bytes = document.ToBytes();

        using WordDocument reloaded = WordDocument.Load(new MemoryStream(bytes));
        WordSmartArt imported = Assert.Single(GetStory(reloaded, footer).SmartArts);
        AssertStoryParts(reloaded, imported, footer, type);
        Assert.Equal("Original story node", imported.GetNodeText(0));
        imported.SetNodeText(0, "Edited story node");
        if (type is WordSmartArtType.BasicProcess or WordSmartArtType.Cycle) {
            Assert.True(imported.CanModifyNodes);
            imported.AddNode("Added after reopening");
        }

        using WordDocument edited = WordDocument.Load(new MemoryStream(reloaded.ToBytes()));
        WordSmartArt editedArt = Assert.Single(GetStory(edited, footer).SmartArts);
        Assert.Equal("Edited story node", editedArt.GetNodeText(0));
        if (type is WordSmartArtType.BasicProcess or WordSmartArtType.Cycle)
            Assert.Equal("Added after reopening", editedArt.GetNodeText(editedArt.NodeCount - 1));
        AssertStoryParts(edited, editedArt, footer, type);
        var errors = new OpenXmlValidator(FileFormatVersions.Microsoft365).Validate(edited._wordprocessingDocument).ToList();
        Assert.True(errors.Count == 0, string.Join(Environment.NewLine,
            errors.Select(error => $"{error.Part?.Uri} {error.Path?.XPath}: {error.Description}")));
    }

    [Fact]
    public void StoryReadersKeepCollidingRelationshipIdsSeparateFromBody() {
        using WordDocument document = WordDocument.Create();
        WordSmartArt body = document.AddSmartArt(WordSmartArtType.CustomSmartArt2);
        WordSmartArt header = document.HeaderDefaultOrCreate.AddParagraph().AddSmartArt(WordSmartArtType.CustomSmartArt1);
        WordSmartArt footer = document.FooterDefaultOrCreate.AddParagraph().AddSmartArt(WordSmartArtType.CustomSmartArt2);
        foreach (var item in new[] { (body, "Body"), (header, "Header"), (footer, "Footer") }) {
            item.Item1.SetNodeText(0, item.Item2);
            OpenXmlPart owner = WordPartOwnership.Resolve(document, item.Item1._drawing);
            RelationshipIds ids = item.Item1._drawing.Descendants<RelationshipIds>().Single();
            StringValue[] values = { ids.DataPart!, ids.LayoutPart!, ids.StylePart!, ids.ColorPart! };
            for (int i = 0; i < values.Length; i++) {
                OpenXmlPart part = owner.GetPartById(values[i].Value!);
                owner.ChangeIdOfPart(part, "rIdDiagram" + i);
                values[i].Value = "rIdDiagram" + i;
            }
            var data = (DiagramDataPart)owner.GetPartById(ids.DataPart!.Value!);
            Dsp.DataModelExtensionBlock extension = data.DataModelRoot!.Descendants<Dsp.DataModelExtensionBlock>().Single();
            OpenXmlPart drawing = owner.GetPartById(extension.RelId!.Value!);
            owner.ChangeIdOfPart(drawing, "rIdCachedDrawing");
            extension.RelId = "rIdCachedDrawing";
        }

        using WordDocument reloaded = WordDocument.Load(new MemoryStream(document.ToBytes()));
        Assert.Equal("Body", Assert.Single(reloaded.SmartArts).GetNodeText(0));
        WordSmartArt headerArt = Assert.Single(GetStory(reloaded, false).SmartArts);
        WordSmartArt footerArt = Assert.Single(GetStory(reloaded, true).SmartArts);
        Assert.Equal("Header", headerArt.GetNodeText(0));
        Assert.Equal("Footer", footerArt.GetNodeText(0));
        Assert.Equal(WordSmartArtType.BasicProcess, headerArt.LayoutType);
        Assert.Equal(WordSmartArtType.Cycle, footerArt.LayoutType);
        Assert.Contains(headerArt.GetPersistedLayoutShapes(), shape => shape.Text == "Header");
        Assert.Contains(footerArt.GetPersistedLayoutShapes(), shape => shape.Text == "Footer");
        headerArt.SetNodeText(0, "Changed header");
        Assert.Equal("Body", Assert.Single(reloaded.SmartArts).GetNodeText(0));
        Assert.Equal("Footer", footerArt.GetNodeText(0));
    }

    private static WordHeaderFooter GetStory(WordDocument document, bool footer) => footer
        ? Assert.IsType<WordFooter>(document.Sections[0].Footer.Default)
        : Assert.IsType<WordHeader>(document.Sections[0].Header.Default);

    private static void AssertStoryParts(WordDocument document, WordSmartArt smartArt, bool footer, WordSmartArtType type) {
        OpenXmlPart owner = WordPartOwnership.Resolve(document, smartArt._drawing);
        if (footer) Assert.IsType<FooterPart>(owner);
        else Assert.IsType<HeaderPart>(owner);
        RelationshipIds ids = smartArt._drawing.Descendants<RelationshipIds>().Single();
        Assert.IsType<DiagramDataPart>(owner.GetPartById(ids.DataPart!.Value!));
        Assert.IsType<DiagramLayoutDefinitionPart>(owner.GetPartById(ids.LayoutPart!.Value!));
        Assert.IsType<DiagramStylePart>(owner.GetPartById(ids.StylePart!.Value!));
        Assert.IsType<DiagramColorsPart>(owner.GetPartById(ids.ColorPart!.Value!));
        Assert.DoesNotContain(document.MainDocumentPartRoot.Parts, part => part.OpenXmlPart is
            DiagramDataPart or DiagramLayoutDefinitionPart or DiagramStylePart or DiagramColorsPart or DiagramPersistLayoutPart);
        if (type is WordSmartArtType.CustomSmartArt1 or WordSmartArtType.CustomSmartArt2) {
            var data = (DiagramDataPart)owner.GetPartById(ids.DataPart!.Value!);
            Dsp.DataModelExtensionBlock extension = data.DataModelRoot!.Descendants<Dsp.DataModelExtensionBlock>().Single();
            Assert.IsType<DiagramPersistLayoutPart>(owner.GetPartById(extension.RelId!.Value!));
            Assert.NotEmpty(smartArt.GetPersistedLayoutShapes());
        }
    }
}

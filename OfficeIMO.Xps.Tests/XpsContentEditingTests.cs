using System;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Xml.Linq;
using Xunit;
using static OfficeIMO.Xps.Tests.XpsLogicalStructureTests;

namespace OfficeIMO.Xps.Tests;

public sealed class XpsContentEditingTests {
    [Fact]
    public void BoundedDisplayDiagnosticsCannotHideDanglingNamesDuringEdits() {
        var doc = Create(XpsFormat.OpenXps, new[] { "A" }); var page = doc.Pages[0]; var ns = doc.StructureNamespace;
        var native = Fragments(doc, null, null, Paragraph(ns, "A"));
        native.AddFirst(Enumerable.Range(0, 100).Select(i => new XElement(XName.Get("Unknown" + i, "urn:extension"))));
        page.ReplaceStoryFragmentsMarkup(native); Assert.Equal(100, page.ReadContentStructure().Diagnostics.Count);
        byte[] before = doc.Save(); var markup = page.GetMarkup(); markup.Elements().Single().SetAttributeValue("Name", "renamed");
        Assert.Throws<InvalidDataException>(() => page.ReplaceMarkup(markup)); Assert.Equal(before, doc.Save());
        native.Descendants(ns + "NamedElement").Single().SetAttributeValue("NameReference", "missing");
        Assert.Throws<InvalidDataException>(() => page.ReplaceStoryFragmentsMarkup(native)); Assert.Equal(before, doc.Save());
        // The sequence-edit path must also reject malformed metadata from a
        // loaded producer package, even when the displayed diagnostic is capped.
        using var stream = new MemoryStream(); stream.Write(before, 0, before.Length);
        using (var zip = new ZipArchive(stream, ZipArchiveMode.Update, true)) {
            string part = doc.PartNames.Single(p => p.EndsWith(".frag")); zip.GetEntry(part)!.Delete();
            using var writer = new StreamWriter(zip.CreateEntry(part).Open()); writer.Write(native.ToString());
        }
        var loaded = XpsDocument.Load(stream.ToArray()); byte[] loadedBefore = loaded.Save();
        Assert.Throws<InvalidDataException>(() => loaded.AddPage()); Assert.Equal(loadedBefore, loaded.Save());
    }

    [Theory]
    [InlineData("root-text")]
    [InlineData("story-text")]
    [InlineData("reference-text")]
    [InlineData("reference-child")]
    public void DocumentStructureEditsRejectKnownContentThatCannotBeRead(string invalid) {
        var doc = Create(XpsFormat.OpenXps, new[] { "A" }); var ns = doc.StructureNamespace;
        doc.Pages[0].ReplaceStoryFragmentsMarkup(Fragments(doc, "body", "part", Paragraph(ns, "A")));
        var markup = Story(doc, "body", (1, "part")); var story = markup.Element(ns + "Story")!;
        var reference = story.Element(ns + "StoryFragmentReference")!;
        if (invalid == "root-text") markup.Add("literal");
        if (invalid == "story-text") story.Add("literal");
        if (invalid == "reference-text") reference.Add("literal");
        if (invalid == "reference-child") reference.Add(new XElement(XName.Get("Unknown", "urn:extension")));
        byte[] before = doc.Save(); Assert.Throws<InvalidDataException>(() => doc.Documents[0].ReplaceDocumentStructureMarkup(markup));
        Assert.Equal(before, doc.Save()); Assert.Null(doc.Documents[0].GetDocumentStructureMarkup());
    }
    [Fact]
    public void RepeatedNativeReferencesCannotAmplifyResolvedTextWithoutBound() {
        var doc = Create(XpsFormat.OpenXps, new[] { "A" }); var page = doc.Pages[0];
        var markup = page.GetMarkup(); markup.Elements().Single().SetAttributeValue("UnicodeString", new string('A', 1_000_000));
        page.ReplaceMarkup(markup); byte[] before = doc.Save();
        var native = Fragments(doc, null, null, Paragraph(doc.StructureNamespace, Enumerable.Repeat("A", 17).ToArray()));
        Assert.Throws<InvalidDataException>(() => page.ReplaceStoryFragmentsMarkup(native)); Assert.Equal(before, doc.Save());
    }
    [Theory]
    [InlineData(XpsFormat.Xps)]
    [InlineData(XpsFormat.OpenXps)]
    public void PageAndNamedReferencesCanBeEditedAtomically(XpsFormat format) {
        var doc = Create(format, new[] { "A" }); var page = doc.Pages[0]; var ns = doc.StructureNamespace;
        page.ReplaceStoryFragmentsMarkup(Fragments(doc, "body", "part", Paragraph(ns, "A")));
        doc.Documents[0].ReplaceDocumentStructureMarkup(Story(doc, "body", (1, "part")));
        byte[] before = doc.Save(); var markup = page.GetMarkup(); markup.Elements().Single().SetAttributeValue("Name", "B");
        Assert.Throws<InvalidDataException>(() => page.ReplaceMarkup(markup)); Assert.Equal(before, doc.Save());
        var fragments = page.GetStoryFragmentsMarkup()!;
        fragments.Descendants(ns + "NamedElement").Single().SetAttributeValue("NameReference", "B");
        page.ReplaceMarkup(markup, fragments);
        var reopened = XpsDocument.Load(doc.Save()); Assert.True(reopened.ReadLogicalStructure().IsComplete);
        Assert.Equal("B", Assert.Single(reopened.ReadLogicalStructure().Stories).Blocks[0].Children[0].Content!.Name);
        Assert.Equal("A", reopened.ReadLogicalStructure().Stories[0].Blocks[0].Text);
        Assert.Equal("body", (string?)reopened.Documents[0].GetDocumentStructureMarkup()!.Element(ns + "Story")!.Attribute("StoryName"));
    }

    [Fact]
    public void DeclaredFragmentRemovalAndChangingItToHeaderAreRejected() {
        var doc = Create(XpsFormat.OpenXps, new[] { "A" }); var page = doc.Pages[0]; var ns = doc.StructureNamespace;
        page.ReplaceStoryFragmentsMarkup(Fragments(doc, "body", "part", Paragraph(ns, "A")));
        doc.Documents[0].ReplaceDocumentStructureMarkup(Story(doc, "body", (1, "part"))); byte[] before = doc.Save();
        var markup = page.GetStoryFragmentsMarkup()!;
        markup.Element(ns + "StoryFragment")!.SetAttributeValue("FragmentName", "other");
        Assert.Throws<InvalidDataException>(() => page.ReplaceStoryFragmentsMarkup(markup));
        markup.Element(ns + "StoryFragment")!.SetAttributeValue("FragmentName", "part");
        markup.Element(ns + "StoryFragment")!.SetAttributeValue("FragmentType", "Header");
        Assert.Throws<InvalidDataException>(() => page.ReplaceStoryFragmentsMarkup(markup)); Assert.Equal(before, doc.Save());
        // Removing the owning story association first permits an intentional fragment rename.
        doc.Documents[0].ReplaceDocumentStructureMarkup(new XElement(ns + "DocumentStructure"));
        markup.Element(ns + "StoryFragment")!.SetAttributeValue("StoryName", null);
        page.ReplaceStoryFragmentsMarkup(markup); Assert.Equal(XpsStoryFragmentType.Header, Assert.Single(page.ReadContentStructure().Fragments).Type);
    }

    [Fact]
    public void PageMovesRemapDeclaredAddressesWithoutChangingTheLogicalStory() {
        var doc = Create(XpsFormat.OpenXps, new[] { "A" }, new[] { "B" }); var ns = doc.StructureNamespace;
        for (int i = 0; i < 2; i++) doc.Pages[i].ReplaceStoryFragmentsMarkup(Fragments(doc, "body", null, Paragraph(ns, i == 0 ? "A" : "B")));
        doc.Documents[0].ReplaceDocumentStructureMarkup(Story(doc, "body", (1, null), (2, null)));
        doc.Documents[0].MovePage(1, 0); var result = XpsDocument.Load(doc.Save()).ReadLogicalStructure();
        Assert.True(result.IsComplete); Assert.Equal("AB", Assert.Single(Assert.Single(result.Stories).Blocks).Text);
        Assert.Equal(new[] { 1, 0 }, result.Stories[0].Fragments.Select(f => f.PageIndex));
        doc.Documents[0].RemovePageAt(1); Assert.Equal("B", Assert.Single(Assert.Single(doc.ReadLogicalStructure().Stories).Blocks).Text);
    }

    [Theory]
    [InlineData("type")]
    [InlineData("nesting")]
    [InlineData("span")]
    [InlineData("name")]
    public void InvalidNativeMetadataNeverCommitsAPartOrRelationship(string invalid) {
        var doc = Create(XpsFormat.OpenXps, new[] { "A" }); var ns = doc.StructureNamespace;
        var native = Fragments(doc, null, null, Paragraph(ns, "A"));
        if (invalid == "type") native.Element(ns + "StoryFragment")!.Attribute("FragmentType")!.Remove();
        if (invalid == "nesting") native.Descendants(ns + "ParagraphStructure").Single().Add(new XElement(ns + "SectionStructure"));
        if (invalid == "span") native = Fragments(doc, null, null, new XElement(ns + "TableStructure", new XElement(ns + "TableRowGroupStructure",
            new XElement(ns + "TableRowStructure", new XElement(ns + "TableCellStructure", new XAttribute("RowSpan", "0"))))));
        if (invalid == "name") native.Descendants(ns + "NamedElement").Single().SetAttributeValue("NameReference", "missing");
        byte[] before = doc.Save(); Assert.Throws<InvalidDataException>(() => doc.Pages[0].ReplaceStoryFragmentsMarkup(native));
        Assert.Equal(before, doc.Save()); Assert.Null(doc.Pages[0].GetStoryFragmentsMarkup());
    }

    [Fact]
    public void NativeMetadataUsesOwningApisAndRejectsUnresolvedDocumentAddresses() {
        var doc = Create(XpsFormat.Xps, new[] { "A" }); var ns = doc.StructureNamespace;
        doc.Pages[0].ReplaceStoryFragmentsMarkup(Fragments(doc, "body", "part", Paragraph(ns, "A")));
        byte[] before = doc.Save();
        Assert.Throws<InvalidDataException>(() => doc.Documents[0].ReplaceDocumentStructureMarkup(Story(doc, "body", (1, "missing"))));
        Assert.Equal(before, doc.Save()); Assert.Null(doc.Documents[0].GetDocumentStructureMarkup());
        doc.Documents[0].ReplaceDocumentStructureMarkup(Story(doc, "body", (1, "part")));
        foreach (string part in doc.PartNames.Where(p => p.EndsWith(".frag") || p.EndsWith(".struct")))
            Assert.Throws<ArgumentException>(() => doc.ReplaceResource(part, doc.GetPartBytes(part)));
    }

    [Fact]
    public void SharedFragmentsReplacementMustRemainValidForEveryPageOwner() {
        var doc = Create(XpsFormat.Xps, new[] { "A", "B" }, new[] { "A" }); var ns = doc.StructureNamespace;
        doc.Pages[0].ReplaceStoryFragmentsMarkup(Fragments(doc, null, null, Paragraph(ns, "A")));
        var sharedPart = doc.PartNames.Single(p => p.EndsWith(".frag"));
        using var stream = new MemoryStream(); byte[] bytes = doc.Save(); stream.Write(bytes, 0, bytes.Length);
        var other = doc.Pages[1]; string relName = other.PartName.Replace("/Pages/", "/Pages/_rels/") + ".rels";
        using (var zip = new ZipArchive(stream, ZipArchiveMode.Update, true)) {
            zip.GetEntry(relName)?.Delete();
            using var writer = new StreamWriter(zip.CreateEntry(relName).Open());
            writer.Write(new XElement(XpsPackage.Relationships + "Relationships", new XElement(XpsPackage.Relationships + "Relationship",
                new XAttribute("Id", "story"), new XAttribute("Type", XpsPackage.Namespace(doc.Format) + "/storyfragments"), new XAttribute("Target", "/" + sharedPart))).ToString());
        }
        var loaded = XpsDocument.Load(stream.ToArray()); byte[] before = loaded.Save();
        var replacement = Fragments(loaded, null, null, Paragraph(ns, "B"));
        replacement.AddFirst(Enumerable.Range(0, 100).Select(i => new XElement(XName.Get("Unknown" + i, "urn:extension"))));
        Assert.Throws<InvalidDataException>(() => loaded.Pages[0].ReplaceStoryFragmentsMarkup(replacement));
        Assert.Equal(before, loaded.Save()); Assert.Equal("A", Assert.Single(loaded.Pages[1].ReadContentStructure().Fragments).Blocks[0].Text);
    }
}

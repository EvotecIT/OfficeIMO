using System;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Xunit;

namespace OfficeIMO.Xps.Tests;

public sealed class XpsDocumentStructureTests {
    private const string StructurePart = "Documents/1/Structure/DocumentStructure.struct";

    [Theory]
    [InlineData(XpsFormat.Xps)]
    [InlineData(XpsFormat.OpenXps)]
    public void StructuralEditsPreserveStoryPagesOutlineTargetsAndExtensions(XpsFormat format) {
        var doc = Fixture(format); var document = doc.Documents[0];
        string target = doc.Pages[1].PartName;
        document.InsertPage(0, 120, 90);
        Assert.Equal(new[] { "3", "4" }, StoryPages(doc));
        Assert.Equal("/" + target, (string?)Structure(doc).Descendants(doc.StructureNamespace + "OutlineEntry").Single().Attribute("OutlineTarget"));
        document.MovePage(3, 0);
        Assert.Equal(new[] { "4", "1" }, StoryPages(doc));
        document.RemovePageAt(3);
        Assert.Equal(new[] { "1" }, StoryPages(doc));
        var reopened = XpsDocument.Load(doc.Save());
        Assert.Equal(new[] { "1" }, StoryPages(reopened));
        Assert.Equal("keep", (string?)Structure(reopened).Attribute(XName.Get("custom", "urn:extension")));
        Assert.Equal("3", (string?)Structure(reopened).Element(XName.Get("StoryFragmentReference", "urn:extension"))!.Attribute("Page"));
        // A removed outline target remains explicitly unresolved, not redirected to
        // whichever page occupies its former numeric address.
        Assert.Equal("/" + target, (string?)Structure(reopened).Descendants(reopened.StructureNamespace + "OutlineEntry").Single().Attribute("OutlineTarget"));
    }

    [Fact]
    public void StoryPageNumbersRemainGlobalAcrossDocumentMovesAndSharedPages() {
        var doc = Fixture(XpsFormat.OpenXps); var first = doc.Documents[0];
        var second = doc.AddDocument(); second.InsertPage(0, first.Pages[1]);
        first.MovePageTo(2, second, 1);
        Assert.Equal(new[] { "2", "4" }, StoryPages(doc));
        doc.MoveDocument(1, 0);
        Assert.Equal(new[] { "1", "2" }, StoryPages(doc));
        second.RemovePageAt(0);
        Assert.Equal(new[] { "3", "1" }, StoryPages(doc));
    }

    [Theory]
    [InlineData("Page")]
    [InlineData("Root")]
    [InlineData("External")]
    public void InvalidKnownStructureRejectsEditsAtomically(string invalid) {
        var doc = Fixture(XpsFormat.Xps, invalid); byte[] before = doc.Save();
        Assert.Throws<InvalidDataException>(() => doc.Documents[0].MovePage(0, 1));
        Assert.Equal(before, doc.Save());
        Assert.Equal(3, doc.Pages.Count);
    }

    [Fact]
    public void RemovingLastStoryPageRemovesOnlyItsKnownReferenceAndEmptyStory() {
        var doc = Fixture(XpsFormat.OpenXps);
        doc.Documents[0].RemovePageAt(2); doc.Documents[0].RemovePageAt(1);
        Assert.Empty(Structure(doc).Elements(doc.StructureNamespace + "Story"));
        Assert.NotNull(Structure(doc).Element(XName.Get("StoryFragmentReference", "urn:extension")));
    }

    private static XElement Structure(XpsDocument doc) => XElement.Parse(Encoding.UTF8.GetString(doc.GetPartBytes(StructurePart)));
    private static string?[] StoryPages(XpsDocument doc) => Structure(doc).Elements(doc.StructureNamespace + "Story")
        .Elements(doc.StructureNamespace + "StoryFragmentReference").Select(e => (string?)e.Attribute("Page")).ToArray();

    internal static XpsDocument Fixture(XpsFormat format, string? invalid = null) {
        var doc = XpsDocument.Create(format); doc.AddPage(); doc.AddPage(); doc.AddPage();
        XNamespace ns = doc.StructureNamespace, ext = "urn:extension";
        var structure = new XElement(ns + (invalid == "Root" ? "Wrong" : "DocumentStructure"), new XAttribute(ext + "custom", "keep"),
            new XElement(ns + "DocumentStructure.Outline", new XElement(ns + "DocumentOutline",
                new XElement(ns + "OutlineEntry", new XAttribute("Description", "Second"), new XAttribute("OutlineLevel", "1"), new XAttribute("OutlineTarget", "/FixedDocumentSequence.fdseq#2")))),
            new XElement(ns + "Story", new XAttribute("StoryName", "body"),
                new XElement(ns + "StoryFragmentReference", new XAttribute("Page", invalid == "Page" ? "9" : "2"), new XAttribute("FragmentName", "partA")),
                new XElement(ns + "StoryFragmentReference", new XAttribute("Page", "3"))),
            new XElement(ext + "StoryFragmentReference", new XAttribute("Page", "3")));
        doc.AddResource(StructurePart, Encoding.UTF8.GetBytes(structure.ToString()), XpsPackage.Type("documentstructure"));
        using var output = new MemoryStream(); var bytes = doc.Save(); output.Write(bytes, 0, bytes.Length);
        using (var zip = new ZipArchive(output, ZipArchiveMode.Update, true)) {
            var relationship = new XElement(XpsPackage.Relationships + "Relationships", new XElement(XpsPackage.Relationships + "Relationship",
                new XAttribute("Id", "structure"), new XAttribute("Type", XpsPackage.Namespace(format) + "/documentstructure"),
                new XAttribute("Target", "Structure/DocumentStructure.struct"), new XAttribute("TargetMode", invalid == "External" ? "External" : "Internal")));
            using var writer = new StreamWriter(zip.CreateEntry("Documents/1/_rels/FixedDocument.fdoc.rels").Open()); writer.Write(relationship.ToString());
        }
        return XpsDocument.Load(output.ToArray());
    }
}

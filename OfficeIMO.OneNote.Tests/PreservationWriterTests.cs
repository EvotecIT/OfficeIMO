namespace OfficeIMO.OneNote.Tests;

public sealed class PreservationWriterTests {
    [Theory]
    [InlineData("testOneNote.one")]
    [InlineData("testOneNote2016.one")]
    [InlineData("testOneNoteEmbeddedWordDoc.one")]
    [InlineData("testOneNoteFromOffice365.one")]
    [InlineData("testOneNoteFromOffice365-2.one")]
    public void PublicFixtureRetainsCurrentPageSemanticsAcrossReadEditWriteRead(string fixture) {
        OneNoteSection source = OneNoteSectionReader.Read(FixturePath(fixture));
        string[] titles = source.Pages.Select(page => page.Title).ToArray();
        string[] text = source.Pages.Select(ExtractText).ToArray();
        string[][] versions = source.Pages
            .Select(page => page.VersionHistory.Select(version => version.Title + "\n" + ExtractText(version)).ToArray())
            .ToArray();
        string[][] conflicts = source.Pages
            .Select(page => page.ConflictPages.Select(conflict => conflict.Title + "\n" + ExtractText(conflict)).ToArray())
            .ToArray();
        int opaqueCount = source.UnknownObjects.Count + source.Pages.Sum(page => page.UnknownObjects.Count);

        byte[] rewritten = OneNoteSectionWriter.Write(source);
        OneNoteSection roundTrip = OneNoteSectionReader.Read(new MemoryStream(rewritten));

        Assert.Equal(titles, roundTrip.Pages.Select(page => page.Title).ToArray());
        Assert.Equal(text, roundTrip.Pages.Select(ExtractText).ToArray());
        for (int index = 0; index < roundTrip.Pages.Count; index++) {
            Assert.Equal(versions[index], roundTrip.Pages[index].VersionHistory.Select(version => version.Title + "\n" + ExtractText(version)).ToArray());
            Assert.Equal(conflicts[index], roundTrip.Pages[index].ConflictPages.Select(conflict => conflict.Title + "\n" + ExtractText(conflict)).ToArray());
        }
        Assert.True(roundTrip.UnknownObjects.Count + roundTrip.Pages.Sum(page => page.UnknownObjects.Count) >= opaqueCount);
    }

    [Fact]
    public void RoundTripPreservesUnknownGraphDataWithoutReattachingDeletedTypedContent() {
        OneNoteWriteGraph graph = new OneNoteWriteGraphBuilder().BuildSection(CreateSection());
        OneNoteWriteObjectSpace pageSpace = graph.ObjectSpaces[1];
        OneNoteWriteObject pageNode = pageSpace.Objects.Single(item => item.Jcid == OneNoteSchema.JcidPageNode);
        OneNoteExtendedGuid outlineId = Assert.Single(Property(pageNode, OneNoteSchema.ElementChildNodes).References);
        int outlineIndex = pageSpace.Objects.ToList().FindIndex(item => item.Id.Equals(outlineId));
        OneNoteWriteObject outline = pageSpace.Objects[outlineIndex];
        OneNoteWriteProperty relationship = Property(outline, OneNoteSchema.ElementChildNodes);
        OneNoteExtendedGuid typedChildId = Assert.Single(relationship.References);
        OneNoteExtendedGuid unknownId = new OneNoteExtendedGuid(new Guid("EFE27D87-7D73-48EC-833C-632E159AA3C6"), 1, 17);
        uint unknownScalarId = Raw(OneNotePropertyType.UInt32, 0x003F0101);

        var updatedProperties = outline.Properties
            .Where(property => PropertyKey(property) != PropertyKey(relationship))
            .Concat(new[] {
                new OneNoteWriteProperty(
                    relationship.RawId,
                    references: new[] { typedChildId, unknownId },
                    preserveRawId: true),
                new OneNoteWriteProperty(unknownScalarId, scalar: 0xC0FFEE, preserveRawId: true)
            });
        pageSpace.Objects[outlineIndex] = new OneNoteWriteObject(outline.Id, outline.Jcid, updatedProperties);
        pageSpace.Objects.Add(new OneNoteWriteObject(
            unknownId,
            0x000600FE,
            new[] { new OneNoteWriteProperty(Raw(OneNotePropertyType.UInt32, 0x003F0102), scalar: 42, preserveRawId: true) }));

        OneNoteSection loaded = OneNoteSectionReader.Read(new MemoryStream(OneNoteRevisionStoreWriter.Write(graph)));
        OneNotePage loadedPage = Assert.Single(loaded.Pages);
        OneNoteOutline loadedOutline = Assert.Single(loadedPage.Outlines);
        Assert.Single(loadedOutline.Children.OfType<OneNoteParagraph>());
        Assert.Contains(loadedPage.UnknownObjects, item => unknownId.Equals(item.Id));

        loadedOutline.Children.Clear();
        loadedPage.Title = "Edited while preserving opaque data";
        byte[] rewritten = OneNoteSectionWriter.Write(loaded);

        OneNoteRevisionStore store = OneNoteRevisionStoreReader.Read(new MemoryStream(rewritten));
        var materializer = new OneNoteObjectSpaceMaterializer(store);
        OneNoteMaterializedObjectSpace rewrittenPageSpace = Assert.IsType<OneNoteMaterializedObjectSpace>(materializer.TryGetCurrentSpace(loadedPage.Id!));
        OneNoteRevisionStoreObject rewrittenOutline = Assert.IsType<OneNoteRevisionStoreObject>(rewrittenPageSpace.GetObject(outlineId));
        OneNotePropertyValue unknownScalar = Assert.Single(
            rewrittenOutline.PropertySet!.Properties,
            property => property.RawPropertyId == unknownScalarId);
        Assert.Equal(0xC0FFEEUL, unknownScalar.ScalarValue);
        Assert.Equal(unknownId, Assert.Single(OneNoteSemanticMapper.GetReferences(rewrittenOutline, OneNoteSchema.ElementChildNodes)));
        Assert.Contains(rewrittenPageSpace.Objects, item => item.Id.Equals(unknownId) && item.Jcid.Value == 0x000600FE);

        OneNoteSection roundTrip = OneNoteSectionReader.Read(new MemoryStream(rewritten));
        Assert.Equal("Edited while preserving opaque data", Assert.Single(roundTrip.Pages).Title);
        Assert.Empty(Assert.Single(Assert.Single(roundTrip.Pages).Outlines).Children);
        Assert.Contains(Assert.Single(roundTrip.Pages).UnknownObjects, item => unknownId.Equals(item.Id));
    }

    private static OneNoteSection CreateSection() {
        var section = new OneNoteSection { Name = "Preservation" };
        var page = new OneNotePage { Title = "Before" };
        var outline = new OneNoteOutline();
        var paragraph = new OneNoteParagraph();
        paragraph.Runs.Add(new OneNoteTextRun { Text = "Typed content to delete" });
        outline.Children.Add(paragraph);
        page.Outlines.Add(outline);
        section.Pages.Add(page);
        return section;
    }

    private static string ExtractText(OneNotePage page) => string.Join("\n", page.Outlines
        .Cast<OneNoteElement>()
        .Concat(page.DirectContent)
        .SelectMany(ExtractText));

    private static IEnumerable<string> ExtractText(OneNoteElement element) {
        if (element is OneNoteParagraph paragraph) {
            foreach (OneNoteTextRun run in paragraph.Runs) yield return run.Text;
            foreach (OneNoteElement child in paragraph.Children)
            foreach (string text in ExtractText(child)) yield return text;
        } else if (element is OneNoteOutline outline) {
            foreach (OneNoteElement child in outline.Children)
            foreach (string text in ExtractText(child)) yield return text;
        } else if (element is OneNoteTable table) {
            foreach (OneNoteTableRow row in table.Rows)
            foreach (OneNoteTableCell cell in row.Cells)
            foreach (OneNoteElement child in cell.Content)
            foreach (string text in ExtractText(child)) yield return text;
        }
    }

    private static string FixturePath(string fileName) => Path.Combine(AppContext.BaseDirectory, "Fixtures", fileName);

    private static OneNoteWriteProperty Property(OneNoteWriteObject item, uint propertyId) =>
        item.Properties.Single(property => PropertyKey(property) == (propertyId & 0x03FFFFFFU));

    private static uint PropertyKey(OneNoteWriteProperty property) => property.RawId & 0x03FFFFFFU;

    private static uint Raw(OneNotePropertyType type, uint id) => ((uint)type << 26) | (id & 0x03FFFFFFU);
}

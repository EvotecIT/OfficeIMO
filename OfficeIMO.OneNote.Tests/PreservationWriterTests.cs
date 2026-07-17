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
    public void SharedNativePageSeriesRetainsEveryCurrentAndVersionPage() {
        var section = new OneNoteSection { Name = "Shared series" };
        for (int index = 0; index < 3; index++) {
            var page = new OneNotePage { Title = "Page " + index };
            var outline = new OneNoteOutline();
            var paragraph = new OneNoteParagraph();
            paragraph.Runs.Add(new OneNoteTextRun { Text = "Content " + index });
            outline.Children.Add(paragraph);
            page.Outlines.Add(outline);
            section.Pages.Add(page);
        }
        section.Pages[1].VersionHistory.Add(new OneNotePage { Title = "Page 1 previous" });

        OneNoteWriteGraph graph = new OneNoteWriteGraphBuilder().BuildSection(section);
        OneNoteWriteObjectSpace sectionSpace = graph.ObjectSpaces[0];
        OneNoteWriteObject[] series = sectionSpace.Objects
            .Where(item => item.Jcid == OneNoteSchema.JcidPageSeriesNode)
            .ToArray();
        Assert.Equal(3, series.Length);
        OneNoteExtendedGuid[] pageSpaceIds = series
            .SelectMany(item => Property(item, OneNoteSchema.ChildGraphSpaceElementNodes).References)
            .ToArray();
        OneNoteExtendedGuid[] cachedMetadataIds = series
            .SelectMany(item => Property(item, OneNoteSchema.MetaDataObjectsAboveGraphSpace).References)
            .ToArray();

        OneNoteWriteObject sharedSeries = series[0];
        OneNoteWriteProperty[] sharedProperties = sharedSeries.Properties.Select(property => {
            uint key = PropertyKey(property);
            if (key == (OneNoteSchema.ChildGraphSpaceElementNodes & 0x03FFFFFFU)) {
                return new OneNoteWriteProperty(
                    property.RawId,
                    references: pageSpaceIds,
                    referenceKind: OneNoteWriteReferenceKind.ObjectSpace,
                    preserveRawId: true);
            }
            if (key == (OneNoteSchema.MetaDataObjectsAboveGraphSpace & 0x03FFFFFFU)) {
                return new OneNoteWriteProperty(
                    property.RawId,
                    references: cachedMetadataIds,
                    preserveRawId: true);
            }
            return property;
        }).ToArray();
        int sharedSeriesIndex = sectionSpace.Objects.ToList().FindIndex(item => item.Id.Equals(sharedSeries.Id));
        sectionSpace.Objects[sharedSeriesIndex] = new OneNoteWriteObject(sharedSeries.Id, sharedSeries.Jcid, sharedProperties);
        foreach (OneNoteWriteObject duplicate in series.Skip(1)) sectionSpace.Objects.Remove(duplicate);

        int sectionRootIndex = sectionSpace.Objects.ToList().FindIndex(item => item.Jcid == OneNoteSchema.JcidSectionNode);
        OneNoteWriteObject sectionRoot = sectionSpace.Objects[sectionRootIndex];
        OneNoteWriteProperty[] rootProperties = sectionRoot.Properties.Select(property =>
            PropertyKey(property) == (OneNoteSchema.ElementChildNodes & 0x03FFFFFFU)
                ? new OneNoteWriteProperty(property.RawId, references: new[] { sharedSeries.Id }, preserveRawId: true)
                : property).ToArray();
        sectionSpace.Objects[sectionRootIndex] = new OneNoteWriteObject(sectionRoot.Id, sectionRoot.Jcid, rootProperties);

        byte[] native = OneNoteRevisionStoreWriter.Write(graph);
        OneNoteSection loaded = OneNoteSectionReader.Read(new MemoryStream(native));
        Assert.Equal(new[] { "Page 0", "Page 1", "Page 2" }, loaded.Pages.Select(page => page.Title));
        Assert.Equal("Page 1 previous", Assert.Single(loaded.Pages[1].VersionHistory).Title);

        byte[] rewritten = OneNoteSectionWriter.Write(loaded);
        OneNoteSection roundTrip = OneNoteSectionReader.Read(new MemoryStream(rewritten));

        Assert.Equal(new[] { "Page 0", "Page 1", "Page 2" }, roundTrip.Pages.Select(page => page.Title));
        Assert.Equal("Page 1 previous", Assert.Single(roundTrip.Pages[1].VersionHistory).Title);
        AssertNativeSeriesLayout(rewritten, new[] { 3 });

        OneNoteSection edited = OneNoteSectionReader.Read(new MemoryStream(native));
        edited.Pages.Insert(1, new OneNotePage { Title = "Inserted" });
        byte[] editedBytes = OneNoteSectionWriter.Write(edited);
        OneNoteSection editedRoundTrip = OneNoteSectionReader.Read(new MemoryStream(editedBytes));
        Assert.Equal(
            new[] { "Page 0", "Inserted", "Page 1", "Page 2" },
            editedRoundTrip.Pages.Select(page => page.Title));
        Assert.Equal("Page 1 previous", Assert.Single(editedRoundTrip.Pages[2].VersionHistory).Title);
        AssertNativeSeriesLayout(editedBytes, new[] { 1, 1, 2 });

        var notebook = new OneNoteNotebook { Name = "Shared series package" };
        notebook.Sections.Add(OneNoteSectionReader.Read(new MemoryStream(native)));
        OneNoteNotebook packageRoundTrip = OneNotePackageReader.Read(
            new MemoryStream(OneNotePackageWriter.Write(notebook)),
            "shared-series.onepkg");
        OneNoteSection packagedSection = Assert.Single(packageRoundTrip.Sections);
        Assert.Equal(new[] { "Page 0", "Page 1", "Page 2" }, packagedSection.Pages.Select(page => page.Title));
        Assert.Equal("Page 1 previous", Assert.Single(packagedSection.Pages[1].VersionHistory).Title);
    }

    [Fact]
    public void RelationshipMembershipDefinesConflictAndVersionSemanticsDuringDefaultValidation() {
        var section = new OneNoteSection { Name = "Relationships" };
        var current = new OneNotePage { Title = "Current" };
        current.ConflictPages.Add(new OneNotePage { Title = "Conflict without flag" });
        current.VersionHistory.Add(new OneNotePage { Title = "Version without flag" });
        section.Pages.Add(current);

        OneNoteSection roundTrip = OneNoteSectionReader.Read(
            new MemoryStream(OneNoteSectionWriter.Write(section)));

        OneNotePage page = Assert.Single(roundTrip.Pages);
        Assert.True(Assert.Single(page.ConflictPages).IsConflictPage);
        Assert.True(Assert.Single(page.VersionHistory).IsVersionHistoryPage);
    }

    [Fact]
    public void RoundTripValidatorRejectsTableTopologyLoss() {
        var expected = new OneNoteSection { Name = "Topology" };
        var expectedPage = new OneNotePage { Title = "Page" };
        var expectedOutline = new OneNoteOutline();
        var expectedTable = new OneNoteTable();
        var expectedRow = new OneNoteTableRow();
        expectedRow.Cells.Add(new OneNoteTableCell());
        expectedRow.Cells.Add(new OneNoteTableCell());
        expectedTable.Rows.Add(expectedRow);
        expectedOutline.Children.Add(expectedTable);
        expectedPage.Outlines.Add(expectedOutline);
        expected.Pages.Add(expectedPage);

        var actual = new OneNoteSection { Name = "Topology" };
        var actualPage = new OneNotePage { Title = "Page" };
        var actualOutline = new OneNoteOutline();
        var actualTable = new OneNoteTable();
        var actualRow = new OneNoteTableRow();
        actualRow.Cells.Add(new OneNoteTableCell());
        actualTable.Rows.Add(actualRow);
        actualOutline.Children.Add(actualTable);
        actualPage.Outlines.Add(actualOutline);
        actual.Pages.Add(actualPage);

        OneNoteFormatException exception = Assert.Throws<OneNoteFormatException>(
            () => OneNoteWriteRoundTripValidator.ValidateSection(expected, actual));
        Assert.Equal("ONENOTE_WRITE_ROUNDTRIP_SEMANTICS", exception.Code);
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

    private static void AssertNativeSeriesLayout(byte[] data, int[] expectedRunLengths) {
        OneNoteSection section = OneNoteSectionReader.Read(new MemoryStream(data));
        OneNoteSectionPreservationState preservation = Assert.IsType<OneNoteSectionPreservationState>(section.PreservationState);
        OneNoteMaterializedObjectSpace sectionSpace = preservation.SectionSpace;
        OneNoteRevisionStoreObject sectionRoot = Assert.IsType<OneNoteRevisionStoreObject>(sectionSpace.GetRoot(1));
        OneNoteRevisionStoreObject[] series = OneNoteSemanticMapper
            .GetReferences(sectionRoot, OneNoteSchema.ElementChildNodes)
            .Select(id => Assert.IsType<OneNoteRevisionStoreObject>(sectionSpace.GetObject(id)))
            .ToArray();

        Assert.Equal(expectedRunLengths, series
            .Select(item => OneNoteSemanticMapper.GetReferences(item, OneNoteSchema.ChildGraphSpaceElementNodes).Count)
            .ToArray());

        OneNoteExtendedGuid[] pageIds = series
            .SelectMany(item => OneNoteSemanticMapper.GetReferences(item, OneNoteSchema.ChildGraphSpaceElementNodes))
            .ToArray();
        OneNoteExtendedGuid[] metadataIds = series
            .SelectMany(item => OneNoteSemanticMapper.GetReferences(item, OneNoteSchema.MetaDataObjectsAboveGraphSpace))
            .ToArray();
        Assert.Equal(section.Pages.Select(page => page.Id!), pageIds);
        Assert.Equal(section.Pages.Select(page => preservation.GetCachedPageMetadataId(page)!), metadataIds);
        Assert.Equal(metadataIds.Length, metadataIds.Distinct().Count());
        Assert.Equal(
            section.Pages.Select(page => page.Title),
            metadataIds.Select(id => ReadUnicodeProperty(
                Assert.IsType<OneNoteRevisionStoreObject>(sectionSpace.GetObject(id)),
                OneNoteSchema.CachedTitleString)));
    }

    private static string? ReadUnicodeProperty(OneNoteRevisionStoreObject item, uint propertyId) {
        OneNotePropertyValue? property = item.PropertySet?.Properties.LastOrDefault(
            value => value.Id == (propertyId & 0x03FFFFFFU));
        if (property?.Data == null) return null;
        byte[] data = property.Data.ToArray(1024 * 1024);
        return System.Text.Encoding.Unicode.GetString(data).TrimEnd('\0');
    }

    private static OneNoteWriteProperty Property(OneNoteWriteObject item, uint propertyId) =>
        item.Properties.Single(property => PropertyKey(property) == (propertyId & 0x03FFFFFFU));

    private static uint PropertyKey(OneNoteWriteProperty property) => property.RawId & 0x03FFFFFFU;

    private static uint Raw(OneNotePropertyType type, uint id) => ((uint)type << 26) | (id & 0x03FFFFFFU);
}

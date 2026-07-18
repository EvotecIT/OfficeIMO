namespace OfficeIMO.OneNote.Tests;

public sealed class NotebookWriterTests {
    [Fact]
    public void TableOfContentsUsesMsoTransactionChecksum() {
        byte[] data = OneNoteTableOfContentsWriter.Write(CreateNotebook());
        OneNoteFileHeader header = OneNoteFileProbe.ReadHeader(new MemoryStream(data));
        OneNoteFileChunkReference transaction = Assert.IsType<OneNoteFileChunkReference>(header.TransactionLog);
        int offset = checked((int)transaction.Offset);
        int sentinelOffset = offset;
        while (BitConverter.ToUInt32(data, sentinelOffset) != 1U) sentinelOffset += 8;

        uint expected = OneNoteCrc32.ComputeMso(data.Skip(offset).Take(sentinelOffset - offset).ToArray());
        Assert.Equal(expected, BitConverter.ToUInt32(data, sentinelOffset + 4));
        OneNoteRevisionStoreReader.Read(new MemoryStream(data));
    }

    [Fact]
    public void TableOfContentsEmitsDependencyOverridesAfterObjectDeclarations() {
        OneNoteRevisionStore store = OneNoteRevisionStoreReader.Read(
            new MemoryStream(OneNoteTableOfContentsWriter.Write(CreateNotebook())));
        OneNoteFileNodeList revision = Assert.Single(store.FileNodeLists, list =>
            list.Nodes.Any(node => node.Id == OneNoteFileNodeId.RevisionManifestStart4));

        int lastDeclaration = revision.Nodes
            .Select((node, index) => new { node, index })
            .Where(item => item.node.Id == OneNoteFileNodeId.ObjectDeclarationWithRefCount || item.node.Id == OneNoteFileNodeId.ObjectDeclarationWithRefCount2)
            .Max(item => item.index);
        int dependencies = revision.Nodes
            .Select((node, index) => new { node, index })
            .Single(item => item.node.Id == OneNoteFileNodeId.ObjectInfoDependencyOverrides)
            .index;
        int firstRoot = revision.Nodes
            .Select((node, index) => new { node, index })
            .Where(item => item.node.Id == OneNoteFileNodeId.RootObjectReference2)
            .Min(item => item.index);

        Assert.True(lastDeclaration < dependencies);
        Assert.True(dependencies < firstRoot);
    }

    [Theory]
    [InlineData(OneNoteStorageFormat.RevisionStore)]
    [InlineData(OneNoteStorageFormat.FileSynchronizationPackage)]
    public void TableOfContentsRoundTripsRootHierarchy(OneNoteStorageFormat storageFormat) {
        OneNoteNotebook original = CreateNotebook();

        byte[] data = OneNoteTableOfContentsWriter.Write(original, new OneNoteWriterOptions {
            StorageFormat = storageFormat
        });
        OneNoteNotebook result = OneNoteNotebookReader.Read(new MemoryStream(data));

        OneNoteFileHeader header = OneNoteFileProbe.ReadHeader(new MemoryStream(data));
        Assert.Equal(OneNoteFileKind.TableOfContents, header.FileKind);
        Assert.Equal(storageFormat, header.StorageFormat);
        Assert.Equal(storageFormat, result.TableOfContentsStorageFormat);
        Assert.Equal(original.Id, result.Id);
        Assert.Equal(original.Id, header.FileId);
        Assert.Equal("Root", Assert.Single(result.Sections).Name);
        Assert.Equal("Group", Assert.Single(result.SectionGroups).Name);
    }

    [Fact]
    public void LoadedFssHttpTableOfContentsPreservesItsStorageFormatByDefault() {
        byte[] source = OneNoteTableOfContentsWriter.Write(CreateNotebook(), new OneNoteWriterOptions {
            StorageFormat = OneNoteStorageFormat.FileSynchronizationPackage
        });
        OneNoteNotebook loaded = OneNoteNotebookReader.Read(new MemoryStream(source));

        byte[] rewritten = OneNoteTableOfContentsWriter.Write(loaded);

        OneNoteFileHeader header = OneNoteFileProbe.ReadHeader(new MemoryStream(rewritten));
        OneNoteNotebook result = OneNoteNotebookReader.Read(new MemoryStream(rewritten));
        Assert.Equal(OneNoteStorageFormat.FileSynchronizationPackage, header.StorageFormat);
        Assert.Equal(OneNoteStorageFormat.FileSynchronizationPackage, result.TableOfContentsStorageFormat);
        Assert.Equal(loaded.Id, result.Id);
        Assert.Equal("Root", Assert.Single(result.Sections).Name);
        Assert.Equal("Group", Assert.Single(result.SectionGroups).Name);
    }

    [Fact]
    public void PackageRoundTripsNestedNotebookAndUsesManagedCabinet() {
        OneNoteNotebook original = CreateNotebook();

        byte[] data = OneNotePackageWriter.Write(original);
        OneNoteNotebook result = OneNotePackageReader.Read(new MemoryStream(data), "Writer.onepkg");

        Assert.Equal(new byte[] { (byte)'M', (byte)'S', (byte)'C', (byte)'F' }, data.Take(4));
        Assert.Equal(original.Id, result.Id);
        Assert.Equal("Root page", Assert.Single(Assert.Single(result.Sections).Pages).Title);
        OneNoteSectionGroup group = Assert.Single(result.SectionGroups);
        Assert.Equal("Nested page", Assert.Single(Assert.Single(group.Sections).Pages).Title);
        Assert.Empty(result.Diagnostics);
    }

    [Fact]
    public void NotebookReadersPreserveNativeSectionDisplayNameInsteadOfSanitizedTocName() {
        const string displayName = "Roadmap: Q1";
        OneNoteNotebook original = CreateNotebook();
        original.Sections[0].Name = displayName;

        byte[] package = OneNotePackageWriter.Write(original);
        OneNoteNotebook packaged = OneNotePackageReader.Read(new MemoryStream(package), "Display.onepkg");

        Assert.Equal(displayName, Assert.Single(packaged.Sections).Name);

        string root = Path.Combine(Path.GetTempPath(), "OfficeIMO-OneNote-Display-" + Guid.NewGuid().ToString("N"));
        try {
            OneNoteNotebookWriter.Write(original, root);
            Assert.True(File.Exists(Path.Combine(root, "Roadmap_ Q1.one")));

            OneNoteNotebook directory = OneNoteNotebookReader.Read(Path.Combine(root, "Open Notebook.onetoc2"));
            Assert.Equal(displayName, Assert.Single(directory.Sections).Name);
        } finally {
            if (Directory.Exists(root)) Directory.Delete(root, true);
        }
    }

    [Fact]
    public void PackageWriterUsesRecycleBinSemanticFlagAndReaderKeepsItOptIn() {
        OneNoteNotebook original = CreateNotebook();
        original.SectionGroups.Add(new OneNoteSectionGroup {
            Name = "Deleted items",
            IsRecycleBin = true
        });

        byte[] data = OneNotePackageWriter.Write(original);
        OneNoteNotebook currentOnly = OneNotePackageReader.Read(new MemoryStream(data), "Recycle.onepkg");
        OneNoteNotebook withRecycleBin = OneNotePackageReader.Read(
            new MemoryStream(data),
            "Recycle.onepkg",
            new OneNoteNotebookReaderOptions { IncludeRecycleBin = true });

        Assert.DoesNotContain(currentOnly.SectionGroups, group => group.IsRecycleBin);
        OneNoteSectionGroup recycleBin = Assert.Single(withRecycleBin.SectionGroups, group => group.IsRecycleBin);
        Assert.Equal("OneNote_RecycleBin", recycleBin.Name);
        Assert.Empty(recycleBin.Sections);
        Assert.Empty(recycleBin.SectionGroups);
        Assert.DoesNotContain(withRecycleBin.Diagnostics, diagnostic => diagnostic.Code == "ONENOTE_TOC_GROUP_MISSING");
    }

    [Fact]
    public void DirectoryWriterCreatesNativeHierarchyAndRefusesExistingContent() {
        string root = Path.Combine(Path.GetTempPath(), "OfficeIMO-OneNote-" + Guid.NewGuid().ToString("N"));
        try {
            OneNoteNotebook original = CreateNotebook();
            OneNoteNotebookWriter.Write(original, root);

            OneNoteNotebook result = OneNoteNotebookReader.Read(Path.Combine(root, "Open Notebook.onetoc2"));
            Assert.Equal(original.Id, result.Id);
            Assert.Equal("Root page", Assert.Single(Assert.Single(result.Sections).Pages).Title);
            Assert.True(File.Exists(Path.Combine(root, "Group", "Open Notebook.onetoc2")));
            Assert.Throws<IOException>(() => OneNoteNotebookWriter.Write(CreateNotebook(), root));
        } finally {
            if (Directory.Exists(root)) Directory.Delete(root, true);
        }
    }

    [Fact]
    public void PackageWriterRejectsCyclicSectionGroupsBeforeDescending() {
        var notebook = new OneNoteNotebook { Name = "Cycle" };
        var group = new OneNoteSectionGroup { Name = "Loop" };
        group.SectionGroups.Add(group);
        notebook.SectionGroups.Add(group);

        OneNoteFormatException exception = Assert.Throws<OneNoteFormatException>(() =>
            OneNotePackageWriter.Write(notebook, new OneNoteWriterOptions { ValidateRoundTrip = false }));

        Assert.Equal("ONENOTE_WRITE_GROUP_CYCLE", exception.Code);
    }

    [Fact]
    public void PackageWriterRejectsSharedSectionGroupInstances() {
        var notebook = new OneNoteNotebook { Name = "Shared" };
        var group = new OneNoteSectionGroup { Name = "Repeated" };
        notebook.SectionGroups.Add(group);
        notebook.SectionGroups.Add(group);

        OneNoteFormatException exception = Assert.Throws<OneNoteFormatException>(() =>
            OneNotePackageWriter.Write(notebook, new OneNoteWriterOptions { ValidateRoundTrip = false }));

        Assert.Equal("ONENOTE_WRITE_SHARED_GROUP", exception.Code);
    }

    [Fact]
    public void PackageWriterRejectsSectionGroupDepthPastTheConfiguredLimit() {
        OneNoteNotebook notebook = CreateGroupChain(4);

        OneNoteFormatException exception = Assert.Throws<OneNoteFormatException>(() =>
            OneNotePackageWriter.Write(notebook, new OneNoteWriterOptions {
                MaxSectionGroupDepth = 3,
                ValidateRoundTrip = false
            }));

        Assert.Equal("ONENOTE_WRITE_GROUP_DEPTH", exception.Code);
    }

    [Fact]
    public void PackageWriterRoundTripUsesTheConfiguredSectionGroupDepth() {
        int depth = OneNoteNotebookReaderOptions.DefaultMaxSectionGroupDepth + 1;

        byte[] package = OneNotePackageWriter.Write(CreateGroupChain(depth), new OneNoteWriterOptions {
            MaxSectionGroupDepth = depth
        });

        Assert.NotEmpty(package);
    }

    [Fact]
    public void WritersReserveTheTocFileNameWhenAllocatingGroupDirectories() {
        var notebook = new OneNoteNotebook { Name = "Reserved" };
        var group = new OneNoteSectionGroup { Name = "Open Notebook.onetoc2" };
        group.Sections.Add(CreateSection("Nested", "Nested page"));
        notebook.SectionGroups.Add(group);

        byte[] package = OneNotePackageWriter.Write(notebook);
        string[] packageNames = OneNoteCabinetArchiveReader
            .Read(package, 16 * 1024 * 1024, 16 * 1024 * 1024, 10)
            .Select(entry => entry.Name.Replace('\\', '/'))
            .ToArray();

        Assert.Contains("Open Notebook.onetoc2", packageNames);
        Assert.Contains("Open Notebook (2).onetoc2/Open Notebook.onetoc2", packageNames);
        Assert.DoesNotContain(packageNames, name => name.StartsWith("Open Notebook.onetoc2/", StringComparison.OrdinalIgnoreCase));

        string root = Path.Combine(Path.GetTempPath(), "OfficeIMO-OneNote-Reserved-" + Guid.NewGuid().ToString("N"));
        try {
            OneNoteNotebookWriter.Write(notebook, root);

            Assert.True(File.Exists(Path.Combine(root, "Open Notebook.onetoc2")));
            Assert.True(Directory.Exists(Path.Combine(root, "Open Notebook (2).onetoc2")));
            Assert.True(File.Exists(Path.Combine(root, "Open Notebook (2).onetoc2", "Open Notebook.onetoc2")));
        } finally {
            if (Directory.Exists(root)) Directory.Delete(root, true);
        }
    }

    [Fact]
    public void PackageWriterRejectsKnownLazyPayloadAgainstRemainingAggregateBudgetWithoutOpeningIt() {
        OneNoteSection first = CreateSection("First", "First page");
        long firstSectionBytes = OneNoteSectionWriter.Write(first, new OneNoteWriterOptions {
            ValidateRoundTrip = false
        }).LongLength;
        bool opened = false;
        OneNoteSection second = CreateSection("Second", "Second page");
        second.Pages[0].DirectContent.Add(new OneNoteEmbeddedFile {
            FileName = "bounded.bin",
            Payload = OneNoteBinaryPayload.FromStreamFactory(() => {
                opened = true;
                return new MemoryStream(new byte[64]);
            }, 64)
        });
        var notebook = new OneNoteNotebook { Name = "Bounded" };
        notebook.Sections.Add(first);
        notebook.Sections.Add(second);

        IOException exception = Assert.Throws<IOException>(() => OneNotePackageWriter.Write(notebook, new OneNoteWriterOptions {
            MaxOutputBytes = firstSectionBytes + 32,
            ValidateRoundTrip = false
        }));

        Assert.Contains("limit", exception.Message, StringComparison.OrdinalIgnoreCase);
        Assert.False(opened);
    }

    [Fact]
    public void NotebookReaderTreatsCaseDistinctTableOfContentsPathsSeparatelyOnCaseSensitivePlatforms() {
        if (Path.DirectorySeparatorChar == '\\') return;

        string root = Path.Combine(Path.GetTempPath(), "OfficeIMO-OneNote-CaseSensitive-" + Guid.NewGuid().ToString("N"));
        try {
            Directory.CreateDirectory(root);
            Guid firstGroupId = Guid.NewGuid();
            Guid secondGroupId = Guid.NewGuid();
            var first = new OneNoteNotebook { Id = firstGroupId, Name = "Projects" };
            first.Sections.Add(CreateSection("Upper", "Upper page"));
            var second = new OneNoteNotebook { Id = secondGroupId, Name = "projects" };
            second.Sections.Add(CreateSection("Lower", "Lower page"));
            string upperPath = Path.Combine(root, "Projects");
            string lowerPath = Path.Combine(root, "projects");
            OneNoteNotebookWriter.Write(first, upperPath);
            if (Directory.Exists(lowerPath)) return;
            OneNoteNotebookWriter.Write(second, lowerPath);

            OneNoteWriteGraph rootToc = new OneNoteWriteGraphBuilder().BuildTableOfContents(
                Guid.NewGuid(),
                Guid.Empty,
                "Open Notebook.onetoc2",
                new[] {
                    new OneNoteTocWriteEntry(firstGroupId, "Projects", 0, null),
                    new OneNoteTocWriteEntry(secondGroupId, "projects", 1, null)
                },
                null,
                null);
            File.WriteAllBytes(
                Path.Combine(root, "Open Notebook.onetoc2"),
                OneNoteRevisionStoreWriter.Write(rootToc));

            OneNoteNotebook result = OneNoteNotebookReader.Read(Path.Combine(root, "Open Notebook.onetoc2"));

            Assert.Equal(new[] { "Projects", "projects" }, result.SectionGroups.Select(group => group.Name).ToArray());
            Assert.Equal("Upper", Assert.Single(result.SectionGroups[0].Sections).Name);
            Assert.Equal("Lower", Assert.Single(result.SectionGroups[1].Sections).Name);
            Assert.DoesNotContain(result.Diagnostics, diagnostic => diagnostic.Code == "ONENOTE_TOC_CYCLE");
        } finally {
            if (Directory.Exists(root)) Directory.Delete(root, true);
        }
    }

    [Fact]
    public void NotebookPathResolutionRejectsTraversalThroughCaseVariantSibling() {
        string parent = Path.Combine(Path.GetTempPath(), "OfficeIMO-Notebook-MixedCase");
        string escaped = ".." + Path.DirectorySeparatorChar +
                         "OFFICEIMO-NOTEBOOK-MIXEDCASE" + Path.DirectorySeparatorChar +
                         "private.one";

        Assert.Null(OneNoteNotebookReader.ResolveChildPath(parent, escaped));
        Assert.Equal(
            Path.GetFullPath(Path.Combine(parent, "Section.one")),
            OneNoteNotebookReader.ResolveChildPath(parent, "Section.one"));
    }

#if NET8_0_OR_GREATER
    [Fact]
    public void NotebookReaderRejectsSymbolicLinkSectionsAndGroups() {
        string root = Path.Combine(Path.GetTempPath(), "OfficeIMO-OneNote-Links-" + Guid.NewGuid().ToString("N"));
        string outside = Path.Combine(Path.GetTempPath(), "OfficeIMO-OneNote-Outside-" + Guid.NewGuid().ToString("N"));
        try {
            OneNoteNotebookWriter.Write(CreateNotebook(), root);
            Directory.CreateDirectory(outside);
            string outsideSection = Path.Combine(outside, "Private.one");
            File.WriteAllBytes(outsideSection, OneNoteSectionWriter.Write(CreateSection("Private", "External content")));
            string outsideGroup = Path.Combine(outside, "PrivateGroup");
            var outsideNotebook = new OneNoteNotebook { Name = "Private group" };
            outsideNotebook.Sections.Add(CreateSection("Private nested", "External nested content"));
            OneNoteNotebookWriter.Write(outsideNotebook, outsideGroup);

            string linkedSection = Path.Combine(root, "Root.one");
            string linkedGroup = Path.Combine(root, "Group");
            File.Delete(linkedSection);
            Directory.Delete(linkedGroup, true);
            File.CreateSymbolicLink(linkedSection, outsideSection);
            Directory.CreateSymbolicLink(linkedGroup, outsideGroup);

            OneNoteNotebook result = OneNoteNotebookReader.Read(Path.Combine(root, "Open Notebook.onetoc2"));

            Assert.Empty(result.Sections);
            Assert.Empty(result.SectionGroups);
            Assert.Equal(2, result.Diagnostics.Count(diagnostic => diagnostic.Code == "ONENOTE_TOC_PATH"));
            Assert.Null(OneNoteNotebookReader.ResolveChildPath(root, "Root.one"));
            Assert.Null(OneNoteNotebookReader.ResolveChildPath(root, "Group"));
        } finally {
            if (Directory.Exists(root)) Directory.Delete(root, true);
            if (Directory.Exists(outside)) Directory.Delete(outside, true);
        }
    }

    [Fact]
    public void NotebookReaderRejectsSymbolicLinkNestedTableOfContents() {
        string root = Path.Combine(Path.GetTempPath(), "OfficeIMO-OneNote-TocLink-" + Guid.NewGuid().ToString("N"));
        string outside = Path.Combine(Path.GetTempPath(), "OfficeIMO-OneNote-OutsideToc-" + Guid.NewGuid().ToString("N"));
        try {
            OneNoteNotebookWriter.Write(CreateNotebook(), root);
            var outsideNotebook = new OneNoteNotebook { Name = "Private group" };
            outsideNotebook.Sections.Add(CreateSection("Private nested", "External nested content"));
            OneNoteNotebookWriter.Write(outsideNotebook, outside);

            string nestedToc = Path.Combine(root, "Group", "Open Notebook.onetoc2");
            File.Delete(nestedToc);
            File.CreateSymbolicLink(nestedToc, Path.Combine(outside, "Open Notebook.onetoc2"));

            OneNoteNotebook result = OneNoteNotebookReader.Read(Path.Combine(root, "Open Notebook.onetoc2"));

            Assert.Single(result.Sections);
            Assert.Empty(Assert.Single(result.SectionGroups).Sections);
            Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "ONENOTE_TOC_GROUP_MISSING");
        } finally {
            if (Directory.Exists(root)) Directory.Delete(root, true);
            if (Directory.Exists(outside)) Directory.Delete(outside, true);
        }
    }
#endif

    [Fact]
    public void PackageReadWritePreservesOpaqueRootEntryAndNestedTocObjects() {
        OneNoteNotebook loaded = OneNotePackageReader.Read(
            new MemoryStream(OneNotePackageWriter.Write(CreateNotebook())),
            "opaque.onepkg");
        OneNoteSectionGroup loadedGroup = Assert.Single(loaded.SectionGroups);
        AddOpaqueScalar(FindRoot(loaded.UnknownObjects, loaded.TableOfContentsRootObjectId), 0x1400ABCD, 101);
        AddOpaqueScalar(FindEntry(loaded.UnknownObjects, Assert.Single(loaded.Sections).Id!.Value), 0x1400ABCE, 202);
        AddOpaqueScalar(FindRoot(loadedGroup.UnknownObjects, loadedGroup.TableOfContentsRootObjectId), 0x1400ABCF, 303);
        OneNoteExtendedGuid extraId = new OneNoteExtendedGuid(Guid.NewGuid(), 1, 17);
        var extra = new OneNoteOpaqueObject {
            Id = extraId,
            Jcid = OneNoteSchema.JcidPropertyContainer,
            Ordinal = loaded.UnknownObjects.Count
        };
        AddOpaqueScalar(extra, 0x1400ABD0, 404);
        loaded.UnknownObjects.Add(extra);

        OneNoteNotebook roundTrip = OneNotePackageReader.Read(
            new MemoryStream(OneNotePackageWriter.Write(loaded)),
            "opaque-roundtrip.onepkg");

        AssertOpaqueScalar(FindRoot(roundTrip.UnknownObjects, roundTrip.TableOfContentsRootObjectId), 0x1400ABCD, 101);
        AssertOpaqueScalar(FindEntry(roundTrip.UnknownObjects, Assert.Single(roundTrip.Sections).Id!.Value), 0x1400ABCE, 202);
        OneNoteSectionGroup roundTripGroup = Assert.Single(roundTrip.SectionGroups);
        AssertOpaqueScalar(FindRoot(roundTripGroup.UnknownObjects, roundTripGroup.TableOfContentsRootObjectId), 0x1400ABCF, 303);
        AssertOpaqueScalar(Assert.Single(roundTrip.UnknownObjects, item => extraId.Equals(item.Id)), 0x1400ABD0, 404);
    }

    private static OneNoteNotebook CreateNotebook() {
        var notebook = new OneNoteNotebook {
            Id = new Guid("9f84c4d1-a8f6-4fdb-8cb9-7c5d7bb6e2a1"),
            Name = "Writer",
            ColorArgb = 0xFF123456U,
            HistoryEnabled = true
        };
        notebook.Sections.Add(CreateSection("Root", "Root page"));
        var group = new OneNoteSectionGroup { Name = "Group" };
        group.Sections.Add(CreateSection("Nested", "Nested page"));
        notebook.SectionGroups.Add(group);
        return notebook;
    }

    private static OneNoteNotebook CreateGroupChain(int groupCount) {
        var notebook = new OneNoteNotebook { Name = "Groups" };
        var root = new OneNoteSectionGroup { Name = "Group 0" };
        OneNoteSectionGroup parent = root;
        for (int index = 1; index < groupCount; index++) {
            var child = new OneNoteSectionGroup { Name = "Group " + index };
            parent.SectionGroups.Add(child);
            parent = child;
        }
        notebook.SectionGroups.Add(root);
        return notebook;
    }

    private static OneNoteSection CreateSection(string name, string title) {
        var section = new OneNoteSection { Name = name };
        var page = new OneNotePage { Title = title };
        var paragraph = new OneNoteParagraph();
        paragraph.Runs.Add(new OneNoteTextRun { Text = title + " content" });
        page.DirectContent.Add(paragraph);
        section.Pages.Add(page);
        return section;
    }

    private static OneNoteOpaqueObject FindRoot(
        IEnumerable<OneNoteOpaqueObject> objects,
        OneNoteExtendedGuid? rootId) => Assert.Single(objects, item => rootId != null && rootId.Equals(item.Id));

    private static OneNoteOpaqueObject FindEntry(IEnumerable<OneNoteOpaqueObject> objects, Guid fileId) =>
        Assert.Single(objects, item => item.Properties.Any(property => {
            if ((property.PropertyId & 0x03FFFFFFU) != (OneNoteSchema.FileIdentityGuid & 0x03FFFFFFU)) return false;
            byte[] data = property.GetRawData();
            return data.Length == 16 && new Guid(data) == fileId;
        }));

    private static void AddOpaqueScalar(OneNoteOpaqueObject target, uint propertyId, ulong value) {
        target.Properties.Add(new OneNoteOpaqueProperty {
            PropertyId = propertyId,
            ValueType = OneNotePropertyValueType.UInt32,
            Ordinal = target.Properties.Count,
            ScalarValue = value
        });
    }

    private static void AssertOpaqueScalar(OneNoteOpaqueObject target, uint propertyId, ulong value) {
        OneNoteOpaqueProperty property = Assert.Single(
            target.Properties,
            item => (item.PropertyId & 0x03FFFFFFU) == (propertyId & 0x03FFFFFFU));
        Assert.Equal(value, property.ScalarValue);
    }
}

using System.Security.Cryptography;
using System.Text;

namespace OfficeIMO.OneNote.Tests;

public sealed class SectionWriterTests {
    [Fact]
    public void CreatesAndReadsDesktopSectionWithHierarchyAndRichText() {
        OneNoteSection original = CreateSection();

        byte[] data = OneNoteSectionWriter.Write(original);
        using var stream = new MemoryStream(data);
        OneNoteFileHeader header = OneNoteFileProbe.ReadHeader(stream);
        OneNoteSection roundTrip = OneNoteSectionReader.Read(stream);

        Assert.Equal(OneNoteStorageFormat.RevisionStore, header.StorageFormat);
        Assert.Equal(OneNoteFileKind.Section, header.FileKind);
        Assert.Equal("Writer sample", roundTrip.Name);
        Assert.Equal(0xFF336699U, roundTrip.ColorArgb);
        Assert.Collection(roundTrip.Pages,
            page => {
                Assert.Equal("Parent page", page.Title);
                Assert.Equal(0, page.Level);
                OneNoteParagraph paragraph = Assert.Single(Assert.Single(page.Outlines).Children.OfType<OneNoteParagraph>());
                Assert.Collection(paragraph.Runs,
                    run => {
                        Assert.Equal("Hello ", run.Text);
                        Assert.True(run.Style.Bold);
                        Assert.Equal("Aptos", run.Style.FontFamily);
                        Assert.Equal(12, run.Style.FontSize);
                        Assert.Equal(0x0415U, run.Style.LanguageId);
                    },
                    run => {
                        Assert.Equal("world", run.Text);
                        Assert.True(run.Style.Italic);
                        Assert.Equal("https://example.test/", run.Hyperlink);
                    });
            },
            page => {
                Assert.Equal("Child page", page.Title);
                Assert.Equal(1, page.Level);
            });
    }

    [Theory]
    [InlineData(OneNoteStorageFormat.RevisionStore)]
    [InlineData(OneNoteStorageFormat.FileSynchronizationPackage)]
    public void MaximumPageHierarchyLevelRoundTripsWithoutArithmeticOverflow(OneNoteStorageFormat storageFormat) {
        var section = new OneNoteSection { Name = "Hierarchy" };
        section.Pages.Add(new OneNotePage { Title = "Extreme level", Level = int.MaxValue });

        byte[] data = OneNoteSectionWriter.Write(section, new OneNoteWriterOptions {
            StorageFormat = storageFormat
        });
        OneNoteSection result = OneNoteSectionReader.Read(new MemoryStream(data));

        Assert.Equal(int.MaxValue, Assert.Single(result.Pages).Level);
    }

    [Fact]
    public void EmitsCommittedDesktopPhysicalStructures() {
        byte[] data = OneNoteSectionWriter.Write(CreateSection());

        OneNoteRevisionStore store = OneNoteRevisionStoreReader.Read(new MemoryStream(data));

        Assert.Equal((ulong)data.Length, store.Header.ExpectedFileLength);
        Assert.Equal(1U, store.Header.TransactionCount);
        Assert.NotNull(store.Header.RootFileNodeList);
        Assert.NotNull(store.Header.TransactionLog);
        Assert.Contains(store.FileNodeLists, list => list.Nodes.Any(node => node.Id == OneNoteFileNodeId.ObjectGroupStart));
        Assert.Contains(store.FileNodeLists, list => list.Nodes.Any(node => node.Id == OneNoteFileNodeId.ObjectInfoDependencyOverrides));
        Assert.All(
            store.Objects.Where(item => item.Jcid.IsReadOnly),
            item => {
                Assert.Equal(OneNoteFileNodeId.ReadOnlyObjectDeclaration2RefCount, item.DeclarationNode.Id);
                byte[] propertyData = Assert.IsType<OneNoteBinaryPayload>(item.RawPropertyData).ToArray(1024 * 1024);
                byte[] expectedHash;
                using (MD5 md5 = MD5.Create()) expectedHash = md5.ComputeHash(propertyData);
                byte[] encodedDeclaration = item.DeclarationNode.EncodedData.ToArray(1024 * 1024);
                Assert.Equal(expectedHash, encodedDeclaration.Skip(encodedDeclaration.Length - expectedHash.Length));
            });
        Assert.All(
            store.Objects.Where(item => !item.Jcid.IsReadOnly && item.FileDataReference == null),
            item => Assert.Equal(OneNoteFileNodeId.ObjectDeclaration2RefCount, item.DeclarationNode.Id));
    }

    [Fact]
    public void EmitsRequiredSectionIdentityAndCachedPageMetadataGraph() {
        OneNoteWriteGraph graph = new OneNoteWriteGraphBuilder().BuildSection(CreateSection());
        OneNoteWriteObjectSpace sectionSpace = graph.ObjectSpaces[0];

        OneNoteWriteObject sectionNode = Assert.Single(sectionSpace.Objects, item => item.Jcid == OneNoteSchema.JcidSectionNode);
        AssertGuidProperty(sectionNode, OneNoteSchema.NotebookManagementEntityGuid);
        Assert.NotNull(Property(sectionNode, OneNoteSchema.TopologyCreationTimestamp).Scalar);

        OneNoteWriteObject sectionMetadata = Assert.Single(sectionSpace.Objects, item => item.Jcid == OneNoteSchema.JcidSectionMetadata);
        Assert.Equal(40UL, Property(sectionMetadata, OneNoteSchema.SchemaRevisionInOrderToRead).Scalar);
        Assert.Equal(40UL, Property(sectionMetadata, OneNoteSchema.SchemaRevisionInOrderToWrite).Scalar);
        Assert.NotNull(Property(sectionMetadata, OneNoteSchema.NotebookColor).Scalar);

        OneNoteWriteObject[] series = sectionSpace.Objects.Where(item => item.Jcid == OneNoteSchema.JcidPageSeriesNode).ToArray();
        OneNoteWriteObject[] cachedMetadata = sectionSpace.Objects.Where(item => item.Jcid == OneNoteSchema.JcidPageMetadata).ToArray();
        Assert.Equal(graph.ObjectSpaces.Count - 1, series.Length);
        Assert.Equal(series.Length, cachedMetadata.Length);

        for (int index = 0; index < series.Length; index++) {
            AssertGuidProperty(series[index], OneNoteSchema.NotebookManagementEntityGuid);
            Assert.NotNull(Property(series[index], OneNoteSchema.TopologyCreationTimestamp).Scalar);
            Assert.Equal(cachedMetadata[index].Id, Assert.Single(Property(series[index], OneNoteSchema.MetaDataObjectsAboveGraphSpace).References));

            OneNoteWriteObject pageMetadata = Assert.Single(graph.ObjectSpaces[index + 1].Objects, item => item.Jcid == OneNoteSchema.JcidPageMetadata);
            Assert.Equal(
                Property(cachedMetadata[index], OneNoteSchema.NotebookManagementEntityGuid).Data,
                Property(pageMetadata, OneNoteSchema.NotebookManagementEntityGuid).Data);
            Assert.Equal(40UL, Property(pageMetadata, OneNoteSchema.SchemaRevisionInOrderToRead).Scalar);
            Assert.Equal(40UL, Property(pageMetadata, OneNoteSchema.SchemaRevisionInOrderToWrite).Scalar);

            OneNoteWriteObject revisionMetadata = Assert.Single(graph.ObjectSpaces[index + 1].Objects, item => item.Jcid == OneNoteSchema.JcidRevisionMetadata);
            Assert.NotNull(Property(revisionMetadata, OneNoteSchema.LastModifiedTimestamp).Scalar);
            OneNoteWriteObjectSpace pageSpace = graph.ObjectSpaces[index + 1];
            OneNoteExtendedGuid[] runStyleIds = pageSpace.Objects
                .Where(item => item.Jcid == OneNoteSchema.JcidRichTextNode)
                .SelectMany(item => item.Properties.Where(property => (property.RawId & 0x7FFFFFFFU) == OneNoteSchema.TextRunFormatting))
                .SelectMany(property => property.References)
                .Distinct()
                .ToArray();
            Assert.All(
                pageSpace.Objects.Where(item => runStyleIds.Contains(item.Id)),
                item => Assert.NotNull(Property(item, OneNoteSchema.LanguageId).Scalar));
        }
    }

    [Fact]
    public void EmitsNativePageTitleStructureSeparatelyFromBodyContent() {
        OneNoteWriteObjectSpace pageSpace = new OneNoteWriteGraphBuilder().BuildSection(CreateSection()).ObjectSpaces[1];
        OneNoteWriteObject pageNode = Assert.Single(pageSpace.Objects, item => item.Jcid == OneNoteSchema.JcidPageNode);
        OneNoteExtendedGuid titleId = Assert.Single(Property(pageNode, OneNoteSchema.StructureElementChildNodes).References);
        OneNoteWriteObject title = Assert.Single(pageSpace.Objects, item => item.Id == titleId);

        Assert.Equal(OneNoteSchema.JcidTitleNode, title.Jcid);
        OneNoteExtendedGuid titleOutlineId = Assert.Single(Property(title, OneNoteSchema.ElementChildNodes).References);
        OneNoteWriteObject titleOutline = Assert.Single(pageSpace.Objects, item => item.Id == titleOutlineId);
        Assert.True(IsTrue(Property(titleOutline, OneNoteSchema.EnforceOutlineStructure)));
        Assert.True(IsTrue(Property(titleOutline, OneNoteSchema.IsTitleText)));

        OneNoteExtendedGuid titleElementId = Assert.Single(Property(titleOutline, OneNoteSchema.ElementChildNodes).References);
        OneNoteWriteObject titleElement = Assert.Single(pageSpace.Objects, item => item.Id == titleElementId);
        Assert.True(IsTrue(Property(titleElement, OneNoteSchema.CannotBeSelected)));
        Assert.True(IsTrue(Property(titleElement, OneNoteSchema.IsTitleText)));
        Assert.NotNull(Property(titleElement, OneNoteSchema.CreationTimestamp).Scalar);

        OneNoteExtendedGuid titleTextId = Assert.Single(Property(titleElement, OneNoteSchema.ContentChildNodes).References);
        OneNoteWriteObject titleText = Assert.Single(pageSpace.Objects, item => item.Id == titleTextId);
        Assert.Equal(OneNoteSchema.JcidRichTextNode, titleText.Jcid);
        Assert.True(IsTrue(Property(titleText, OneNoteSchema.IsTitleText)));
        OneNoteWriteProperty encodedTitle = titleText.Properties.Single(property =>
            (property.RawId & 0x7FFFFFFFU) == OneNoteSchema.RichEditTextUnicode ||
            (property.RawId & 0x7FFFFFFFU) == OneNoteSchema.TextExtendedAscii);
        string titleValue = (encodedTitle.RawId & 0x7FFFFFFFU) == OneNoteSchema.TextExtendedAscii
            ? Encoding.ASCII.GetString(encodedTitle.Data!)
            : Encoding.Unicode.GetString(encodedTitle.Data!).TrimEnd('\0');
        Assert.Equal("Parent page", titleValue);

        OneNoteExtendedGuid[] bodyIds = Property(pageNode, OneNoteSchema.ElementChildNodes).References.ToArray();
        Assert.DoesNotContain(titleOutlineId, bodyIds);
    }

    [Fact]
    public void NormalizesDirectPageContentIntoRenderableOutline() {
        var section = new OneNoteSection { Name = "Direct content" };
        var page = new OneNotePage { Title = "Direct page" };
        var paragraph = new OneNoteParagraph();
        paragraph.Runs.Add(new OneNoteTextRun { Text = "Outside an outline" });
        page.DirectContent.Add(paragraph);
        page.DirectContent.Add(new OneNoteEmbeddedFile {
            FileName = "direct.bin",
            Payload = OneNoteBinaryPayload.FromBytes(new byte[] { 1, 2, 3 })
        });
        section.Pages.Add(page);

        byte[] data = OneNoteSectionWriter.Write(section);
        Assert.Empty(page.DirectContent);
        OneNoteOutline normalizedOutline = Assert.Single(page.Outlines);
        Assert.Same(paragraph, normalizedOutline.Children[0]);
        Assert.NotNull(normalizedOutline.Id);
        OneNoteWriteObjectSpace pageSpace = new OneNoteWriteGraphBuilder().BuildSection(section).ObjectSpaces[1];
        OneNoteWriteObject pageNode = Assert.Single(pageSpace.Objects, item => item.Jcid == OneNoteSchema.JcidPageNode);
        OneNoteExtendedGuid outlineId = Assert.Single(Property(pageNode, OneNoteSchema.ElementChildNodes).References);
        OneNoteWriteObject outline = Assert.Single(pageSpace.Objects, item => item.Id == outlineId);
        Assert.Equal(OneNoteSchema.JcidOutlineNode, outline.Jcid);
        Assert.Equal(2, Property(outline, OneNoteSchema.ElementChildNodes).References.Count);
        Assert.Equal(1.0F, FloatValue(Property(outline, OneNoteSchema.OffsetFromParentHorizontal)));
        Assert.Equal(2.4F, FloatValue(Property(outline, OneNoteSchema.OffsetFromParentVertical)));
        Assert.Equal(normalizedOutline.Id, outline.Id);

        OneNoteSection roundTrip = OneNoteSectionReader.Read(new MemoryStream(data));

        OneNotePage result = Assert.Single(roundTrip.Pages);
        Assert.Empty(result.DirectContent);
        Assert.Collection(Assert.Single(result.Outlines).Children,
            element => Assert.Equal("Outside an outline", Assert.Single(Assert.IsType<OneNoteParagraph>(element).Runs).Text),
            element => {
                OneNoteEmbeddedFile file = Assert.IsType<OneNoteEmbeddedFile>(element);
                Assert.Equal("direct.bin", file.FileName);
                Assert.Equal(new byte[] { 1, 2, 3 }, file.Payload!.ToArray(16));
            });
    }

    [Fact]
    public void PlacesRootManifestAfterFirstObjectSpaceManifest() {
        OneNoteRevisionStore store = OneNoteRevisionStoreReader.Read(new MemoryStream(OneNoteSectionWriter.Write(CreateSection())));
        OneNoteFileNodeId[] declarations = store.RootFileNodeList.Nodes
            .Where(node => node.Id == OneNoteFileNodeId.ObjectSpaceManifestListReference || node.Id == OneNoteFileNodeId.ObjectSpaceManifestRoot)
            .Select(node => node.Id)
            .ToArray();

        Assert.Equal(OneNoteFileNodeId.ObjectSpaceManifestListReference, declarations[0]);
        Assert.Equal(OneNoteFileNodeId.ObjectSpaceManifestRoot, declarations[1]);
        Assert.Equal(1, declarations.Count(id => id == OneNoteFileNodeId.ObjectSpaceManifestRoot));
    }

    [Fact]
    public void EmitsAndValidatesTransactionChecksum() {
        byte[] data = OneNoteSectionWriter.Write(CreateSection());
        OneNoteFileHeader header = OneNoteFileProbe.ReadHeader(new MemoryStream(data));
        OneNoteFileChunkReference transaction = Assert.IsType<OneNoteFileChunkReference>(header.TransactionLog);
        int offset = checked((int)transaction.Offset);
        int sentinelOffset = offset + (ReadCommittedListCount(data, offset) * 8);

        uint expected = OneNoteCrc32.Compute(data.Skip(offset).Take(sentinelOffset - offset).ToArray());
        Assert.Equal(1U, BitConverter.ToUInt32(data, sentinelOffset));
        Assert.Equal(expected, BitConverter.ToUInt32(data, sentinelOffset + 4));

        data[sentinelOffset + 4] ^= 0x01;
        OneNoteFormatException exception = Assert.Throws<OneNoteFormatException>(() =>
            OneNoteRevisionStoreReader.Read(new MemoryStream(data)));
        Assert.Equal("ONENOTE_TRANSACTION_CHECKSUM", exception.Code);
    }

    [Fact]
    public void DestinationStreamRemainsOpen() {
        using var destination = new MemoryStream();

        OneNoteSectionWriter.Write(CreateSection(), destination);
        destination.WriteByte(0);

        Assert.True(destination.Length > 1);
    }

    [Fact]
    public void RejectsUnsupportedTypedContentInsteadOfDroppingIt() {
        var section = new OneNoteSection { Name = "Unsupported" };
        var page = new OneNotePage { Title = "Ink" };
        page.DirectContent.Add(new OneNoteInk());
        section.Pages.Add(page);

        OneNoteFormatException exception = Assert.Throws<OneNoteFormatException>(() => OneNoteSectionWriter.Write(section));

        Assert.Equal("ONENOTE_WRITE_UNSUPPORTED_INK", exception.Code);
    }

    [Fact]
    public void RejectsRawMathInsteadOfFlatteningItToPlainText() {
        var section = new OneNoteSection { Name = "Unsupported" };
        var page = new OneNotePage { Title = "Math" };
        page.DirectContent.Add(new OneNoteMath { Text = "x+y", MathMl = "<math><mi>x</mi></math>" });
        section.Pages.Add(page);

        OneNoteFormatException exception = Assert.Throws<OneNoteFormatException>(() => OneNoteSectionWriter.Write(section));

        Assert.Equal("ONENOTE_WRITE_UNSUPPORTED_MATH", exception.Code);
    }

    [Fact]
    public void RoundTripsTablesListsParagraphStylesAssetsAndMath() {
        var section = new OneNoteSection { Name = "Content" };
        var page = new OneNotePage { Title = "Kinds", MostRecentAuthor = "Recent author" };
        var outline = new OneNoteOutline();

        var listed = new OneNoteParagraph { List = new OneNoteListInfo { Ordered = true, Format = 3, Level = 2, Restart = true, DisplayIndex = 4, FontFamily = "Aptos" } };
        listed.Style.Alignment = OneNoteParagraphAlignment.Center;
        listed.Style.SpaceAfter = 6;
        listed.Author = new OneNoteAuthor { Name = "Paragraph author" };
        listed.Runs.Add(new OneNoteTextRun { Text = "List item" });
        outline.Children.Add(listed);

        var table = new OneNoteTable { BordersVisible = true };
        table.ColumnWidths.Add(120);
        var row = new OneNoteTableRow();
        var cell = new OneNoteTableCell { ShadingColorArgb = 0xFF112233 };
        var cellText = new OneNoteParagraph();
        cellText.Runs.Add(new OneNoteTextRun { Text = "Cell" });
        cell.Content.Add(cellText);
        row.Cells.Add(cell);
        table.Rows.Add(row);
        outline.Children.Add(table);

        outline.Children.Add(new OneNoteImage {
            FileName = "pixel.png",
            AltText = "pixel",
            WidthHalfInches = 2.5,
            HeightHalfInches = 1.25,
            Payload = OneNoteBinaryPayload.FromBytes(new byte[] { 1, 2, 3, 4 })
        });
        outline.Children.Add(new OneNoteEmbeddedFile {
            FileName = "notes.txt",
            SourcePath = "C:\\source\\notes.txt",
            Payload = OneNoteBinaryPayload.FromBytes(new byte[] { 5, 6, 7 })
        });
        outline.Children.Add(new OneNoteMath { Text = "x+y" });
        page.Outlines.Add(outline);
        section.Pages.Add(page);

        OneNoteWriteObjectSpace pageSpace = new OneNoteWriteGraphBuilder().BuildSection(section).ObjectSpaces[1];
        var timeBearingTypes = new HashSet<uint> {
            OneNoteSchema.JcidPageNode,
            OneNoteSchema.JcidOutlineNode,
            OneNoteSchema.JcidOutlineElementNode,
            OneNoteSchema.JcidRichTextNode,
            OneNoteSchema.JcidNumberListNode,
            OneNoteSchema.JcidTableNode,
            OneNoteSchema.JcidTableRowNode,
            OneNoteSchema.JcidTableCellNode,
            OneNoteSchema.JcidImageNode,
            OneNoteSchema.JcidEmbeddedFileNode
        };
        OneNoteWriteObject[] timeBearingObjects = pageSpace.Objects.Where(item => timeBearingTypes.Contains(item.Jcid)).ToArray();
        Assert.NotEmpty(timeBearingObjects);
        Assert.All(timeBearingObjects, item => Assert.NotNull(Property(item, OneNoteSchema.LastModifiedTime).Scalar));

        OneNoteSection roundTrip = OneNoteSectionReader.Read(new MemoryStream(OneNoteSectionWriter.Write(section)));
        OneNoteOutline result = Assert.Single(Assert.Single(roundTrip.Pages).Outlines);
        OneNoteParagraph resultList = Assert.IsType<OneNoteParagraph>(result.Children[0]);
        Assert.True(resultList.List!.Ordered);
        Assert.Equal(3U, resultList.List.Format);
        Assert.Equal(2, resultList.List.Level);
        Assert.Equal(4, resultList.List.DisplayIndex);
        Assert.Equal(OneNoteParagraphAlignment.Center, resultList.Style.Alignment);
        Assert.Equal(6, resultList.Style.SpaceAfter);
        Assert.Equal("Paragraph author", resultList.Author!.Name);
        Assert.Equal("Recent author", roundTrip.Pages[0].MostRecentAuthor);

        OneNoteTable resultTable = Assert.IsType<OneNoteTable>(result.Children[1]);
        Assert.True(resultTable.BordersVisible);
        Assert.Equal(120, Assert.Single(resultTable.ColumnWidths));
        Assert.Equal(0xFF112233U, Assert.Single(Assert.Single(resultTable.Rows).Cells).ShadingColorArgb);
        OneNoteImage resultImage = Assert.IsType<OneNoteImage>(result.Children[2]);
        Assert.Equal(new byte[] { 1, 2, 3, 4 }, resultImage.Payload!.ToArray(16));
        Assert.Equal(2.5, resultImage.WidthHalfInches);
        Assert.Equal(1.25, resultImage.HeightHalfInches);
        OneNoteEmbeddedFile resultFile = Assert.IsType<OneNoteEmbeddedFile>(result.Children[3]);
        Assert.Equal(new byte[] { 5, 6, 7 }, resultFile.Payload!.ToArray(16));
        OneNoteMath resultMath = Assert.IsType<OneNoteMath>(result.Children[4]);
        Assert.Equal("x+y", resultMath.Text);
    }

    [Fact]
    public void RoundTripsNormalAndTaskTagsThroughNestedPropertySets() {
        DateTime created = new DateTime(2026, 7, 16, 8, 0, 0, DateTimeKind.Utc);
        DateTime completed = created.AddHours(1);
        var section = new OneNoteSection { Name = "Tags" };
        var page = new OneNotePage { Title = "Tagged" };
        var outline = new OneNoteOutline();
        var paragraph = new OneNoteParagraph();
        paragraph.Runs.Add(new OneNoteTextRun { Text = "Tagged text" });
        paragraph.Tags.Add(new OneNoteTag {
            ActionItemType = 8,
            Shape = 1,
            Label = "Important",
            IsCheckable = true,
            CreatedUtc = created,
            TextColorArgb = 0x000000FF,
            HighlightColorArgb = 0x0000FFFF
        });
        paragraph.Tags.Add(new OneNoteTag {
            ActionItemType = 105,
            Shape = 89,
            IsTask = true,
            IsCheckable = true,
            IsCompleted = true,
            IsUnsynchronized = true,
            DueUtc = created.AddDays(1),
            CreatedUtc = created,
            CompletedUtc = completed
        });
        outline.Children.Add(paragraph);
        page.Outlines.Add(outline);
        section.Pages.Add(page);

        byte[] data = OneNoteSectionWriter.Write(section);
        OneNoteRevisionStore store = OneNoteRevisionStoreReader.Read(new MemoryStream(data));
        OneNoteRevisionStoreObject definition = Assert.Single(store.Objects, item => item.Jcid.Value == OneNoteSchema.JcidNoteTagSharedDefinition);
        Assert.Equal(1U, definition.ReferenceCount);
        OneNoteSection roundTrip = OneNoteSectionReader.Read(new MemoryStream(data));
        OneNoteParagraph result = Assert.IsType<OneNoteParagraph>(Assert.Single(Assert.Single(roundTrip.Pages).Outlines).Children[0]);
        Assert.Collection(result.Tags,
            normal => {
                Assert.False(normal.IsTask);
                Assert.Equal(8U, normal.ActionItemType);
                Assert.Equal(1U, normal.Shape);
                Assert.Equal("Important", normal.Label);
                Assert.Equal(0x000000FFU, normal.TextColorArgb);
                Assert.Equal(0x0000FFFFU, normal.HighlightColorArgb);
                Assert.Equal(created, normal.CreatedUtc);
            },
            task => {
                Assert.True(task.IsTask);
                Assert.Equal(105U, task.ActionItemType);
                Assert.Equal(89U, task.Shape);
                Assert.True(task.IsCompleted);
                Assert.True(task.IsUnsynchronized);
                Assert.Equal(created.AddDays(1), task.DueUtc);
                Assert.Equal(created, task.CreatedUtc);
                Assert.Equal(completed, task.CompletedUtc);
            });
    }

    [Fact]
    public void NormalTagDefaultShapeHonorsNonCheckableState() {
        DateTime created = new DateTime(2026, 7, 16, 9, 0, 0, DateTimeKind.Utc);
        var section = new OneNoteSection { Name = "Tags" };
        var page = new OneNotePage { Title = "Tagged" };
        var paragraph = new OneNoteParagraph();
        paragraph.Runs.Add(new OneNoteTextRun { Text = "Reference" });
        paragraph.Tags.Add(new OneNoteTag {
            ActionItemType = 8,
            Label = "Reference",
            IsCheckable = false,
            CreatedUtc = created
        });
        page.DirectContent.Add(paragraph);
        section.Pages.Add(page);

        OneNoteSection roundTrip = OneNoteSectionReader.Read(new MemoryStream(OneNoteSectionWriter.Write(section)));

        OneNotePage resultPage = Assert.Single(roundTrip.Pages);
        Assert.Empty(resultPage.DirectContent);
        OneNoteParagraph resultParagraph = Assert.IsType<OneNoteParagraph>(Assert.Single(Assert.Single(resultPage.Outlines).Children));
        OneNoteTag result = Assert.Single(resultParagraph.Tags);
        Assert.False(result.IsCheckable);
        Assert.Equal(13U, result.Shape);
        Assert.True(result.IsCompleted);
        Assert.Equal(created, result.CompletedUtc);
    }

    [Fact]
    public void RejectsImpossibleNonCheckableTaskTag() {
        var section = new OneNoteSection { Name = "Tasks" };
        var page = new OneNotePage { Title = "Task" };
        var paragraph = new OneNoteParagraph();
        paragraph.Tags.Add(new OneNoteTag { IsTask = true, IsCheckable = false });
        page.DirectContent.Add(paragraph);
        section.Pages.Add(page);

        OneNoteFormatException exception = Assert.Throws<OneNoteFormatException>(() => OneNoteSectionWriter.Write(section));

        Assert.Equal("ONENOTE_WRITE_TASK_TAG_CHECKABILITY", exception.Code);
    }

    [Fact]
    public void PageDeletionMarkerCanBeSetAndClearedAcrossPreservedWrites() {
        var section = new OneNoteSection { Name = "Deleted pages" };
        section.Pages.Add(new OneNotePage { Title = "Recoverable", IsDeleted = true });

        OneNoteSection deleted = OneNoteSectionReader.Read(new MemoryStream(OneNoteSectionWriter.Write(section)));
        Assert.True(Assert.Single(deleted.Pages).IsDeleted);

        deleted.Pages[0].IsDeleted = false;
        OneNoteSection restored = OneNoteSectionReader.Read(new MemoryStream(OneNoteSectionWriter.Write(deleted)));

        Assert.False(Assert.Single(restored.Pages).IsDeleted);
    }

    [Fact]
    public void RoundTripsConflictPageObjectSpacesWithoutPromotingThemToTopLevelPages() {
        var section = new OneNoteSection { Name = "Conflicts" };
        var page = new OneNotePage { Title = "Current" };
        var conflict = new OneNotePage { Title = "Conflict copy", IsConflictPage = true };
        var conflictParagraph = new OneNoteParagraph();
        conflictParagraph.Runs.Add(new OneNoteTextRun { Text = "Conflicting content" });
        conflict.DirectContent.Add(conflictParagraph);
        page.ConflictPages.Add(conflict);
        section.Pages.Add(page);

        byte[] data = OneNoteSectionWriter.Write(section);
        OneNoteRevisionStore store = OneNoteRevisionStoreReader.Read(new MemoryStream(data));
        var materializer = new OneNoteObjectSpaceMaterializer(store);
        OneNoteMaterializedObjectSpace sectionSpace = materializer.FindCurrentSpaceByRootJcid(
            OneNoteSchema.JcidSectionNode,
            "TEST_SECTION",
            "The generated section object space was not materialized.");
        OneNoteRevisionStoreObject sectionRoot = Assert.IsType<OneNoteRevisionStoreObject>(sectionSpace.GetRoot(1));
        OneNoteExtendedGuid seriesId = Assert.Single(OneNoteSemanticMapper.GetReferences(sectionRoot, OneNoteSchema.ElementChildNodes));
        OneNoteRevisionStoreObject series = Assert.IsType<OneNoteRevisionStoreObject>(sectionSpace.GetObject(seriesId));
        OneNoteExtendedGuid currentSpaceId = Assert.Single(OneNoteSemanticMapper.GetReferences(series, OneNoteSchema.ChildGraphSpaceElementNodes));
        OneNoteMaterializedObjectSpace currentSpace = Assert.IsType<OneNoteMaterializedObjectSpace>(materializer.TryGetCurrentSpace(currentSpaceId));
        OneNoteRevisionStoreObject currentManifest = Assert.IsType<OneNoteRevisionStoreObject>(currentSpace.GetRoot(1));
        OneNoteExtendedGuid conflictSpaceId = Assert.Single(OneNoteSemanticMapper.GetReferences(currentManifest, OneNoteSchema.ChildGraphSpaceElementNodes));
        OneNoteMaterializedObjectSpace conflictSpace = Assert.IsType<OneNoteMaterializedObjectSpace>(materializer.TryGetCurrentSpace(conflictSpaceId));
        Assert.Equal(OneNoteSchema.JcidPageManifestNode, Assert.IsType<OneNoteRevisionStoreObject>(conflictSpace.GetRoot(1)).Jcid.Value);
        Assert.Equal(OneNoteSchema.JcidConflictPageMetadata, Assert.IsType<OneNoteRevisionStoreObject>(conflictSpace.GetRoot(2)).Jcid.Value);

        OneNoteSection roundTrip = OneNoteSectionReader.Read(new MemoryStream(data));

        OneNotePage current = Assert.Single(roundTrip.Pages);
        Assert.Equal("Current", current.Title);
        OneNotePage conflictResult = Assert.Single(current.ConflictPages);
        Assert.True(conflictResult.IsConflictPage);
        Assert.Equal("Conflict copy", conflictResult.Title);
        Assert.Empty(conflictResult.DirectContent);
        OneNoteParagraph conflictResultParagraph = Assert.IsType<OneNoteParagraph>(Assert.Single(Assert.Single(conflictResult.Outlines).Children));
        Assert.Equal("Conflicting content", string.Concat(conflictResultParagraph.Runs.Select(run => run.Text)));
    }

    [Fact]
    public void RoundTripsVersionHistoryThroughDesktopRevisionContexts() {
        OneNoteSection section = CreateVersionedSection();

        byte[] data = OneNoteSectionWriter.Write(section);

        AssertVersionHistoryRoundTrip(data);
    }

    [Fact]
    public void RoundTripsVersionHistoryThroughFssHttpCells() {
        OneNoteSection section = CreateVersionedSection();

        byte[] data = OneNoteSectionWriter.Write(section, new OneNoteWriterOptions {
            StorageFormat = OneNoteStorageFormat.FileSynchronizationPackage
        });

        Assert.Equal(OneNoteStorageFormat.FileSynchronizationPackage, OneNoteFileProbe.ReadHeader(new MemoryStream(data)).StorageFormat);
        AssertVersionHistoryRoundTrip(data);
    }

    [Theory]
    [InlineData(OneNoteStorageFormat.RevisionStore)]
    [InlineData(OneNoteStorageFormat.FileSynchronizationPackage)]
    public void RoundTripsNestedConflictAndVersionPagesWithAssets(OneNoteStorageFormat storageFormat) {
        var section = new OneNoteSection { Name = "Nested related pages" };
        var current = new OneNotePage { Title = "Current" };
        var conflict = new OneNotePage { Title = "Conflict", IsConflictPage = true };
        var conflictVersion = new OneNotePage { Title = "Conflict version", IsVersionHistoryPage = true };
        conflictVersion.DirectContent.Add(new OneNoteImage {
            FileName = "historical.png",
            MediaType = "image/png",
            Payload = OneNoteBinaryPayload.FromBytes(new byte[] { 1, 2, 3 })
        });
        conflict.VersionHistory.Add(conflictVersion);
        current.ConflictPages.Add(conflict);

        var version = new OneNotePage { Title = "Version", IsVersionHistoryPage = true };
        var versionConflict = new OneNotePage { Title = "Version conflict", IsConflictPage = true };
        versionConflict.DirectContent.Add(new OneNoteEmbeddedFile {
            FileName = "evidence.bin",
            Payload = OneNoteBinaryPayload.FromBytes(new byte[] { 4, 5, 6 })
        });
        version.ConflictPages.Add(versionConflict);
        current.VersionHistory.Add(version);
        section.Pages.Add(current);

        byte[] data = OneNoteSectionWriter.Write(section, new OneNoteWriterOptions { StorageFormat = storageFormat });
        OneNotePage roundTrip = Assert.Single(OneNoteSectionReader.Read(new MemoryStream(data)).Pages);

        OneNotePage nestedVersion = Assert.Single(Assert.Single(roundTrip.ConflictPages).VersionHistory);
        Assert.Equal("Conflict version", nestedVersion.Title);
        OneNoteImage image = Assert.IsType<OneNoteImage>(Assert.Single(Assert.Single(nestedVersion.Outlines).Children));
        Assert.Equal(new byte[] { 1, 2, 3 }, image.Payload!.ToArray(16));

        OneNotePage nestedConflict = Assert.Single(Assert.Single(roundTrip.VersionHistory).ConflictPages);
        Assert.Equal("Version conflict", nestedConflict.Title);
        OneNoteEmbeddedFile file = Assert.IsType<OneNoteEmbeddedFile>(Assert.Single(Assert.Single(nestedConflict.Outlines).Children));
        Assert.Equal(new byte[] { 4, 5, 6 }, file.Payload!.ToArray(16));
    }

    [Fact]
    public void LoadedFssHttpSectionPreservesItsStorageFormatByDefault() {
        OneNoteSection source = OneNoteSectionReader.Read(Path.Combine(AppContext.BaseDirectory, "Fixtures", "testOneNoteFromOffice365.one"));
        Assert.Equal(OneNoteStorageFormat.FileSynchronizationPackage, source.StorageFormat);

        byte[] data = OneNoteSectionWriter.Write(source);

        Assert.Equal(OneNoteStorageFormat.FileSynchronizationPackage, OneNoteFileProbe.ReadHeader(new MemoryStream(data)).StorageFormat);
        OneNoteSection roundTrip = OneNoteSectionReader.Read(new MemoryStream(data));
        Assert.Equal(source.Pages.Select(page => page.Title), roundTrip.Pages.Select(page => page.Title));
    }

    [Fact]
    public void EmitsRequiredHeaderCellAndOfficialFileNameCrc() {
        Assert.Equal(0xCEBE8422U, OneNoteCrc32.ComputeFileName("Example.one"));
        OneNoteSection section = CreateSection();
        OneNoteWriteGraph graph = new OneNoteWriteGraphBuilder().BuildSection(section, Guid.NewGuid(), "Example.one");

        OneNoteRevisionStore store = OneNoteRevisionStoreReader.Read(new MemoryStream(OneNotePackageStoreWriter.Write(graph)));

        OneNoteRevisionStoreObject header = Assert.Single(store.Objects, item => item.Id.Identifier == new Guid("B4760B1A-FBDF-4AE3-9D08-53219D8A8D21"));
        Assert.Equal(OneNoteSchema.JcidPropertyContainer, header.Jcid.Value);
        Assert.Equal(0xCEBE8422U, header.PropertySet!.Find(OneNoteSchema.FileNameCrc)!.ScalarValue);
    }

    private static OneNoteSection CreateSection() {
        var section = new OneNoteSection { Name = "Writer sample", ColorArgb = 0xFF336699U };
        var parent = new OneNotePage {
            Title = "Parent page",
            Level = 0,
            CreatedUtc = new DateTime(2026, 7, 15, 10, 0, 0, DateTimeKind.Utc),
            LastModifiedUtc = new DateTime(2026, 7, 15, 11, 0, 0, DateTimeKind.Utc),
            OriginalAuthor = "OfficeIMO"
        };
        var outline = new OneNoteOutline { Layout = new OneNoteLayout { X = 10, Y = 20, Width = 400 } };
        var paragraph = new OneNoteParagraph();
        var first = new OneNoteTextRun { Text = "Hello " };
        first.Style.Bold = true;
        first.Style.FontFamily = "Aptos";
        first.Style.FontSize = 12;
        first.Style.LanguageId = 0x0415U;
        var second = new OneNoteTextRun { Text = "world", Hyperlink = "https://example.test/" };
        second.Style.Italic = true;
        paragraph.Runs.Add(first);
        paragraph.Runs.Add(second);
        outline.Children.Add(paragraph);
        parent.Outlines.Add(outline);
        section.Pages.Add(parent);
        section.Pages.Add(new OneNotePage { Title = "Child page", Level = 1 });
        return section;
    }

    private static OneNoteSection CreateVersionedSection() {
        DateTime created = new DateTime(2026, 7, 15, 9, 0, 0, DateTimeKind.Utc);
        var section = new OneNoteSection { Name = "Versions" };
        var page = new OneNotePage {
            Title = "Current page",
            CreatedUtc = created,
            LastModifiedUtc = created.AddHours(2),
            MostRecentAuthor = "Current author"
        };
        var currentOutline = new OneNoteOutline();
        var currentParagraph = new OneNoteParagraph();
        currentParagraph.Runs.Add(new OneNoteTextRun { Text = "Current content" });
        currentOutline.Children.Add(currentParagraph);
        page.Outlines.Add(currentOutline);

        var version = new OneNotePage {
            Title = "Historical page",
            CreatedUtc = created,
            LastModifiedUtc = created.AddHours(1),
            MostRecentAuthor = "Historical author",
            IsVersionHistoryPage = true
        };
        var versionOutline = new OneNoteOutline();
        var versionParagraph = new OneNoteParagraph();
        versionParagraph.Runs.Add(new OneNoteTextRun { Text = "Historical content" });
        versionOutline.Children.Add(versionParagraph);
        version.Outlines.Add(versionOutline);
        page.VersionHistory.Add(version);
        section.Pages.Add(page);
        return section;
    }

    private static void AssertVersionHistoryRoundTrip(byte[] data) {
        OneNoteRevisionStore store = OneNoteRevisionStoreReader.Read(new MemoryStream(data));
        OneNoteSection roundTrip = OneNoteSectionReader.Read(new MemoryStream(data));
        OneNotePage current = Assert.Single(roundTrip.Pages);
        OneNotePage version = Assert.Single(current.VersionHistory);
        Assert.Equal("Current page", current.Title);
        Assert.Equal("Historical page", version.Title);
        Assert.True(version.IsVersionHistoryPage);
        Assert.NotNull(version.RevisionContextId);
        Assert.Equal("Historical author", version.MostRecentAuthor);
        OneNoteParagraph paragraph = Assert.IsType<OneNoteParagraph>(Assert.Single(Assert.Single(version.Outlines).Children));
        Assert.Equal("Historical content", string.Concat(paragraph.Runs.Select(run => run.Text)));

        OneNoteExtendedGuid pageSpaceId = Assert.IsType<OneNoteExtendedGuid>(current.Id);
        OneNoteRevisionManifest[] pageRevisions = store.Revisions.Where(revision => pageSpaceId.Equals(revision.ObjectSpaceId)).ToArray();
        Assert.Contains(pageRevisions, revision => revision.RoleAssociations.Any(association =>
            association.Role == 1 &&
            association.ContextId?.Identifier == new Guid("7111497F-1B6B-4209-9491-C98B04CF4C5A") &&
            association.ContextId.Value == 1));
        Assert.Contains(pageRevisions, revision => revision.RoleAssociations.Any(association =>
            association.Role == 1 && association.ContextId != null && association.ContextId.Equals(version.RevisionContextId)));
    }

    private static int ReadCommittedListCount(byte[] data, int transactionOffset) {
        int count = 0;
        while (BitConverter.ToUInt32(data, transactionOffset + (count * 8)) != 1U) count++;
        return count;
    }

    private static OneNoteWriteProperty Property(OneNoteWriteObject item, uint id) =>
        Assert.Single(item.Properties, property => (property.RawId & 0x7FFFFFFFU) == id);

    private static bool IsTrue(OneNoteWriteProperty property) => (property.RawId & 0x80000000U) != 0;

    private static float FloatValue(OneNoteWriteProperty property) =>
        BitConverter.ToSingle(BitConverter.GetBytes((uint)Assert.IsType<ulong>(property.Scalar)), 0);

    private static void AssertGuidProperty(OneNoteWriteObject item, uint id) {
        byte[] data = Assert.IsType<byte[]>(Property(item, id).Data);
        Assert.Equal(16, data.Length);
        Assert.NotEqual(Guid.Empty, new Guid(data));
    }
}

namespace OfficeIMO.OneNote.Tests;

public sealed class IdentityStabilityTests {
    [Theory]
    [InlineData(OneNoteStorageFormat.RevisionStore)]
    [InlineData(OneNoteStorageFormat.FileSynchronizationPackage)]
    public void RepeatedSectionWritesRetainAssignedLogicalIdentities(OneNoteStorageFormat storageFormat) {
        OneNoteSection section = CreateSection("Standalone");
        var options = new OneNoteWriterOptions { StorageFormat = storageFormat };

        byte[] first = OneNoteSectionWriter.Write(section, options);
        AssertSectionIdentitiesAssigned(section);
        string[] assigned = CaptureSectionIdentities(section);
        byte[] second = OneNoteSectionWriter.Write(section, options);

        OneNoteSection firstRead = OneNoteSectionReader.Read(new MemoryStream(first));
        OneNoteSection secondRead = OneNoteSectionReader.Read(new MemoryStream(second));
        Assert.Equal(assigned, CaptureSectionIdentities(section));
        Assert.Equal(assigned, CaptureSectionIdentities(firstRead));
        Assert.Equal(assigned, CaptureSectionIdentities(secondRead));
    }

    [Fact]
    public void RepeatedNotebookWritesRetainHierarchyAndNestedTocIdentities() {
        var notebook = new OneNoteNotebook { Name = "Stable notebook" };
        notebook.Sections.Add(CreateSection("Root"));
        var group = new OneNoteSectionGroup { Name = "Group" };
        group.Sections.Add(CreateSection("Nested"));
        notebook.SectionGroups.Add(group);

        byte[] firstToc = OneNoteTableOfContentsWriter.Write(notebook);
        AssertNotebookIdentitiesAssigned(notebook, includeSectionContent: false);
        string[] assignedToc = CaptureNotebookIdentities(notebook);
        byte[] secondToc = OneNoteTableOfContentsWriter.Write(notebook);
        Assert.Equal(
            CaptureNotebookIdentities(OneNoteNotebookReader.Read(new MemoryStream(firstToc))),
            CaptureNotebookIdentities(OneNoteNotebookReader.Read(new MemoryStream(secondToc))));
        Assert.Equal(assignedToc, CaptureNotebookIdentities(notebook));

        byte[] firstPackage = OneNotePackageWriter.Write(notebook);
        AssertNotebookIdentitiesAssigned(notebook, includeSectionContent: true);
        string[] assignedPackage = CaptureNotebookIdentities(notebook);
        byte[] secondPackage = OneNotePackageWriter.Write(notebook);
        OneNoteNotebook firstRead = OneNotePackageReader.Read(new MemoryStream(firstPackage), "first.onepkg");
        OneNoteNotebook secondRead = OneNotePackageReader.Read(new MemoryStream(secondPackage), "second.onepkg");

        Assert.Equal(assignedPackage, CaptureNotebookIdentities(notebook));
        Assert.Equal(assignedPackage, CaptureNotebookIdentities(firstRead));
        Assert.Equal(assignedPackage, CaptureNotebookIdentities(secondRead));
    }

    [Theory]
    [InlineData(OneNoteStorageFormat.RevisionStore)]
    [InlineData(OneNoteStorageFormat.FileSynchronizationPackage)]
    public void IndependentlyEditedSharedAuthorsReceiveStableDistinctIdentities(OneNoteStorageFormat storageFormat) {
        var section = new OneNoteSection { Name = "Authors" };
        var page = new OneNotePage { Title = "Authors" };
        var shared = new OneNoteAuthor { Name = "Shared author" };
        OneNoteParagraph first = Paragraph("First");
        first.Author = shared;
        OneNoteParagraph second = Paragraph("Second");
        second.Author = shared;
        page.DirectContent.Add(first);
        page.DirectContent.Add(second);
        section.Pages.Add(page);
        var options = new OneNoteWriterOptions { StorageFormat = storageFormat };

        OneNoteSection loaded = OneNoteSectionReader.Read(new MemoryStream(OneNoteSectionWriter.Write(section, options)));
        OneNoteParagraph[] loadedParagraphs = BodyParagraphs(loaded.Pages[0]);
        OneNoteParagraph loadedFirst = loadedParagraphs[0];
        OneNoteParagraph loadedSecond = loadedParagraphs[1];
        Assert.Equal(loadedFirst.Author!.ObjectId, loadedSecond.Author!.ObjectId);

        loadedSecond.Author.Name = "Edited author";
        OneNoteSection edited = OneNoteSectionReader.Read(new MemoryStream(OneNoteSectionWriter.Write(loaded, options)));
        OneNoteParagraph[] editedParagraphs = BodyParagraphs(edited.Pages[0]);
        OneNoteParagraph editedFirst = editedParagraphs[0];
        OneNoteParagraph editedSecond = editedParagraphs[1];
        Assert.Equal("Shared author", editedFirst.Author!.Name);
        Assert.Equal("Edited author", editedSecond.Author!.Name);
        Assert.NotEqual(editedFirst.Author.ObjectId, editedSecond.Author.ObjectId);

        OneNoteSection repeated = OneNoteSectionReader.Read(new MemoryStream(OneNoteSectionWriter.Write(edited, options)));
        OneNoteParagraph[] repeatedParagraphs = BodyParagraphs(repeated.Pages[0]);
        Assert.Equal(editedFirst.Author.ObjectId, repeatedParagraphs[0].Author!.ObjectId);
        Assert.Equal(editedSecond.Author.ObjectId, repeatedParagraphs[1].Author!.ObjectId);
    }

    private static OneNoteParagraph[] BodyParagraphs(OneNotePage page) {
        Assert.Empty(page.DirectContent);
        return Assert.Single(page.Outlines).Children.OfType<OneNoteParagraph>().ToArray();
    }

    private static OneNoteSection CreateSection(string name) {
        DateTime created = new DateTime(2026, 7, 16, 8, 0, 0, DateTimeKind.Utc);
        var section = new OneNoteSection { Name = name };
        var page = new OneNotePage {
            Title = name + " page",
            CreatedUtc = created,
            LastModifiedUtc = created.AddMinutes(5)
        };

        var outline = new OneNoteOutline();
        outline.Children.Add(Paragraph("Outlined text"));
        page.Outlines.Add(outline);

        OneNoteParagraph listed = Paragraph("Bold", " link");
        listed.Runs[0].Style.Bold = true;
        listed.Runs[1].Hyperlink = "https://example.invalid/item";
        listed.List = new OneNoteListInfo { Ordered = true, Level = 1, DisplayIndex = 2 };
        listed.Style.Alignment = OneNoteParagraphAlignment.Center;
        listed.Author = new OneNoteAuthor { Name = "Stable author" };
        listed.Tags.Add(new OneNoteTag {
            ActionItemType = 0,
            Label = "Important",
            IsCheckable = false
        });
        page.DirectContent.Add(listed);

        var table = new OneNoteTable { BordersVisible = true };
        table.ColumnWidths.Add(180);
        var row = new OneNoteTableRow();
        var cell = new OneNoteTableCell { ShadingColorArgb = 0xFFF0F0F0U };
        cell.Content.Add(Paragraph("Cell"));
        row.Cells.Add(cell);
        table.Rows.Add(row);
        page.DirectContent.Add(table);

        page.DirectContent.Add(new OneNoteImage {
            FileName = "diagram.png",
            MediaType = "image/png",
            Payload = OneNoteBinaryPayload.FromBytes(new byte[] { 1, 2, 3 })
        });
        page.DirectContent.Add(new OneNoteEmbeddedFile {
            FileName = "notes.txt",
            MediaType = "text/plain",
            Payload = OneNoteBinaryPayload.FromBytes(new byte[] { 4, 5, 6 })
        });
        page.DirectContent.Add(new OneNoteMath { Text = "x + y" });

        var conflict = new OneNotePage {
            Title = "Conflict",
            IsConflictPage = true,
            CreatedUtc = created,
            LastModifiedUtc = created
        };
        conflict.DirectContent.Add(Paragraph("Conflict text"));
        page.ConflictPages.Add(conflict);

        var version = new OneNotePage {
            Title = "Version",
            IsVersionHistoryPage = true,
            CreatedUtc = created,
            LastModifiedUtc = created
        };
        version.DirectContent.Add(Paragraph("Version text"));
        page.VersionHistory.Add(version);

        section.Pages.Add(page);
        return section;
    }

    private static OneNoteParagraph Paragraph(params string[] runs) {
        var paragraph = new OneNoteParagraph();
        foreach (string text in runs) paragraph.Runs.Add(new OneNoteTextRun { Text = text });
        return paragraph;
    }

    private static void AssertNotebookIdentitiesAssigned(OneNoteNotebook notebook, bool includeSectionContent) {
        AssertValid(notebook.Id);
        Assert.NotNull(notebook.TableOfContentsRootObjectId);
        foreach (OneNoteSection section in notebook.Sections) {
            if (includeSectionContent) AssertSectionIdentitiesAssigned(section);
            else AssertValid(section.Id);
        }
        foreach (OneNoteSectionGroup group in notebook.SectionGroups) {
            AssertGroupIdentitiesAssigned(group, includeSectionContent);
        }
    }

    private static void AssertGroupIdentitiesAssigned(OneNoteSectionGroup group, bool includeSectionContent) {
        AssertValid(group.Id);
        if (!includeSectionContent) return;
        Assert.NotNull(group.TableOfContentsRootObjectId);
        foreach (OneNoteSection section in group.Sections) {
            AssertSectionIdentitiesAssigned(section);
        }
        foreach (OneNoteSectionGroup child in group.SectionGroups) {
            AssertGroupIdentitiesAssigned(child, includeSectionContent: true);
        }
    }

    private static void AssertSectionIdentitiesAssigned(OneNoteSection section) {
        AssertValid(section.Id);
        foreach (OneNotePage page in section.Pages) AssertPageIdentitiesAssigned(page);
    }

    private static void AssertPageIdentitiesAssigned(OneNotePage page) {
        Assert.NotNull(page.Id);
        Assert.NotNull(page.PreservationIds.ManifestId);
        Assert.NotNull(page.PreservationIds.MetadataId);
        Assert.NotNull(page.PreservationIds.RevisionMetadataId);
        Assert.NotNull(page.PreservationIds.PageNodeId);
        Assert.NotNull(page.PreservationIds.TitleNodeId);
        Assert.NotNull(page.PreservationIds.TitleOutlineId);
        Assert.NotNull(page.PreservationIds.TitleElementId);
        Assert.NotNull(page.PreservationIds.TitleTextId);
        if (page.IsVersionHistoryPage) {
            Assert.NotNull(page.RevisionContextId);
            Assert.NotNull(page.PreservationIds.VersionProxyId);
        }
        foreach (OneNoteOutline outline in page.Outlines) AssertElementIdentitiesAssigned(outline);
        foreach (OneNoteElement element in page.DirectContent) AssertElementIdentitiesAssigned(element);
        foreach (OneNotePage conflict in page.ConflictPages) AssertPageIdentitiesAssigned(conflict);
        foreach (OneNotePage version in page.VersionHistory) AssertPageIdentitiesAssigned(version);
    }

    private static void AssertElementIdentitiesAssigned(OneNoteElement element) {
        Assert.NotNull(element.Id);
        if (element.Author != null) Assert.NotNull(element.Author.ObjectId);
        foreach (OneNoteTag tag in element.Tags.Where(tag => tag.ActionItemType < 100)) Assert.NotNull(tag.DefinitionId);
        if (element is OneNoteOutline outline) {
            if (outline.WrapperList != null) Assert.NotNull(outline.WrapperList.ObjectId);
            foreach (OneNoteElement child in outline.Children) AssertElementIdentitiesAssigned(child);
        } else if (element is OneNoteParagraph paragraph) {
            Assert.NotNull(paragraph.ContentObjectId);
            if (paragraph.List != null) Assert.NotNull(paragraph.List.ObjectId);
            if (paragraph.Runs.Count > 1) {
                foreach (OneNoteTextRun run in paragraph.Runs) Assert.NotNull(run.StyleObjectId);
            }
            foreach (OneNoteElement child in paragraph.Children) AssertElementIdentitiesAssigned(child);
        } else if (element is OneNoteTable table) {
            foreach (OneNoteTableRow row in table.Rows) {
                Assert.NotNull(row.ObjectId);
                foreach (OneNoteTableCell cell in row.Cells) {
                    Assert.NotNull(cell.ObjectId);
                    foreach (OneNoteElement child in cell.Content) AssertElementIdentitiesAssigned(child);
                }
            }
        }
        if (element is OneNoteBinaryElement binary) {
            Assert.NotNull(binary.PayloadObjectId);
            Assert.True(binary.PayloadFileDataId.HasValue);
        }
    }

    private static void AssertValid(Guid? identity) {
        Assert.True(identity.HasValue);
        Assert.NotEqual(Guid.Empty, identity.Value);
    }

    private static string[] CaptureNotebookIdentities(OneNoteNotebook notebook) {
        var result = new List<string>();
        Add(result, "notebook", notebook.Id);
        Add(result, "notebook/toc-root", notebook.TableOfContentsRootObjectId);
        for (int index = 0; index < notebook.Sections.Count; index++) {
            AddSection(result, "section[" + index + "]", notebook.Sections[index]);
        }
        for (int index = 0; index < notebook.SectionGroups.Count; index++) {
            AddGroup(result, "group[" + index + "]", notebook.SectionGroups[index]);
        }
        return result.ToArray();
    }

    private static void AddGroup(ICollection<string> result, string path, OneNoteSectionGroup group) {
        Add(result, path, group.Id);
        Add(result, path + "/toc-root", group.TableOfContentsRootObjectId);
        for (int index = 0; index < group.Sections.Count; index++) {
            AddSection(result, path + "/section[" + index + "]", group.Sections[index]);
        }
        for (int index = 0; index < group.SectionGroups.Count; index++) {
            AddGroup(result, path + "/group[" + index + "]", group.SectionGroups[index]);
        }
    }

    private static string[] CaptureSectionIdentities(OneNoteSection section) {
        var result = new List<string>();
        AddSection(result, "section", section);
        return result.ToArray();
    }

    private static void AddSection(ICollection<string> result, string path, OneNoteSection section) {
        Add(result, path, section.Id);
        for (int index = 0; index < section.Pages.Count; index++) {
            AddPage(result, path + "/page[" + index + "]", section.Pages[index]);
        }
    }

    private static void AddPage(ICollection<string> result, string path, OneNotePage page) {
        Add(result, path, page.Id);
        Add(result, path + "/context", page.RevisionContextId);
        Add(result, path + "/manifest", page.PreservationIds.ManifestId);
        Add(result, path + "/metadata", page.PreservationIds.MetadataId);
        Add(result, path + "/revision-metadata", page.PreservationIds.RevisionMetadataId);
        Add(result, path + "/page-node", page.PreservationIds.PageNodeId);
        Add(result, path + "/title-node", page.PreservationIds.TitleNodeId);
        Add(result, path + "/title-outline", page.PreservationIds.TitleOutlineId);
        Add(result, path + "/title-element", page.PreservationIds.TitleElementId);
        Add(result, path + "/title-text", page.PreservationIds.TitleTextId);
        Add(result, path + "/version-proxy", page.PreservationIds.VersionProxyId);
        for (int index = 0; index < page.Outlines.Count; index++) {
            AddElement(result, path + "/outline[" + index + "]", page.Outlines[index]);
        }
        for (int index = 0; index < page.DirectContent.Count; index++) {
            AddElement(result, path + "/direct[" + index + "]", page.DirectContent[index]);
        }
        for (int index = 0; index < page.ConflictPages.Count; index++) {
            AddPage(result, path + "/conflict[" + index + "]", page.ConflictPages[index]);
        }
        for (int index = 0; index < page.VersionHistory.Count; index++) {
            AddPage(result, path + "/version[" + index + "]", page.VersionHistory[index]);
        }
    }

    private static void AddElement(ICollection<string> result, string path, OneNoteElement element) {
        Add(result, path, element.Id);
        Add(result, path + "/author", element.Author?.ObjectId);
        for (int index = 0; index < element.Tags.Count; index++) {
            Add(result, path + "/tag[" + index + "]", element.Tags[index].DefinitionId);
        }
        if (element is OneNoteOutline outline) {
            Add(result, path + "/wrapper-list", outline.WrapperList?.ObjectId);
            for (int index = 0; index < outline.Children.Count; index++) {
                AddElement(result, path + "/child[" + index + "]", outline.Children[index]);
            }
        } else if (element is OneNoteParagraph paragraph) {
            Add(result, path + "/content", paragraph.ContentObjectId);
            Add(result, path + "/list", paragraph.List?.ObjectId);
            Add(result, path + "/paragraph-style", paragraph.Style.ObjectId);
            for (int index = 0; index < paragraph.Runs.Count; index++) {
                Add(result, path + "/run-style[" + index + "]", paragraph.Runs[index].StyleObjectId);
            }
            for (int index = 0; index < paragraph.Children.Count; index++) {
                AddElement(result, path + "/child[" + index + "]", paragraph.Children[index]);
            }
        } else if (element is OneNoteTable table) {
            for (int rowIndex = 0; rowIndex < table.Rows.Count; rowIndex++) {
                OneNoteTableRow row = table.Rows[rowIndex];
                string rowPath = path + "/row[" + rowIndex + "]";
                Add(result, rowPath, row.ObjectId);
                for (int cellIndex = 0; cellIndex < row.Cells.Count; cellIndex++) {
                    OneNoteTableCell cell = row.Cells[cellIndex];
                    string cellPath = rowPath + "/cell[" + cellIndex + "]";
                    Add(result, cellPath, cell.ObjectId);
                    for (int index = 0; index < cell.Content.Count; index++) {
                        AddElement(result, cellPath + "/content[" + index + "]", cell.Content[index]);
                    }
                }
            }
        }
        if (element is OneNoteBinaryElement binary) {
            Add(result, path + "/payload-object", binary.PayloadObjectId);
            Add(result, path + "/payload-data", binary.PayloadFileDataId);
        }
    }

    private static void Add(ICollection<string> result, string path, Guid? identity) {
        if (identity.HasValue) result.Add(path + "=" + identity.Value.ToString("D"));
    }

    private static void Add(ICollection<string> result, string path, OneNoteExtendedGuid? identity) {
        if (identity != null) result.Add(path + "=" + identity);
    }
}

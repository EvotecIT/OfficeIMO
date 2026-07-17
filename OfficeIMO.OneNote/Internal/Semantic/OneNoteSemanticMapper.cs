namespace OfficeIMO.OneNote;

internal static partial class OneNoteSemanticMapper {
    private const uint JcidSectionNode = 0x00060007;
    private const uint JcidPageSeriesNode = 0x00060008;
    private const uint JcidPageNode = 0x0006000B;
    private const uint JcidOutlineNode = 0x0006000C;
    private const uint JcidOutlineElementNode = 0x0006000D;
    private const uint JcidRichTextNode = 0x0006000E;
    private const uint JcidImageNode = 0x00060011;
    private const uint JcidNumberListNode = 0x00060012;
    private const uint JcidOutlineGroup = 0x00060019;
    private const uint JcidTableNode = 0x00060022;
    private const uint JcidTableRowNode = 0x00060023;
    private const uint JcidTableCellNode = 0x00060024;
    private const uint JcidTitleNode = 0x0006002C;
    private const uint JcidPageMetadata = 0x00020030;
    private const uint JcidSectionMetadata = 0x00020031;
    private const uint JcidEmbeddedFileNode = 0x00060035;
    private const uint JcidPageManifestNode = 0x00060037;
    private const uint JcidConflictPageMetadata = 0x00020038;
    private const uint JcidRevisionMetadata = 0x00020044;
    private const uint JcidAuthor = 0x00120001;

    private const uint ContentChildNodes = 0x24001C1F;
    private const uint ElementChildNodes = 0x24001C20;
    private const uint ListNodes = 0x24001C26;
    private const uint RichEditTextUnicode = 0x1C001C22;
    private const uint ChildGraphSpaceElementNodes = 0x2C001D63;
    private const uint CachedTitleString = 0x1C001CF3;
    private const uint CachedTitleStringFromPage = 0x1C001D3C;
    private const uint StructureElementChildNodes = 0x24001D5F;
    private const uint Author = 0x1C001D75;
    private const uint AuthorOriginal = 0x20001D78;
    private const uint AuthorMostRecent = 0x20001D79;
    private const uint SectionDisplayName = 0x1C00349B;
    private const uint NotebookColor = 0x14001CBE;
    private const uint PageLevel = 0x14001DFF;
    private const uint TopologyCreationTimestamp = 0x18001C65;
    private const uint LastModifiedTimestamp = 0x18001D77;
    private const uint LastModifiedTime = 0x14001D7A;
    private const uint IsConflictPage = 0x08001D7C;
    private const uint IsDeletedGraphSpaceContent = 0x1C001DE9;
    private const uint PictureContainer = 0x20001C3F;
    private const uint EmbeddedFileContainer = 0x20001D9B;
    private const uint EmbeddedFileName = 0x1C001D9C;
    private const uint SourceFilePath = 0x1C001D9D;
    private const uint ImageFilename = 0x1C001DD7;
    private const uint ImageAltText = 0x1C001E58;
    private const uint IRecordMedia = 0x14001D24;
    private const uint PictureWidth = 0x140034CD;
    private const uint PictureHeight = 0x140034CE;
    private const uint PageWidth = 0x14001C01;
    private const uint PageHeight = 0x14001C02;
    private const uint LayoutMaxWidth = 0x14001C1B;
    private const uint LayoutMaxHeight = 0x14001C1C;
    private const uint OffsetFromParentHorizontal = 0x14001C14;
    private const uint OffsetFromParentVertical = 0x14001C15;
    private const uint LayoutTightLayout = 0x08001C00;
    private const uint OutlineElementRtl = 0x08001C34;
    private const uint TableBordersVisible = 0x08001D5E;
    private const uint TableColumnWidths = 0x1C001D66;
    private const uint CellShadingColor = 0x14001E26;
    private const uint TextRunIndex = 0x1C001E12;
    private const uint TextRunFormatting = 0x24001E13;
    private const uint Hyperlink = 0x08001E14;
    private const uint HyperlinkProtected = 0x08001E19;
    private const uint HyperlinkUrl = 0x1C001E20;
    private const uint ParagraphStyle = 0x2000342C;
    private const uint ParagraphStyleId = 0x1C00345A;
    private const uint Bold = 0x08001C04;
    private const uint Italic = 0x08001C05;
    private const uint Underline = 0x08001C06;
    private const uint Strikethrough = 0x08001C07;
    private const uint Superscript = 0x08001C08;
    private const uint Subscript = 0x08001C09;
    private const uint Font = 0x1C001C0A;
    private const uint FontSize = 0x10001C0B;
    private const uint FontColor = 0x14001C0C;
    private const uint Highlight = 0x14001C0D;
    private const uint ParagraphAlignment = 0x0C003477;
    private const uint ParagraphSpaceBefore = 0x1400342E;
    private const uint ParagraphSpaceAfter = 0x1400342F;
    private const uint ParagraphLineSpacingExact = 0x14003430;
    private const uint TextExtendedAscii = 0x1C003498;
    private const uint MathFormatting = 0x08003401;

    public static OneNoteSection MapSection(OneNoteRevisionStore store, OneNoteReaderOptions options) {
        var materializer = new OneNoteObjectSpaceMaterializer(store);
        var pageMapping = new PageMappingState(options);
        OneNoteMaterializedObjectSpace sectionSpace = materializer.FindCurrentSpaceByRootJcid(
            JcidSectionNode,
            "ONENOTE_SECTION_OBJECT_SPACE",
            "No current section object space could be materialized.");
        var section = new OneNoteSection {
            Id = store.Header.FileId,
            Name = string.Empty,
            StorageFormat = store.Header.StorageFormat
        };

        OneNoteRevisionStoreObject? metadata = sectionSpace.GetRoot(2);
        if (metadata?.Jcid.Value == JcidSectionMetadata) {
            section.Name = ReadString(metadata, SectionDisplayName) ?? string.Empty;
            section.ColorArgb = ReadUInt32(metadata, NotebookColor);
        }
        foreach (OneNoteRevisionManifest manifest in materializer.GetRevisionChain(sectionSpace.Revision)) {
            section.Revisions.Add(MapRevision(manifest, manifest == sectionSpace.Revision));
        }

        OneNoteRevisionStoreObject? sectionRoot = sectionSpace.GetRoot(1);
        if (sectionRoot == null || sectionRoot.Jcid.Value != JcidSectionNode) {
            throw new OneNoteFormatException("ONENOTE_SECTION_ROOT", "The current root object space does not resolve to a section node.");
        }
        foreach (OneNoteExtendedGuid pageSeriesId in GetReferences(sectionRoot, ElementChildNodes)) {
            OneNoteRevisionStoreObject? pageSeries = sectionSpace.GetObject(pageSeriesId);
            if (pageSeries?.Jcid.Value != JcidPageSeriesNode) continue;
            foreach (OneNoteExtendedGuid objectSpaceId in GetReferences(pageSeries, ChildGraphSpaceElementNodes)) {
                OneNotePage? page = MapPage(materializer, objectSpaceId, options, pageMapping);
                if (page != null) section.Pages.Add(page);
            }
        }
        PreserveUnknownObjects(section.UnknownObjects, sectionSpace);
        if (options.PreserveUnknownData) {
            section.PreservationState = OneNoteSectionPreservationState.Capture(materializer, sectionSpace, section.Pages);
        }
        return section;
    }

    private static OneNotePage? MapPage(
        OneNoteObjectSpaceMaterializer materializer,
        OneNoteExtendedGuid objectSpaceId,
        OneNoteReaderOptions options,
        PageMappingState pageMapping) {
        return MapPage(materializer, objectSpaceId, null, false, options, pageMapping);
    }

    private static OneNotePage? MapPage(
        OneNoteObjectSpaceMaterializer materializer,
        OneNoteExtendedGuid objectSpaceId,
        OneNoteExtendedGuid? contextId,
        bool isVersionHistoryPage,
        OneNoteReaderOptions options,
        PageMappingState pageMapping) {
        string pageKey = OneNoteObjectSpaceMaterializer.GetSpaceKey(objectSpaceId, contextId);
        if (!pageMapping.TryEnter(pageKey)) return null;
        try {
            OneNoteMaterializedObjectSpace? space = materializer.TryGetSpace(objectSpaceId, contextId);
            if (space == null) return null;
            OneNoteRevisionStoreObject? manifest = space.GetRoot(1);
            if (manifest?.Jcid.Value != JcidPageManifestNode) return null;
            OneNoteRevisionStoreObject? metadata = space.GetRoot(2);
            OneNoteRevisionStoreObject? revisionMetadata = space.GetRoot(4);
            OneNoteRevisionStoreObject? pageNode = GetReferences(manifest, ContentChildNodes)
                .Select(space.GetObject)
                .FirstOrDefault(item => item?.Jcid.Value == JcidPageNode);
            if (pageNode == null) return null;

            var page = new OneNotePage {
                Id = objectSpaceId,
                RevisionContextId = contextId,
                Title = ReadString(metadata, CachedTitleString) ?? ReadString(pageNode, CachedTitleStringFromPage) ?? string.Empty,
                Level = Math.Max(0, unchecked((int)(ReadUInt32(metadata, PageLevel) ?? 1U)) - 1),
                CreatedUtc = ReadFileTime(metadata, TopologyCreationTimestamp),
                LastModifiedUtc = ReadFileTime(revisionMetadata, LastModifiedTimestamp) ?? ReadTime32(pageNode, LastModifiedTime),
                OriginalAuthor = ReadString(pageNode, Author),
                MostRecentAuthor = ReadReferencedAuthor(space, revisionMetadata, AuthorMostRecent),
                IsConflictPage = ReadBoolean(metadata, IsConflictPage) ?? metadata?.Jcid.Value == JcidConflictPageMetadata,
                IsVersionHistoryPage = isVersionHistoryPage,
                IsDeleted = ReadData(metadata, IsDeletedGraphSpaceContent) != null,
                Width = ReadFloat(pageNode, PageWidth),
                Height = ReadFloat(pageNode, PageHeight)
            };
            page.PreservationIds.ManifestId = manifest.Id;
            page.PreservationIds.MetadataId = metadata?.Id;
            page.PreservationIds.RevisionMetadataId = revisionMetadata?.Id;
            page.PreservationIds.PageNodeId = pageNode.Id;
            CaptureTitleIds(page, space, pageNode);
            foreach (OneNoteRevisionManifest revision in materializer.GetRevisionChain(space.Revision)) {
                page.Revisions.Add(MapRevision(revision, revision == space.Revision));
            }

            foreach (OneNoteExtendedGuid childId in GetReferences(pageNode, ElementChildNodes)) {
                OneNoteElement? element = BuildElement(space, childId, materializer, options, 0, new HashSet<OneNoteExtendedGuid>());
                if (element is OneNoteOutline outline && !outline.IsOutlineElementWrapper) page.Outlines.Add(outline);
                else if (element != null) page.DirectContent.Add(element);
            }

            if (string.IsNullOrWhiteSpace(page.Title)) {
                foreach (OneNoteExtendedGuid titleId in GetReferences(pageNode, StructureElementChildNodes)) {
                    OneNoteRevisionStoreObject? title = space.GetObject(titleId);
                    if (title?.Jcid.Value != JcidTitleNode) continue;
                    page.Title = string.Join(" ", GetReferences(title, ElementChildNodes)
                        .Select(id => BuildElement(space, id, materializer, options, 0, new HashSet<OneNoteExtendedGuid>()))
                        .Where(element => element != null)
                        .SelectMany(EnumerateText)
                        .Where(text => !string.IsNullOrWhiteSpace(text))).Trim();
                }
            }

            foreach (OneNoteExtendedGuid conflictSpaceId in GetReferences(manifest, ChildGraphSpaceElementNodes)) {
                OneNotePage? conflict = MapPage(materializer, conflictSpaceId, options, pageMapping);
                if (conflict != null) {
                    conflict.IsConflictPage = true;
                    page.ConflictPages.Add(conflict);
                }
            }
            PreserveUnknownObjects(page.UnknownObjects, space);
            if (!isVersionHistoryPage) MapVersionHistory(page, materializer, objectSpaceId, manifest, options, pageMapping);
            return page;
        } finally {
            pageMapping.Exit(pageKey);
        }
    }

    private static OneNoteElement? BuildElement(
        OneNoteMaterializedObjectSpace space,
        OneNoteExtendedGuid id,
        OneNoteObjectSpaceMaterializer materializer,
        OneNoteReaderOptions options,
        int depth,
        HashSet<OneNoteExtendedGuid> path) {
        if (depth >= options.MaxPropertySetDepth || !path.Add(id)) return null;
        try {
            OneNoteRevisionStoreObject? item = space.GetObject(id);
            if (item == null) return null;
            switch (item.Jcid.Value) {
                case JcidOutlineNode:
                case JcidOutlineGroup: {
                    var outline = new OneNoteOutline { Id = id, Layout = ReadLayout(item) };
                    foreach (OneNoteExtendedGuid childId in GetReferences(item, ElementChildNodes)) {
                        OneNoteElement? child = BuildElement(space, childId, materializer, options, depth + 1, path);
                        if (child != null) outline.Children.Add(child);
                    }
                    return outline;
                }
                case JcidOutlineElementNode:
                    return BuildOutlineElement(space, item, materializer, options, depth, path);
                case JcidRichTextNode:
                    return BuildParagraph(space, item);
                case JcidImageNode:
                    return BuildImage(space, item, materializer);
                case JcidEmbeddedFileNode:
                    return BuildEmbeddedFile(space, item, materializer);
                case JcidTableNode:
                    return BuildTable(space, item, materializer, options, depth, path);
                default:
                    return null;
            }
        } finally {
            path.Remove(id);
        }
    }

    private static OneNoteElement BuildOutlineElement(
        OneNoteMaterializedObjectSpace space,
        OneNoteRevisionStoreObject item,
        OneNoteObjectSpaceMaterializer materializer,
        OneNoteReaderOptions options,
        int depth,
        HashSet<OneNoteExtendedGuid> path) {
        OneNoteElement? primary = GetReferences(item, ContentChildNodes)
            .Select(id => BuildElement(space, id, materializer, options, depth + 1, path))
            .FirstOrDefault(element => element != null);
        OneNoteParagraph paragraph = primary as OneNoteParagraph ?? new OneNoteParagraph();
        paragraph.Id = item.Id;
        paragraph.Layout = ReadLayout(item);
        paragraph.Author = ReadReferencedAuthorMetadata(space, item, AuthorMostRecent) ??
            ReadReferencedAuthorMetadata(space, item, AuthorOriginal);
        paragraph.List = BuildListInfo(space, item);
        foreach (OneNoteExtendedGuid childId in GetReferences(item, ElementChildNodes)) {
            OneNoteElement? child = BuildElement(space, childId, materializer, options, depth + 1, path);
            if (child != null) paragraph.Children.Add(child);
        }
        if (primary is OneNoteParagraph && paragraph.Children.Count == 0 && paragraph.List == null &&
            paragraph.Runs.Count > 0 && paragraph.Runs.All(run => run.Style.IsMath == true)) {
            var math = new OneNoteMath {
                Id = item.Id,
                Text = string.Concat(paragraph.Runs.Select(run => run.Text)),
                Layout = paragraph.Layout,
                Author = paragraph.Author
            };
            foreach (OneNoteTag tag in paragraph.Tags) math.Tags.Add(tag);
            return math;
        }
        if (primary != null && !(primary is OneNoteParagraph)) {
            var wrapper = new OneNoteOutline {
                Id = item.Id,
                Layout = paragraph.Layout,
                Author = paragraph.Author,
                IsOutlineElementWrapper = true,
                WrapperList = paragraph.List
            };
            wrapper.Children.Add(primary);
            foreach (OneNoteElement child in paragraph.Children) wrapper.Children.Add(child);
            ApplyTags(wrapper, item, space);
            return wrapper;
        }
        return paragraph;
    }

    private static OneNoteParagraph BuildParagraph(OneNoteMaterializedObjectSpace space, OneNoteRevisionStoreObject item) {
        string text = ReadString(item, RichEditTextUnicode) ?? ReadSingleByteString(item, TextExtendedAscii) ?? string.Empty;
        var paragraph = new OneNoteParagraph { Id = item.Id, ContentObjectId = item.Id, Layout = ReadLayout(item) };
        IReadOnlyList<uint> boundaries = ReadUInt32Array(item, TextRunIndex);
        IReadOnlyList<OneNoteExtendedGuid> styles = GetReferences(item, TextRunFormatting);
        int start = 0;
        int runCount = Math.Max(1, boundaries.Count + 1);
        for (int index = 0; index < runCount; index++) {
            int end = index < boundaries.Count ? ClampTextRunBoundary(boundaries[index], text.Length) : text.Length;
            if (end < start) end = start;
            var run = new OneNoteTextRun { Text = text.Substring(start, end - start) };
            if (index < styles.Count) ApplyTextStyle(run, space.GetObject(styles[index]));
            paragraph.Runs.Add(run);
            start = end;
        }
        if (paragraph.Runs.Count == 0) paragraph.Runs.Add(new OneNoteTextRun { Text = text });

        OneNoteRevisionStoreObject? paragraphStyle = GetReferences(item, ParagraphStyle).Select(space.GetObject).FirstOrDefault(style => style != null);
        if (paragraphStyle != null) {
            paragraph.Style.ObjectId = paragraphStyle.Id;
            paragraph.Style.StyleId = ReadString(paragraphStyle, ParagraphStyleId);
            uint? alignment = ReadUInt32(paragraphStyle, ParagraphAlignment);
            if (alignment.HasValue && alignment.Value <= 3) paragraph.Style.Alignment = (OneNoteParagraphAlignment)alignment.Value;
            paragraph.Style.SpaceBefore = ReadFloat(paragraphStyle, ParagraphSpaceBefore);
            paragraph.Style.SpaceAfter = ReadFloat(paragraphStyle, ParagraphSpaceAfter);
            paragraph.Style.ExactLineSpacing = ReadFloat(paragraphStyle, ParagraphLineSpacingExact);
        }
        ApplyTags(paragraph, item, space);
        return paragraph;
    }

    private static OneNoteImage BuildImage(OneNoteMaterializedObjectSpace space, OneNoteRevisionStoreObject item, OneNoteObjectSpaceMaterializer materializer) {
        var image = new OneNoteImage {
            Id = item.Id,
            FileName = ReadString(item, ImageFilename),
            AltText = ReadString(item, ImageAltText),
            SourcePath = ReadString(item, SourceFilePath),
            Hyperlink = ReadString(item, HyperlinkUrl),
            WidthHalfInches = ReadFloat(item, PictureWidth),
            HeightHalfInches = ReadFloat(item, PictureHeight),
            Layout = ReadLayout(item)
        };
        image.MediaType = ResolveMediaType(image.FileName);
        image.PictureContainerObjectId = GetReferences(item, PictureContainer).FirstOrDefault();
        image.WebPictureContainerObjectId = GetReferences(item, OneNoteSchema.WebPictureContainer14).FirstOrDefault();
        OneNoteRevisionStoreObject? picture = image.PictureContainerObjectId == null
            ? null
            : space.GetObject(image.PictureContainerObjectId);
        OneNoteRevisionStoreObject? webPicture = image.WebPictureContainerObjectId == null
            ? null
            : space.GetObject(image.WebPictureContainerObjectId);
        PopulateBinaryPayload(image, picture, materializer);
        if (image.Payload == null && webPicture != null) {
            PopulateBinaryPayload(image, webPicture, materializer);
            image.PayloadUsesWebPictureContainer = image.Payload != null;
        }
        ApplyTags(image, item, space);
        return image;
    }

    private static OneNoteBinaryElement BuildEmbeddedFile(OneNoteMaterializedObjectSpace space, OneNoteRevisionStoreObject item, OneNoteObjectSpaceMaterializer materializer) {
        OneNoteBinaryElement embedded = CreateEmbeddedElement(item);
        OneNoteRevisionStoreObject? binary = GetReferences(item, EmbeddedFileContainer).Select(space.GetObject).FirstOrDefault(value => value != null);
        PopulateBinaryPayload(embedded, binary, materializer);
        ApplyTags(embedded, item, space);
        return embedded;
    }

    internal static OneNoteBinaryElement CreateEmbeddedElement(OneNoteRevisionStoreObject item) {
        string? fileName = ReadString(item, EmbeddedFileName);
        string? sourcePath = ReadString(item, SourceFilePath);
        uint? recordingKind = ReadUInt32(item, IRecordMedia);
        OneNoteBinaryElement result;
        if (recordingKind == 1 || recordingKind == 2) {
            result = new OneNoteMedia {
                RecordingKind = recordingKind == 1 ? OneNoteMediaKind.Audio : OneNoteMediaKind.Video,
                SourcePath = sourcePath
            };
        } else {
            result = new OneNoteEmbeddedFile { SourcePath = sourcePath };
        }
        result.Id = item.Id;
        result.FileName = fileName;
        result.MediaType = ResolveMediaType(fileName);
        result.Layout = ReadLayout(item);
        return result;
    }

    private static string? ResolveMediaType(string? fileName) {
        switch (Path.GetExtension(fileName ?? string.Empty).ToLowerInvariant()) {
            case ".png": return "image/png";
            case ".jpg":
            case ".jpeg": return "image/jpeg";
            case ".gif": return "image/gif";
            case ".bmp": return "image/bmp";
            case ".tif":
            case ".tiff": return "image/tiff";
            case ".svg": return "image/svg+xml";
            case ".mp3": return "audio/mpeg";
            case ".wav": return "audio/wav";
            case ".wma": return "audio/x-ms-wma";
            case ".wmv": return "video/x-ms-wmv";
            case ".avi": return "video/x-msvideo";
            case ".mpg":
            case ".mpeg": return "video/mpeg";
            case ".mp4": return "video/mp4";
            case ".pdf": return "application/pdf";
            case ".docx": return "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
            case ".xlsx": return "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
            case ".pptx": return "application/vnd.openxmlformats-officedocument.presentationml.presentation";
            default: return null;
        }
    }

    private static OneNoteTable BuildTable(
        OneNoteMaterializedObjectSpace space,
        OneNoteRevisionStoreObject item,
        OneNoteObjectSpaceMaterializer materializer,
        OneNoteReaderOptions options,
        int depth,
        HashSet<OneNoteExtendedGuid> path) {
        var table = new OneNoteTable {
            Id = item.Id,
            BordersVisible = ReadBoolean(item, TableBordersVisible) ?? false,
            Layout = ReadLayout(item)
        };
        byte[]? widths = ReadData(item, TableColumnWidths);
        if (widths != null && widths.Length > 0) {
            int count = Math.Min(widths[0], (byte)((widths.Length - 1) / 4));
            for (int index = 0; index < count; index++) table.ColumnWidths.Add(BitConverter.ToSingle(widths, 1 + index * 4));
        }
        foreach (OneNoteExtendedGuid rowId in GetReferences(item, ElementChildNodes)) {
            OneNoteRevisionStoreObject? rowObject = space.GetObject(rowId);
            if (rowObject?.Jcid.Value != JcidTableRowNode) continue;
            var row = new OneNoteTableRow { ObjectId = rowObject.Id };
            foreach (OneNoteExtendedGuid cellId in GetReferences(rowObject, ElementChildNodes)) {
                OneNoteRevisionStoreObject? cellObject = space.GetObject(cellId);
                if (cellObject?.Jcid.Value != JcidTableCellNode) continue;
                var cell = new OneNoteTableCell { ObjectId = cellObject.Id, ShadingColorArgb = ReadUInt32(cellObject, CellShadingColor) };
                foreach (OneNoteExtendedGuid contentId in GetReferences(cellObject, ElementChildNodes)) {
                    OneNoteElement? content = BuildElement(space, contentId, materializer, options, depth + 1, path);
                    if (content != null) cell.Content.Add(content);
                }
                row.Cells.Add(cell);
            }
            table.Rows.Add(row);
        }
        ApplyTags(table, item, space);
        return table;
    }

    private static void PopulateBinaryPayload(
        OneNoteBinaryElement target,
        OneNoteRevisionStoreObject? binary,
        OneNoteObjectSpaceMaterializer materializer) {
        if (binary == null) return;
        target.PayloadObjectId = binary.Id;
        target.PayloadFileExtension = binary.FileExtension;
        if (materializer.TryResolveFileData(binary, out Guid fileDataId, out OneNoteBinaryPayload? payload)) {
            target.PayloadFileDataId = fileDataId;
            target.Payload = payload;
        }
    }

    private static OneNoteLayout ReadLayout(OneNoteRevisionStoreObject item) {
        return new OneNoteLayout {
            X = ReadFloat(item, OffsetFromParentHorizontal),
            Y = ReadFloat(item, OffsetFromParentVertical),
            Width = ReadFloat(item, LayoutMaxWidth),
            Height = ReadFloat(item, LayoutMaxHeight),
            Tight = ReadBoolean(item, LayoutTightLayout),
            RightToLeft = ReadBoolean(item, OutlineElementRtl)
        };
    }

    internal static void ApplyTextStyle(OneNoteTextRun target, OneNoteRevisionStoreObject? style) {
        if (style == null) return;
        target.StyleObjectId = style.Id;
        target.Style.Bold = ReadBoolean(style, Bold);
        target.Style.Italic = ReadBoolean(style, Italic);
        target.Style.Underline = ReadBoolean(style, Underline);
        target.Style.Strikethrough = ReadBoolean(style, Strikethrough);
        target.Style.Superscript = ReadBoolean(style, Superscript);
        target.Style.Subscript = ReadBoolean(style, Subscript);
        target.Style.FontFamily = ReadString(style, Font);
        ushort? fontSize = ReadUInt16(style, FontSize);
        if (fontSize.HasValue) target.Style.FontSize = fontSize.Value / 2.0;
        target.Style.ColorArgb = ReadUInt32(style, FontColor);
        target.Style.HighlightColorArgb = ReadUInt32(style, Highlight);
        target.Style.LanguageId = ReadUInt32(style, OneNoteSchema.LanguageId);
        target.Style.IsMath = ReadBoolean(style, MathFormatting);
        target.HyperlinkProtected = ReadBoolean(style, HyperlinkProtected) ?? false;
        if (ReadBoolean(style, Hyperlink) == true || target.HyperlinkProtected) {
            target.Hyperlink = ReadString(style, HyperlinkUrl);
        }
    }

    private static void CaptureTitleIds(OneNotePage page, OneNoteMaterializedObjectSpace space, OneNoteRevisionStoreObject pageNode) {
        OneNoteRevisionStoreObject? title = GetReferences(pageNode, StructureElementChildNodes)
            .Select(space.GetObject)
            .FirstOrDefault(item => item?.Jcid.Value == JcidTitleNode);
        if (title == null) return;
        page.PreservationIds.TitleNodeId = title.Id;
        OneNoteRevisionStoreObject? outline = GetReferences(title, ElementChildNodes).Select(space.GetObject).FirstOrDefault();
        page.PreservationIds.TitleOutlineId = outline?.Id;
        OneNoteRevisionStoreObject? element = GetReferences(outline, ElementChildNodes).Select(space.GetObject).FirstOrDefault();
        page.PreservationIds.TitleElementId = element?.Id;
        OneNoteRevisionStoreObject? text = GetReferences(element, ContentChildNodes).Select(space.GetObject).FirstOrDefault();
        page.PreservationIds.TitleTextId = text?.Id;
    }

    private static OneNoteRevision MapRevision(OneNoteRevisionManifest source, bool current) {
        return new OneNoteRevision {
            Id = source.Id,
            BaseRevisionId = source.DependencyId,
            IsCurrent = current,
            IsVersionHistory = source.ContextId != null
        };
    }

    private static void PreserveUnknownObjects(IList<OneNoteOpaqueObject> target, OneNoteMaterializedObjectSpace space) {
        int ordinal = 0;
        foreach (OneNoteRevisionStoreObject item in space.Objects) {
            if (!IsKnownJcid(item.Jcid)) target.Add(CreateOpaqueObject(item, ordinal));
            ordinal++;
        }
    }

    internal static OneNoteOpaqueObject CreateOpaqueObject(OneNoteRevisionStoreObject item, int ordinal) {
        var result = new OneNoteOpaqueObject {
            Id = item.Id,
            Jcid = item.Jcid.Value,
            Ordinal = ordinal
        };
        if (item.RawPropertyData != null) result.SetRawData(item.RawPropertyData.ToArray(int.MaxValue));
        if (item.PropertySet != null) {
            foreach (OneNotePropertyValue property in item.PropertySet.Properties) {
                var opaque = new OneNoteOpaqueProperty {
                    PropertyId = property.RawPropertyId,
                    ValueType = MapOpaqueValueType(property.Type),
                    Ordinal = property.Ordinal,
                    BooleanValue = property.BooleanValue,
                    ScalarValue = property.ScalarValue
                };
                if (property.Data != null) opaque.SetRawData(property.Data.ToArray(int.MaxValue));
                foreach (OneNoteExtendedGuid reference in property.ReferencedIds) opaque.ReferencedIds.Add(reference);
                result.Properties.Add(opaque);
            }
        }
        return result;
    }

    private static OneNotePropertyValueType MapOpaqueValueType(OneNotePropertyType type) {
        switch (type) {
            case OneNotePropertyType.NoData: return OneNotePropertyValueType.NoData;
            case OneNotePropertyType.Boolean: return OneNotePropertyValueType.Boolean;
            case OneNotePropertyType.Byte: return OneNotePropertyValueType.Byte;
            case OneNotePropertyType.UInt16: return OneNotePropertyValueType.UInt16;
            case OneNotePropertyType.UInt32: return OneNotePropertyValueType.UInt32;
            case OneNotePropertyType.UInt64: return OneNotePropertyValueType.UInt64;
            case OneNotePropertyType.LengthPrefixedData: return OneNotePropertyValueType.Blob;
            case OneNotePropertyType.ObjectId: return OneNotePropertyValueType.ObjectId;
            case OneNotePropertyType.ObjectIdArray: return OneNotePropertyValueType.ObjectIdArray;
            case OneNotePropertyType.ObjectSpaceId: return OneNotePropertyValueType.ObjectSpaceId;
            case OneNotePropertyType.ObjectSpaceIdArray: return OneNotePropertyValueType.ObjectSpaceIdArray;
            case OneNotePropertyType.ContextId: return OneNotePropertyValueType.ContextId;
            case OneNotePropertyType.ContextIdArray: return OneNotePropertyValueType.ContextIdArray;
            case OneNotePropertyType.PropertySet: return OneNotePropertyValueType.PropertySet;
            case OneNotePropertyType.PropertySetArray: return OneNotePropertyValueType.PropertySetArray;
            default: return OneNotePropertyValueType.Unknown;
        }
    }

    internal static bool IsKnownJcid(OneNoteJcid jcid) {
        if (jcid.IsFileData) return true;
        switch (jcid.Value) {
            case 0x00120001:
            case 0x00020001:
            case JcidSectionNode:
            case JcidPageSeriesNode:
            case JcidPageNode:
            case JcidOutlineNode:
            case JcidOutlineElementNode:
            case JcidRichTextNode:
            case JcidImageNode:
            case JcidNumberListNode:
            case JcidOutlineGroup:
            case JcidTableNode:
            case JcidTableRowNode:
            case JcidTableCellNode:
            case JcidTitleNode:
            case JcidPageMetadata:
            case JcidSectionMetadata:
            case JcidEmbeddedFileNode:
            case JcidPageManifestNode:
            case JcidConflictPageMetadata:
            case JcidVersionHistoryContent:
            case JcidVersionProxy:
            case 0x00120043:
            case JcidRevisionMetadata:
            case 0x00020046:
            case 0x0012004D:
                return true;
            default:
                return false;
        }
    }

    private static IEnumerable<string> EnumerateText(OneNoteElement? element) {
        if (element is OneNoteParagraph paragraph) {
            foreach (OneNoteTextRun run in paragraph.Runs) yield return run.Text;
            foreach (OneNoteElement child in paragraph.Children) {
                foreach (string text in EnumerateText(child)) yield return text;
            }
        } else if (element is OneNoteOutline outline) {
            foreach (OneNoteElement child in outline.Children) {
                foreach (string text in EnumerateText(child)) yield return text;
            }
        } else if (element is OneNoteTable table) {
            foreach (OneNoteTableRow row in table.Rows)
            foreach (OneNoteTableCell cell in row.Cells)
            foreach (OneNoteElement child in cell.Content)
            foreach (string text in EnumerateText(child)) yield return text;
        }
    }

    private static string? ReadReferencedAuthor(OneNoteMaterializedObjectSpace space, OneNoteRevisionStoreObject? item, uint propertyId) {
        return ReadReferencedAuthorMetadata(space, item, propertyId)?.Name;
    }

    private static OneNoteAuthor? ReadReferencedAuthorMetadata(
        OneNoteMaterializedObjectSpace space,
        OneNoteRevisionStoreObject? item,
        uint propertyId) {
        OneNoteRevisionStoreObject? author = GetReferences(item, propertyId)
            .Select(space.GetObject)
            .FirstOrDefault(value => value?.Jcid.Value == JcidAuthor);
        return author == null
            ? null
            : new OneNoteAuthor { ObjectId = author.Id, Name = ReadString(author, Author) };
    }

    internal static IReadOnlyList<OneNoteExtendedGuid> GetReferences(OneNoteRevisionStoreObject? item, uint propertyId) {
        return FindProperty(item?.PropertySet, propertyId)?.ReferencedIds ?? Array.Empty<OneNoteExtendedGuid>();
    }

    internal static string? ReadString(OneNoteRevisionStoreObject? item, uint propertyId) {
        byte[]? data = ReadData(item, propertyId);
        if (data == null || data.Length == 0) return null;
        string value = System.Text.Encoding.Unicode.GetString(data, 0, data.Length - data.Length % 2);
        return value.TrimEnd('\0');
    }

    private static string? ReadSingleByteString(OneNoteRevisionStoreObject? item, uint propertyId) {
        byte[]? data = ReadData(item, propertyId);
        if (data == null || data.Length == 0) return null;
        var characters = new char[data.Length];
        for (int index = 0; index < data.Length; index++) characters[index] = (char)data[index];
        return new string(characters).TrimEnd('\0');
    }

    internal static byte[]? ReadData(OneNoteRevisionStoreObject? item, uint propertyId) {
        OneNoteBinaryPayload? payload = FindProperty(item?.PropertySet, propertyId)?.Data;
        return payload?.ToArray(int.MaxValue);
    }

    internal static bool? ReadBoolean(OneNoteRevisionStoreObject? item, uint propertyId) {
        return FindProperty(item?.PropertySet, propertyId)?.BooleanValue;
    }

    internal static uint? ReadUInt32(OneNoteRevisionStoreObject? item, uint propertyId) {
        ulong? value = FindProperty(item?.PropertySet, propertyId)?.ScalarValue;
        return value.HasValue ? (uint)value.Value : null;
    }

    private static ushort? ReadUInt16(OneNoteRevisionStoreObject? item, uint propertyId) {
        ulong? value = FindProperty(item?.PropertySet, propertyId)?.ScalarValue;
        return value.HasValue ? (ushort)value.Value : null;
    }

    private static double? ReadFloat(OneNoteRevisionStoreObject? item, uint propertyId) {
        byte[]? data = ReadData(item, propertyId);
        return data != null && data.Length == 4 ? BitConverter.ToSingle(data, 0) : (double?)null;
    }

    private static IReadOnlyList<uint> ReadUInt32Array(OneNoteRevisionStoreObject? item, uint propertyId) {
        byte[]? data = ReadData(item, propertyId);
        if (data == null || data.Length % 4 != 0) return Array.Empty<uint>();
        var values = new uint[data.Length / 4];
        for (int index = 0; index < values.Length; index++) values[index] = OneNoteBinary.ReadUInt32(data, index * 4);
        return values;
    }

    private static DateTime? ReadFileTime(OneNoteRevisionStoreObject? item, uint propertyId) {
        ulong? value = FindProperty(item?.PropertySet, propertyId)?.ScalarValue;
        if (!value.HasValue || value.Value == 0 || value.Value > long.MaxValue) return null;
        try { return DateTime.FromFileTimeUtc((long)value.Value); } catch (ArgumentOutOfRangeException) { return null; }
    }

    internal static OneNotePropertyValue? FindProperty(OneNotePropertySet? set, uint propertyId) {
        uint normalized = propertyId & 0x7FFFFFFFU;
        return set?.Properties.LastOrDefault(property => (property.RawPropertyId & 0x7FFFFFFFU) == normalized);
    }

    private static DateTime? ReadTime32(OneNoteRevisionStoreObject? item, uint propertyId) {
        uint? value = ReadUInt32(item, propertyId);
        return value.HasValue ? new DateTime(1980, 1, 1, 0, 0, 0, DateTimeKind.Utc).AddSeconds(value.Value) : (DateTime?)null;
    }

}

using System.Text;

namespace OfficeIMO.OneNote;

internal sealed partial class OneNoteWriteGraphBuilder {
    private const double DefaultBodyOutlineHorizontalOffset = 1.0;
    private const double DefaultBodyOutlineVerticalOffset = 2.4;

    /// <summary>
    /// Normalizes page-level convenience content into the stable root outline required for Microsoft
    /// OneNote to render it. Direct outline-element children are preserved by OneNote but remain invisible,
    /// and the canonical MS-ONE body placement prevents the new outline from colliding with the title.
    /// </summary>
    /// <param name="page">Page whose direct content is moved into a native root outline.</param>
    private static void NormalizeDirectContent(OneNotePage page) {
        OneNoteElement[] movable = page.DirectContent
            .Where(element => !(element is OneNoteImage image && image.IsBackground == true))
            .ToArray();
        if (movable.Length == 0) return;

        var outline = new OneNoteOutline {
            Layout = new OneNoteLayout {
                X = DefaultBodyOutlineHorizontalOffset,
                Y = DefaultBodyOutlineVerticalOffset
            }
        };
        foreach (OneNoteElement child in movable) {
            outline.Children.Add(child);
            page.DirectContent.Remove(child);
        }
        page.Outlines.Add(outline);
    }

    private OneNoteExtendedGuid BuildOutline(OneNoteWriteObjectSpace space, OneNoteOutline outline, uint lastModifiedTime) {
        if (outline.IsOutlineElementWrapper) {
            return BuildOutlineElementWrapper(space, outline, lastModifiedTime);
        }
        EnsureTagTargetSupported(outline);
        var childIds = new List<OneNoteExtendedGuid>();
        foreach (OneNoteElement child in outline.Children) childIds.Add(BuildOutlineChild(space, child, lastModifiedTime));
        OneNoteExtendedGuid id = IdOrNew(outline.Id);
        outline.Id = id;
        var properties = LayoutProperties(outline.Layout);
        properties.Insert(0, Scalar(OneNoteSchema.LastModifiedTime, lastModifiedTime));
        if (childIds.Count > 0) properties.Add(ObjectReferences(OneNoteSchema.ElementChildNodes, childIds));
        space.Objects.Add(new OneNoteWriteObject(id, OneNoteSchema.JcidOutlineNode, properties));
        return id;
    }

    private OneNoteExtendedGuid BuildOutlineElementWrapper(
        OneNoteWriteObjectSpace space,
        OneNoteOutline wrapper,
        uint lastModifiedTime) {
        if (wrapper.Children.Count == 0) {
            throw new OneNoteFormatException("ONENOTE_WRITE_OUTLINE_ELEMENT_CONTENT", "A preserved outline-element wrapper must contain primary content.");
        }

        OneNoteExtendedGuid primaryId = BuildOutlineChild(space, wrapper.Children[0], lastModifiedTime);
        var nestedIds = new List<OneNoteExtendedGuid>();
        foreach (OneNoteElement child in wrapper.Children.Skip(1)) {
            nestedIds.Add(BuildOutlineChild(space, child, lastModifiedTime));
        }

        var properties = LayoutProperties(wrapper.Layout);
        properties.Insert(0, Scalar(OneNoteSchema.LastModifiedTime, lastModifiedTime));
        properties.Add(ObjectReferences(OneNoteSchema.ContentChildNodes, primaryId));
        if (nestedIds.Count > 0) properties.Add(ObjectReferences(OneNoteSchema.ElementChildNodes, nestedIds));
        if (wrapper.WrapperList != null) {
            properties.Add(ObjectReferences(OneNoteSchema.ListNodes, BuildList(space, wrapper.WrapperList, lastModifiedTime)));
            properties.Add(Scalar(
                OneNoteSchema.OutlineElementChildLevel,
                (ulong)Math.Min(byte.MaxValue, Math.Max(1, wrapper.WrapperList.Level + 1))));
        }
        if (!string.IsNullOrWhiteSpace(wrapper.Author?.Name)) {
            properties.Add(ObjectReferences(OneNoteSchema.AuthorMostRecent, BuildAuthor(space, wrapper.Author!)));
        }
        AddTags(space, properties, wrapper.Tags);

        OneNoteExtendedGuid id = IdOrNew(wrapper.Id);
        wrapper.Id = id;
        space.Objects.Add(new OneNoteWriteObject(id, OneNoteSchema.JcidOutlineElementNode, properties));
        return id;
    }

    private OneNoteExtendedGuid BuildOutlineChild(OneNoteWriteObjectSpace space, OneNoteElement element, uint lastModifiedTime) {
        if (element is OneNoteOutline outline) return BuildOutline(space, outline, lastModifiedTime);
        if (element is OneNoteParagraph paragraph) return BuildParagraph(space, paragraph, lastModifiedTime);
        if (element is OneNoteTable table) return BuildTable(space, table, lastModifiedTime);
        if (element is OneNoteImage image) return BuildImage(space, image, lastModifiedTime);
        if (element is OneNoteMedia media) return BuildEmbeddedFile(space, media, lastModifiedTime);
        if (element is OneNoteEmbeddedFile embedded) return BuildEmbeddedFile(space, embedded, lastModifiedTime);
        if (element is OneNoteMath math) return BuildMath(space, math, lastModifiedTime);
        if (element is OneNoteInk ink) return BuildInk(space, ink, lastModifiedTime);
        throw new OneNoteFormatException("ONENOTE_WRITE_UNSUPPORTED_CONTENT", "The initial OneNote writer cannot yet serialize " + element.Kind + " content.");
    }

    private OneNoteExtendedGuid BuildParagraph(OneNoteWriteObjectSpace space, OneNoteParagraph paragraph, uint lastModifiedTime) {
        IReadOnlyList<OneNoteTextRun> serializedRuns = ExpandMathRuns(paragraph.Runs);
        string text = string.Concat(serializedRuns.Select(run => run.Text ?? string.Empty));
        uint languageId = serializedRuns
            .Where(run => run.Style.LanguageId.HasValue)
            .Select(run => run.Style.LanguageId!.Value)
            .DefaultIfEmpty(0x0409U)
            .First();
        var richTextProperties = new List<OneNoteWriteProperty> {
            Scalar(OneNoteSchema.LastModifiedTime, lastModifiedTime),
            Scalar(OneNoteSchema.RichEditTextLanguageId, Math.Min(ushort.MaxValue, languageId)),
            Data(OneNoteSchema.RichEditTextUnicode, Unicode(text))
        };
        bool writeRunFormatting = serializedRuns.Count > 1 ||
            (serializedRuns.Count == 1 && HasExplicitTextStyle(serializedRuns[0]));
        if (writeRunFormatting) {
            var styleIds = new List<OneNoteExtendedGuid>();
            var boundaries = new List<uint>();
            int length = 0;
            for (int index = 0; index < serializedRuns.Count; index++) {
                OneNoteTextRun run = serializedRuns[index];
                length += (run.Text ?? string.Empty).Length;
                if (index < serializedRuns.Count - 1) boundaries.Add((uint)length);
                styleIds.Add(BuildTextStyle(space, run));
            }
            if (boundaries.Count > 0) richTextProperties.Add(Data(OneNoteSchema.TextRunIndex, UInt32Array(boundaries)));
            richTextProperties.Add(ObjectReferences(OneNoteSchema.TextRunFormatting, styleIds));
        }
        if (serializedRuns.Any(run => run.MathDescriptor != null)) {
            var childSets = new List<IReadOnlyList<OneNoteWriteProperty>>(serializedRuns.Count);
            foreach (OneNoteTextRun run in serializedRuns) childSets.Add(MathDescriptorProperties(run.MathDescriptor));
            richTextProperties.Add(new OneNoteWriteProperty(
                OneNoteSchema.MathInlineObjects,
                childPropertySets: childSets,
                childPropertyId: OneNoteSchema.MathInlineObjectElement));
        }
        OneNoteExtendedGuid richTextId = IdOrNew(paragraph.ContentObjectId);
        paragraph.ContentObjectId = richTextId;
        OneNoteExtendedGuid? paragraphStyleId = BuildParagraphStyle(space, paragraph.Style);
        if (paragraphStyleId != null) richTextProperties.Add(ObjectReferences(OneNoteSchema.ParagraphStyle, paragraphStyleId));
        AddTags(space, richTextProperties, paragraph.Tags);
        space.Objects.Add(new OneNoteWriteObject(richTextId, OneNoteSchema.JcidRichTextNode, richTextProperties));

        var nested = new List<OneNoteExtendedGuid>();
        foreach (OneNoteElement child in paragraph.Children) nested.Add(BuildOutlineChild(space, child, lastModifiedTime));
        OneNoteExtendedGuid elementId = IsCompact(paragraph.Id) && !paragraph.Id!.Equals(richTextId) ? paragraph.Id! : _ids.New();
        paragraph.Id = elementId;
        var properties = LayoutProperties(paragraph.Layout);
        properties.Insert(0, Scalar(OneNoteSchema.LastModifiedTime, lastModifiedTime));
        properties.Add(ObjectReferences(OneNoteSchema.ContentChildNodes, richTextId));
        if (nested.Count > 0) properties.Add(ObjectReferences(OneNoteSchema.ElementChildNodes, nested));
        if (paragraph.List != null) {
            properties.Add(ObjectReferences(OneNoteSchema.ListNodes, BuildList(space, paragraph.List, lastModifiedTime)));
            properties.Add(Scalar(OneNoteSchema.OutlineElementChildLevel, (ulong)Math.Min(byte.MaxValue, Math.Max(1, paragraph.List.Level + 1))));
        }
        if (!string.IsNullOrWhiteSpace(paragraph.Author?.Name)) {
            properties.Add(ObjectReferences(OneNoteSchema.AuthorMostRecent, BuildAuthor(space, paragraph.Author!)));
        }
        space.Objects.Add(new OneNoteWriteObject(elementId, OneNoteSchema.JcidOutlineElementNode, properties));
        return elementId;
    }

    private OneNoteExtendedGuid BuildTable(OneNoteWriteObjectSpace space, OneNoteTable table, uint lastModifiedTime) {
        var rowIds = new List<OneNoteExtendedGuid>();
        foreach (OneNoteTableRow row in table.Rows) {
            var cellIds = new List<OneNoteExtendedGuid>();
            foreach (OneNoteTableCell cell in row.Cells) {
                var contentIds = new List<OneNoteExtendedGuid>();
                foreach (OneNoteElement content in cell.Content) contentIds.Add(BuildOutlineChild(space, content, lastModifiedTime));
                var cellProperties = new List<OneNoteWriteProperty> { Scalar(OneNoteSchema.LastModifiedTime, lastModifiedTime) };
                if (contentIds.Count > 0) cellProperties.Add(ObjectReferences(OneNoteSchema.ElementChildNodes, contentIds));
                if (cell.ShadingColorArgb.HasValue) cellProperties.Add(Scalar(OneNoteSchema.CellShadingColor, cell.ShadingColorArgb.Value));
                OneNoteExtendedGuid cellId = IdOrNew(cell.ObjectId);
                cell.ObjectId = cellId;
                space.Objects.Add(new OneNoteWriteObject(cellId, OneNoteSchema.JcidTableCellNode, cellProperties));
                cellIds.Add(cellId);
            }
            OneNoteExtendedGuid rowId = IdOrNew(row.ObjectId);
            row.ObjectId = rowId;
            var rowProperties = new List<OneNoteWriteProperty> { Scalar(OneNoteSchema.LastModifiedTime, lastModifiedTime) };
            if (cellIds.Count > 0) rowProperties.Add(ObjectReferences(OneNoteSchema.ElementChildNodes, cellIds));
            space.Objects.Add(new OneNoteWriteObject(rowId, OneNoteSchema.JcidTableRowNode, rowProperties));
            rowIds.Add(rowId);
        }
        var properties = LayoutProperties(table.Layout);
        properties.Insert(0, Scalar(OneNoteSchema.LastModifiedTime, lastModifiedTime));
        properties.Add(Boolean(OneNoteSchema.TableBordersVisible, table.BordersVisible));
        properties.Add(Data(OneNoteSchema.TableColumnWidths, TableWidths(table.ColumnWidths)));
        if (rowIds.Count > 0) properties.Add(ObjectReferences(OneNoteSchema.ElementChildNodes, rowIds));
        AddTags(space, properties, table.Tags);
        OneNoteExtendedGuid id = IdOrNew(table.Id);
        table.Id = id;
        space.Objects.Add(new OneNoteWriteObject(id, OneNoteSchema.JcidTableNode, properties));
        return id;
    }

    private OneNoteExtendedGuid BuildImage(OneNoteWriteObjectSpace space, OneNoteImage image, uint lastModifiedTime) {
        var properties = LayoutProperties(image.Layout);
        properties.Insert(0, Scalar(OneNoteSchema.LastModifiedTime, lastModifiedTime));
        AddImagePayloadReferences(space, image, properties);
        AddString(properties, OneNoteSchema.ImageFilename, image.FileName);
        AddString(properties, OneNoteSchema.ImageAltText, image.AltText);
        AddString(properties, OneNoteSchema.SourceFilePath, image.SourcePath);
        AddString(properties, OneNoteSchema.HyperlinkUrl, image.Hyperlink);
        AddString(properties, OneNoteSchema.RichEditTextUnicode, image.OcrText);
        if (image.OcrLanguageId.HasValue) properties.Add(Scalar(OneNoteSchema.LanguageId, Math.Min(ushort.MaxValue, image.OcrLanguageId.Value)));
        if (image.DisplayedPageNumber.HasValue) properties.Add(Scalar(OneNoteSchema.DisplayedPageNumber, image.DisplayedPageNumber.Value));
        AddBoolean(properties, OneNoteSchema.IsBackground, image.IsBackground);
        AddBoolean(properties, OneNoteSchema.IsLayoutSizeSetByUser, image.SizeSetByUser);
        if (image.UploadState.HasValue) properties.Add(Scalar(OneNoteSchema.ImageUploadState, image.UploadState.Value));
        if (image.WidthHalfInches.HasValue) properties.Add(Float(OneNoteSchema.PictureWidth, image.WidthHalfInches.Value));
        if (image.HeightHalfInches.HasValue) properties.Add(Float(OneNoteSchema.PictureHeight, image.HeightHalfInches.Value));
        AddTags(space, properties, image.Tags);
        OneNoteExtendedGuid id = IdOrNew(image.Id);
        image.Id = id;
        space.Objects.Add(new OneNoteWriteObject(id, OneNoteSchema.JcidImageNode, properties));
        return id;
    }

    private void AddImagePayloadReferences(
        OneNoteWriteObjectSpace space,
        OneNoteImage image,
        IList<OneNoteWriteProperty> properties) {
        bool hasPreservedReferences = image.PictureContainerObjectId != null || image.WebPictureContainerObjectId != null;
        if (!_preserveUnknownData || !hasPreservedReferences) {
            OneNoteExtendedGuid canonicalId = BuildFileData(space, image, true);
            properties.Add(ObjectReferences(OneNoteSchema.PictureContainer, canonicalId));
            return;
        }

        OneNoteExtendedGuid? selectedId = null;
        if (image.Payload != null) selectedId = BuildFileData(space, image, true);
        if (image.Payload == null && image.PictureContainerObjectId == null && image.WebPictureContainerObjectId == null) {
            throw new OneNoteFormatException("ONENOTE_WRITE_MISSING_PAYLOAD", image.Kind + " content has no binary payload.");
        }

        OneNoteExtendedGuid? pictureId = image.PayloadUsesWebPictureContainer
            ? image.PictureContainerObjectId
            : selectedId ?? image.PictureContainerObjectId;
        OneNoteExtendedGuid? webPictureId = image.PayloadUsesWebPictureContainer
            ? selectedId ?? image.WebPictureContainerObjectId
            : image.WebPictureContainerObjectId;
        if (pictureId != null) properties.Add(ObjectReferences(OneNoteSchema.PictureContainer, pictureId));
        if (webPictureId != null) properties.Add(ObjectReferences(OneNoteSchema.WebPictureContainer14, webPictureId));
    }

    private OneNoteExtendedGuid BuildEmbeddedFile(OneNoteWriteObjectSpace space, OneNoteBinaryElement element, uint lastModifiedTime) {
        OneNoteExtendedGuid binaryId = BuildFileData(space, element, false);
        var properties = LayoutProperties(element.Layout);
        properties.Insert(0, Scalar(OneNoteSchema.LastModifiedTime, lastModifiedTime));
        properties.Add(ObjectReferences(OneNoteSchema.EmbeddedFileContainer, binaryId));
        AddString(properties, OneNoteSchema.EmbeddedFileName, element.FileName);
        string? sourcePath = element is OneNoteMedia media ? media.SourcePath : ((OneNoteEmbeddedFile)element).SourcePath;
        AddString(properties, OneNoteSchema.SourceFilePath, sourcePath);
        if (element is OneNoteMedia recording) {
            OneNoteMediaKind kind = ResolveRecordingKind(recording);
            if (kind != OneNoteMediaKind.Unknown) {
                recording.RecordingKind = kind;
                recording.RecordingId ??= Guid.NewGuid();
                properties.Add(Scalar(OneNoteSchema.RecordMedia, kind == OneNoteMediaKind.Audio ? 1U : 2U));
                properties.Add(Data(OneNoteSchema.AudioRecordingGuid, recording.RecordingId.Value.ToByteArray()));
                if (recording.Duration.HasValue) {
                    properties.Add(Scalar(
                        OneNoteSchema.AudioRecordingDuration,
                        checked((uint)Math.Round(recording.Duration.Value.TotalMilliseconds, MidpointRounding.AwayFromZero))));
                }
            }
        }
        AddTags(space, properties, element.Tags);
        OneNoteExtendedGuid id = IdOrNew(element.Id);
        element.Id = id;
        space.Objects.Add(new OneNoteWriteObject(id, OneNoteSchema.JcidEmbeddedFileNode, properties));
        return id;
    }

    private OneNoteExtendedGuid BuildFileData(OneNoteWriteObjectSpace space, OneNoteBinaryElement element, bool picture) {
        if (element.Payload == null) throw new OneNoteFormatException("ONENOTE_WRITE_MISSING_PAYLOAD", element.Kind + " content has no binary payload.");
        byte[] payload = element.Payload.ToArray(_maxPayloadBytes);
        Guid dataId = element.PayloadFileDataId ?? Guid.NewGuid();
        string extension = element.PayloadFileExtension ?? Path.GetExtension(element.FileName ?? string.Empty);
        OneNoteWriteProperty[] properties = {
            Data(OneNoteSchema.FileDataReference, dataId.ToByteArray()),
            Data(OneNoteSchema.FileDataExtension, Unicode(extension))
        };
        OneNoteExtendedGuid id = IdOrNew(element.PayloadObjectId);
        uint jcid = picture ? OneNoteSchema.JcidPictureData : OneNoteSchema.JcidEmbeddedFileData;
        OneNoteWriteObject? existingById = space.Objects.FirstOrDefault(item => item.Id.Equals(id));
        OneNoteWriteObject? existingByFileData = space.Objects.FirstOrDefault(item => item.FileDataId == dataId);
        OneNoteWriteObject? reusable = existingById ?? existingByFileData;
        if (reusable != null &&
            reusable.Jcid == jcid &&
            PropertiesEqual(reusable.Properties, properties) &&
            BytesEqual(reusable.Blob, payload) &&
            reusable.FileDataId == dataId &&
            string.Equals(reusable.FileExtension, extension, StringComparison.Ordinal)) {
            element.PayloadObjectId = reusable.Id;
            element.PayloadFileDataId = dataId;
            element.PayloadFileExtension = extension;
            return reusable.Id;
        }
        if (existingById != null || existingByFileData != null) {
            id = _ids.New();
            dataId = Guid.NewGuid();
            properties = new[] {
                Data(OneNoteSchema.FileDataReference, dataId.ToByteArray()),
                Data(OneNoteSchema.FileDataExtension, Unicode(extension))
            };
        }
        element.PayloadObjectId = id;
        element.PayloadFileDataId = dataId;
        element.PayloadFileExtension = extension;
        space.Objects.Add(new OneNoteWriteObject(
            id,
            jcid,
            properties,
            payload,
            dataId,
            extension));
        return id;
    }

    private OneNoteExtendedGuid BuildMath(OneNoteWriteObjectSpace space, OneNoteMath math, uint lastModifiedTime) {
        if (math.RawPayload != null) throw new OneNoteFormatException("ONENOTE_WRITE_UNSUPPORTED_MATH_PAYLOAD", "An opaque mathematical payload cannot be regenerated without a structured expression.");
        OfficeIMO.Drawing.OfficeMathExpression expression = OneNoteMathNativeCodec.Canonicalize(math.GetExpression());
        math.Expression = expression;
        math.Text = expression.ToPlainText();
        math.MathMl = OfficeIMO.Drawing.OfficeMathMarkup.ToMathMl(expression);
        math.Latex = OfficeIMO.Drawing.OfficeMathMarkup.ToLatex(expression);
        var paragraph = new OneNoteParagraph {
            Id = math.Id,
            ContentObjectId = math.ContentObjectId,
            Layout = math.Layout,
            Author = math.Author
        };
        foreach (OneNoteTag tag in math.Tags) paragraph.Tags.Add(tag);
        paragraph.AddMath(expression);
        OneNoteExtendedGuid id = BuildParagraph(space, paragraph, lastModifiedTime);
        math.Id = id;
        math.ContentObjectId = paragraph.ContentObjectId;
        return id;
    }

    private static OneNoteMediaKind ResolveRecordingKind(OneNoteMedia recording) {
        if (recording.RecordingKind != OneNoteMediaKind.Unknown) return recording.RecordingKind;
        switch (Path.GetExtension(recording.FileName ?? recording.SourcePath ?? string.Empty).ToLowerInvariant()) {
            case ".wma":
            case ".mp3":
            case ".wav": return OneNoteMediaKind.Audio;
            case ".wmv":
            case ".avi":
            case ".mpg": return OneNoteMediaKind.Video;
            default: return OneNoteMediaKind.Unknown;
        }
    }

    internal static IReadOnlyList<OneNoteTextRun> ExpandMathRuns(IList<OneNoteTextRun> semanticRuns) {
        if (!semanticRuns.Any(run => run.MathExpression != null)) return semanticRuns.ToArray();
        var output = new List<OneNoteTextRun>();
        foreach (OneNoteTextRun semantic in semanticRuns) {
            if (semantic.MathExpression == null) {
                output.Add(semantic);
                continue;
            }
            if (OneNoteMathRunPreservation.CanReuse(semantic)) {
                foreach (OneNoteTextRun native in OneNoteMathRunPreservation.CloneForWrite(semantic)) output.Add(native);
                continue;
            }
            semantic.MathExpression = OneNoteMathNativeCodec.Canonicalize(semantic.MathExpression);
            semantic.Text = semantic.MathExpression.ToPlainText();
            bool first = true;
            foreach (OneNoteMathNativeCodec.EncodedRun encoded in OneNoteMathNativeCodec.Encode(semantic.MathExpression)) {
                var native = new OneNoteTextRun {
                    Text = encoded.Text,
                    MathDescriptor = encoded.Descriptor,
                    Hyperlink = semantic.Hyperlink,
                    HyperlinkProtected = semantic.HyperlinkProtected,
                    StyleObjectId = semantic.StyleObjectId
                };
                CopyTextStyleForWrite(semantic.Style, native.Style);
                native.Style.IsMath = true;
                if (first) {
                    foreach (OneNoteOpaqueProperty property in semantic.UnknownProperties) native.UnknownProperties.Add(property);
                    first = false;
                }
                output.Add(native);
            }
        }
        return output;
    }

    private static void CopyTextStyleForWrite(OneNoteTextStyle source, OneNoteTextStyle destination) {
        destination.FontFamily = source.FontFamily;
        destination.FontSize = source.FontSize;
        destination.ColorArgb = source.ColorArgb;
        destination.HighlightColorArgb = source.HighlightColorArgb;
        destination.Bold = source.Bold;
        destination.Italic = source.Italic;
        destination.Underline = source.Underline;
        destination.Strikethrough = source.Strikethrough;
        destination.Superscript = source.Superscript;
        destination.Subscript = source.Subscript;
        destination.LanguageId = source.LanguageId;
        destination.IsMath = source.IsMath;
    }

    private static IReadOnlyList<OneNoteWriteProperty> MathDescriptorProperties(OneNoteMathInlineDescriptor? descriptor) {
        if (descriptor == null) return Array.Empty<OneNoteWriteProperty>();
        var properties = new List<OneNoteWriteProperty> { Scalar(OneNoteSchema.MathInlineObjectType, descriptor.Type) };
        if (descriptor.Count.HasValue) properties.Add(Scalar(OneNoteSchema.MathInlineObjectCount, descriptor.Count.Value));
        if (descriptor.Column.HasValue) properties.Add(Scalar(OneNoteSchema.MathInlineObjectColumn, descriptor.Column.Value));
        if (descriptor.Alignment.HasValue) properties.Add(Scalar(OneNoteSchema.MathInlineObjectAlignment, descriptor.Alignment.Value));
        if (descriptor.Character.HasValue) properties.Add(Scalar(OneNoteSchema.MathInlineObjectCharacter, descriptor.Character.Value));
        if (descriptor.Character1.HasValue) properties.Add(Scalar(OneNoteSchema.MathInlineObjectCharacter1, descriptor.Character1.Value));
        if (descriptor.Character2.HasValue) properties.Add(Scalar(OneNoteSchema.MathInlineObjectCharacter2, descriptor.Character2.Value));
        return properties;
    }

    private OneNoteExtendedGuid BuildList(OneNoteWriteObjectSpace space, OneNoteListInfo list, uint lastModifiedTime) {
        string format = list.Ordered ? "\uFFFD" + (char)(list.Format ?? 0) : "\u2022";
        string encoded = new string((char)format.Length, 1) + format;
        var properties = new List<OneNoteWriteProperty> {
            Scalar(OneNoteSchema.LastModifiedTime, lastModifiedTime),
            Data(OneNoteSchema.NumberListFormat, Unicode(encoded))
        };
        AddString(properties, OneNoteSchema.ListFont, list.FontFamily);
        if (list.Restart || list.DisplayIndex.HasValue) properties.Add(Scalar(OneNoteSchema.ListRestart, (uint)Math.Max(1, list.DisplayIndex ?? 1)));
        OneNoteExtendedGuid id = IdOrNew(list.ObjectId);
        OneNoteWriteObject? existing = space.Objects.FirstOrDefault(item => item.Id.Equals(id));
        if (existing != null) {
            if (existing.Jcid == OneNoteSchema.JcidNumberListNode && PropertiesEqual(existing.Properties, properties)) {
                list.ObjectId = id;
                return id;
            }
            id = _ids.New();
        }
        list.ObjectId = id;
        space.Objects.Add(new OneNoteWriteObject(id, OneNoteSchema.JcidNumberListNode, properties));
        return id;
    }

    private OneNoteExtendedGuid? BuildParagraphStyle(OneNoteWriteObjectSpace space, OneNoteParagraphStyle style) {
        var properties = new List<OneNoteWriteProperty>();
        AddString(properties, OneNoteSchema.ParagraphStyleId, style.StyleId);
        if (style.Alignment.HasValue) properties.Add(Scalar(OneNoteSchema.ParagraphAlignment, (uint)style.Alignment.Value));
        if (style.SpaceBefore.HasValue) properties.Add(Float(OneNoteSchema.ParagraphSpaceBefore, style.SpaceBefore.Value));
        if (style.SpaceAfter.HasValue) properties.Add(Float(OneNoteSchema.ParagraphSpaceAfter, style.SpaceAfter.Value));
        if (style.ExactLineSpacing.HasValue) properties.Add(Float(OneNoteSchema.ParagraphLineSpacingExact, style.ExactLineSpacing.Value));
        if (properties.Count == 0) return null;
        OneNoteExtendedGuid id = IdOrNew(style.ObjectId);
        OneNoteWriteObject? existing = space.Objects.FirstOrDefault(item => item.Id.Equals(id));
        if (existing != null) {
            if (existing.Jcid == OneNoteSchema.JcidTextStyle && PropertiesEqual(existing.Properties, properties)) {
                style.ObjectId = id;
                return id;
            }
            id = _ids.New();
        }
        style.ObjectId = id;
        space.Objects.Add(new OneNoteWriteObject(id, OneNoteSchema.JcidTextStyle, properties));
        return id;
    }

    private OneNoteExtendedGuid BuildAuthor(OneNoteWriteObjectSpace space, OneNoteAuthor author) {
        var properties = new[] { Data(OneNoteSchema.Author, Unicode(author.Name!)) };
        OneNoteExtendedGuid id = IdOrNew(author.ObjectId);
        OneNoteWriteObject? existing = space.Objects.FirstOrDefault(item => item.Id.Equals(id));
        if (existing != null) {
            if (existing.Jcid == OneNoteSchema.JcidAuthor && PropertiesEqual(existing.Properties, properties)) {
                author.ObjectId = id;
                return id;
            }
            id = _ids.New();
        }
        author.ObjectId = id;
        space.Objects.Add(new OneNoteWriteObject(id, OneNoteSchema.JcidAuthor, properties));
        return id;
    }

    private OneNoteExtendedGuid BuildTextStyle(OneNoteWriteObjectSpace space, OneNoteTextRun run) {
        var properties = new List<OneNoteWriteProperty>();
        AddBoolean(properties, OneNoteSchema.Bold, run.Style.Bold);
        AddBoolean(properties, OneNoteSchema.Italic, run.Style.Italic);
        AddBoolean(properties, OneNoteSchema.Underline, run.Style.Underline);
        AddBoolean(properties, OneNoteSchema.Strikethrough, run.Style.Strikethrough);
        AddBoolean(properties, OneNoteSchema.Superscript, run.Style.Superscript);
        AddBoolean(properties, OneNoteSchema.Subscript, run.Style.Subscript);
        properties.Add(Boolean(OneNoteSchema.Hidden, false));
        properties.Add(Boolean(OneNoteSchema.MathFormatting, run.Style.IsMath ?? false));
        if (!string.IsNullOrWhiteSpace(run.Style.FontFamily)) properties.Add(Data(OneNoteSchema.Font, Unicode(run.Style.FontFamily!)));
        if (run.Style.FontSize.HasValue) properties.Add(Scalar(OneNoteSchema.FontSize, FontSizeHalfPoints(run.Style.FontSize.Value)));
        if (run.Style.ColorArgb.HasValue) properties.Add(Scalar(OneNoteSchema.FontColor, run.Style.ColorArgb.Value));
        if (run.Style.HighlightColorArgb.HasValue) properties.Add(Scalar(OneNoteSchema.Highlight, run.Style.HighlightColorArgb.Value));
        properties.Add(Scalar(OneNoteSchema.Charset, 0));
        properties.Add(Scalar(OneNoteSchema.LanguageId, run.Style.LanguageId ?? 0x0409U));
        bool hasHyperlink = !string.IsNullOrWhiteSpace(run.Hyperlink);
        properties.Add(Boolean(OneNoteSchema.Hyperlink, hasHyperlink));
        properties.Add(Boolean(OneNoteSchema.HyperlinkProtected, hasHyperlink && run.HyperlinkProtected));
        if (hasHyperlink) {
            properties.Add(Data(OneNoteSchema.HyperlinkUrl, Unicode(run.Hyperlink!)));
        }
        OneNoteExtendedGuid id = IdOrNew(run.StyleObjectId);
        OneNoteWriteObject? existing = space.Objects.FirstOrDefault(item => item.Id.Equals(id));
        if (existing != null) {
            if (existing.Jcid == OneNoteSchema.JcidTextStyle && PropertiesEqual(existing.Properties, properties)) {
                run.StyleObjectId = id;
                return id;
            }
            id = _ids.New();
        }
        run.StyleObjectId = id;
        space.Objects.Add(new OneNoteWriteObject(id, OneNoteSchema.JcidTextStyle, properties));
        return id;
    }

    private static bool PropertiesEqual(
        IReadOnlyList<OneNoteWriteProperty> left,
        IReadOnlyList<OneNoteWriteProperty> right) {
        if (left.Count != right.Count) return false;
        for (int index = 0; index < left.Count; index++) {
            OneNoteWriteProperty first = left[index];
            OneNoteWriteProperty second = right[index];
            if (first.RawId != second.RawId ||
                first.Scalar != second.Scalar ||
                first.ReferenceKind != second.ReferenceKind ||
                first.ChildPropertyId != second.ChildPropertyId ||
                !BytesEqual(first.Data, second.Data) ||
                !first.References.SequenceEqual(second.References) ||
                first.ChildPropertySets.Count != second.ChildPropertySets.Count) {
                return false;
            }
        }
        return true;
    }

    private static bool BytesEqual(byte[]? left, byte[]? right) {
        if (ReferenceEquals(left, right)) return true;
        return left != null && right != null && left.SequenceEqual(right);
    }

    private void AddTags(
        OneNoteWriteObjectSpace space,
        ICollection<OneNoteWriteProperty> properties,
        IList<OneNoteTag> tags) {
        if (tags.Count == 0) return;
        if (tags.Count > 9) {
            throw new OneNoteFormatException("ONENOTE_WRITE_TAG_LIMIT", "A OneNote content object cannot contain more than nine note tags.");
        }

        var actionItemTypes = new HashSet<uint>();
        var states = new List<IReadOnlyList<OneNoteWriteProperty>>(tags.Count);
        foreach (OneNoteTag tag in tags) {
            uint actionItemType = tag.ActionItemType ?? (tag.IsTask ? 104U : 0U);
            if (tag.IsTask) {
                if (actionItemType < 100 || actionItemType > 105) {
                    throw new OneNoteFormatException("ONENOTE_WRITE_TASK_TAG_TYPE", "A task tag ActionItemType must be from 100 through 105.");
                }
            } else if (actionItemType >= 100 && actionItemType <= 105) {
                throw new OneNoteFormatException("ONENOTE_WRITE_TAG_TYPE", "ActionItemType values from 100 through 105 require IsTask to be true.");
            }
            if (!actionItemTypes.Add(actionItemType)) {
                throw new OneNoteFormatException("ONENOTE_WRITE_DUPLICATE_TAG_TYPE", "Note-tag ActionItemType values must be unique on a content object.");
            }

            var state = new List<OneNoteWriteProperty>();
            if (tag.IsTask) {
                if (!tag.IsCheckable) {
                    throw new OneNoteFormatException("ONENOTE_WRITE_TASK_TAG_CHECKABILITY", "MS-ONE task tags are always checkable.");
                }
                uint shape = tag.Shape ?? 89U;
                if (shape < 89 || shape > 93) {
                    throw new OneNoteFormatException("ONENOTE_WRITE_TASK_TAG_SHAPE", "A task tag shape must be from 89 through 93.");
                }
                state.Add(Scalar(OneNoteSchema.ActionItemType, actionItemType));
                state.Add(Scalar(OneNoteSchema.NoteTagPropertyStatus, 0));
                state.Add(Scalar(OneNoteSchema.NoteTagShape, shape));
                state.Add(Scalar(OneNoteSchema.ActionItemSchemaVersion, 0));
                state.Add(Scalar(OneNoteSchema.TaskTagDueDate, tag.DueUtc.HasValue ? Time32(tag.DueUtc.Value) : 0U));
            } else {
                OneNoteExtendedGuid definitionId = BuildTagDefinition(space, tag, actionItemType);
                state.Add(ObjectReferences(OneNoteSchema.NoteTagDefinitionOid, definitionId));
            }
            if (tag.CreatedUtc.HasValue) state.Add(Scalar(OneNoteSchema.NoteTagCreated, Time32(tag.CreatedUtc.Value)));
            DateTime? completedUtc = tag.IsCheckable ? tag.CompletedUtc : tag.CreatedUtc;
            if (completedUtc.HasValue) state.Add(Scalar(OneNoteSchema.NoteTagCompleted, Time32(completedUtc.Value)));
            else state.Add(Scalar(OneNoteSchema.NoteTagCompleted, 0));
            uint status = (tag.IsCompleted || !tag.IsCheckable ? 0x01U : 0U) |
                (tag.IsDisabled ? 0x02U : 0U) |
                (tag.IsTask ? 0x04U : 0U) |
                (tag.IsUnsynchronized ? 0x08U : 0U) |
                (tag.IsRemoved ? 0x10U : 0U);
            state.Add(Scalar(OneNoteSchema.ActionItemStatus, status));
            states.Add(state);
        }

        properties.Add(new OneNoteWriteProperty(
            OneNoteSchema.NoteTagStates,
            childPropertySets: states,
            childPropertyId: ((uint)OneNotePropertyType.PropertySet << 26),
            preserveRawId: true));
    }

    private OneNoteExtendedGuid BuildTagDefinition(OneNoteWriteObjectSpace space, OneNoteTag tag, uint actionItemType) {
        OneNoteExtendedGuid id = IdOrNew(tag.DefinitionId);
        uint shape = tag.Shape ?? (tag.IsCheckable ? 3U : 13U);
        if (shape > 143) {
            throw new OneNoteFormatException("ONENOTE_WRITE_TAG_SHAPE", "A normal note-tag shape must be from 0 through 143.");
        }
        if (OneNoteSemanticMapper.IsCheckableTagShape(shape) != tag.IsCheckable) {
            throw new OneNoteFormatException("ONENOTE_WRITE_TAG_CHECKABILITY", "The note-tag shape and IsCheckable value must describe the same MS-ONE tag behavior.");
        }
        var properties = new List<OneNoteWriteProperty> {
            Scalar(OneNoteSchema.ActionItemSchemaVersion, 0),
            Scalar(OneNoteSchema.ActionItemType, actionItemType),
            Scalar(OneNoteSchema.NoteTagShape, shape),
            Scalar(OneNoteSchema.NoteTagPropertyStatus, 9)
        };
        AddString(properties, OneNoteSchema.NoteTagLabel, tag.Label ?? "Tag");
        if (tag.HighlightColorArgb.HasValue) properties.Add(Scalar(OneNoteSchema.NoteTagHighlightColor, tag.HighlightColorArgb.Value));
        if (tag.TextColorArgb.HasValue) properties.Add(Scalar(OneNoteSchema.NoteTagTextColor, tag.TextColorArgb.Value));
        OneNoteWriteObject? existing = space.Objects.FirstOrDefault(item => item.Id.Equals(id));
        if (existing != null) {
            if (existing.Jcid == OneNoteSchema.JcidNoteTagSharedDefinition && PropertiesEqual(existing.Properties, properties)) {
                tag.DefinitionId = id;
                return id;
            }
            id = _ids.New();
        }
        tag.DefinitionId = id;
        space.Objects.Add(new OneNoteWriteObject(id, OneNoteSchema.JcidNoteTagSharedDefinition, properties));
        return id;
    }

    private static void EnsureTagTargetSupported(OneNoteElement element) {
        if (element.Tags.Count > 0) {
            throw new OneNoteFormatException("ONENOTE_WRITE_UNSUPPORTED_TAG_TARGET", "OneNote note tags can be emitted only on paragraphs, tables, images, and embedded files.");
        }
    }

    private static List<OneNoteWriteProperty> LayoutProperties(OneNoteLayout? layout) {
        var properties = new List<OneNoteWriteProperty>();
        if (layout == null) return properties;
        if (layout.X.HasValue) properties.Add(Float(OneNoteSchema.OffsetFromParentHorizontal, layout.X.Value));
        if (layout.Y.HasValue) properties.Add(Float(OneNoteSchema.OffsetFromParentVertical, layout.Y.Value));
        if (layout.Width.HasValue) properties.Add(Float(OneNoteSchema.LayoutMaxWidth, layout.Width.Value));
        if (layout.Height.HasValue) properties.Add(Float(OneNoteSchema.LayoutMaxHeight, layout.Height.Value));
        AddBoolean(properties, OneNoteSchema.LayoutTightLayout, layout.Tight);
        AddBoolean(properties, OneNoteSchema.OutlineElementRtl, layout.RightToLeft);
        if (layout.MinimumWidth.HasValue) properties.Add(Float(OneNoteSchema.LayoutMinimumOutlineWidth, layout.MinimumWidth.Value));
        if (layout.AlignmentInParent.HasValue) properties.Add(Scalar(OneNoteSchema.LayoutAlignmentInParent, layout.AlignmentInParent.Value));
        if (layout.AlignmentSelf.HasValue) properties.Add(Scalar(OneNoteSchema.LayoutAlignmentSelf, layout.AlignmentSelf.Value));
        if (layout.CollisionPriority.HasValue) properties.Add(Scalar(OneNoteSchema.LayoutCollisionPriority, layout.CollisionPriority.Value));
        AddBoolean(properties, OneNoteSchema.LayoutTightAlignment, layout.TightAlignment);
        return properties;
    }

    private static void AddBoolean(ICollection<OneNoteWriteProperty> properties, uint id, bool? value) { if (value.HasValue) properties.Add(Boolean(id, value.Value)); }
    private static void AddString(ICollection<OneNoteWriteProperty> properties, uint id, string? value) { if (!string.IsNullOrWhiteSpace(value)) properties.Add(Data(id, Unicode(value!))); }
    private static OneNoteWriteProperty Data(uint id, byte[] value) => new OneNoteWriteProperty(id, data: value);
    private static OneNoteWriteProperty Scalar(uint id, ulong value) => new OneNoteWriteProperty(id, scalar: value);
    private static OneNoteWriteProperty Boolean(uint id, bool value) => new OneNoteWriteProperty(id, boolean: value);
    private static OneNoteWriteProperty ObjectReferences(uint id, params OneNoteExtendedGuid[] values) => ObjectReferences(id, (IEnumerable<OneNoteExtendedGuid>)values);
    private static OneNoteWriteProperty ObjectReferences(uint id, IEnumerable<OneNoteExtendedGuid> values) => new OneNoteWriteProperty(id, references: values);
    private static OneNoteWriteProperty ObjectSpaceReferences(uint id, params OneNoteExtendedGuid[] values) => new OneNoteWriteProperty(id, references: values, referenceKind: OneNoteWriteReferenceKind.ObjectSpace);
    private static OneNoteWriteProperty ContextReferences(uint id, params OneNoteExtendedGuid[] values) => new OneNoteWriteProperty(id, references: values, referenceKind: OneNoteWriteReferenceKind.Context);
    private static OneNoteWriteProperty Float(uint id, double value) => Scalar(id, BitConverter.ToUInt32(BitConverter.GetBytes((float)value), 0));
    private static ulong FileTime(DateTime value) => (ulong)value.ToUniversalTime().ToFileTimeUtc();
    private static uint Time32(DateTime value) {
        DateTime utc = value.ToUniversalTime();
        long seconds = (utc.Ticks - new DateTime(1980, 1, 1, 0, 0, 0, DateTimeKind.Utc).Ticks) / TimeSpan.TicksPerSecond;
        if (seconds < 0 || seconds > uint.MaxValue) {
            throw new OneNoteFormatException("ONENOTE_WRITE_TIME_RANGE", "A OneNote content timestamp is outside the Time32 range.");
        }
        return (uint)seconds;
    }
    private static byte[] Unicode(string value) => Encoding.Unicode.GetBytes((value ?? string.Empty) + "\0");
    private static byte[] UInt32Array(IEnumerable<uint> values) { using (var stream = new MemoryStream()) { foreach (uint value in values) FssHttpStreamObjectWriter.WriteUInt32(stream, value); return stream.ToArray(); } }
    private static byte[] OutlineIndentDistances() { using (var stream = new MemoryStream()) { FssHttpStreamObjectWriter.WriteUInt32(stream, 4); foreach (float value in new[] { 0.5F, 0F, 0.75F, 0.75F }) { byte[] data = BitConverter.GetBytes(value); stream.Write(data, 0, data.Length); } return stream.ToArray(); } }
    private static byte[] TableWidths(IList<double> values) { using (var stream = new MemoryStream()) { stream.WriteByte((byte)values.Count); foreach (double value in values) { byte[] data = BitConverter.GetBytes((float)value); stream.Write(data, 0, data.Length); } return stream.ToArray(); } }
    private static uint ToUInt32(double value) => (uint)Math.Max(0, Math.Min(uint.MaxValue, Math.Round(value)));
    private static ushort FontSizeHalfPoints(double value) {
        double halfPoints = Math.Round(value * 2, MidpointRounding.AwayFromZero);
        if (double.IsNaN(value) || double.IsInfinity(value) || halfPoints < 12 || halfPoints > 288) {
            throw new OneNoteFormatException("ONENOTE_WRITE_FONT_SIZE", "A OneNote font size must be from 6 through 144 points.");
        }
        return (ushort)halfPoints;
    }
    private static bool HasExplicitTextStyle(OneNoteTextRun run) =>
        run.Style.Bold.HasValue ||
        run.Style.Italic.HasValue ||
        run.Style.Underline.HasValue ||
        run.Style.Strikethrough.HasValue ||
        run.Style.Superscript.HasValue ||
        run.Style.Subscript.HasValue ||
        run.Style.IsMath.HasValue ||
        run.Style.LanguageId.HasValue ||
        run.Style.FontSize.HasValue ||
        run.Style.ColorArgb.HasValue ||
        run.Style.HighlightColorArgb.HasValue ||
        !string.IsNullOrWhiteSpace(run.Style.FontFamily) ||
        !string.IsNullOrWhiteSpace(run.Hyperlink);
    private static bool IsCompact(OneNoteExtendedGuid? id) => id != null && id.Identifier != Guid.Empty && id.Value > 0 && id.Value <= byte.MaxValue;
    private OneNoteExtendedGuid IdOrNew(OneNoteExtendedGuid? id) => IsCompact(id) ? id! : _ids.New();

    private static Guid? ReadGuidProperty(OneNoteRevisionStoreObject? item, uint propertyId) {
        byte[]? data = OneNoteSemanticMapper.ReadData(item, propertyId);
        return data != null && data.Length >= 16 ? new Guid(data.Take(16).ToArray()) : (Guid?)null;
    }
}

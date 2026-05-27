using System.Globalization;

namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    public static byte[] Write(PdfDoc doc, IEnumerable<IPdfBlock> blocks, PdfOptions opts, string? title, string? author, string? subject, string? keywords) {
        // Layout blocks into pages and create per-page content streams.
        var layout = LayoutBlocks(blocks, opts);
        ValidateNamedDestinationLinks(layout.Pages);

        // Build PDF objects as byte arrays, then assemble with xref.
        var objects = new List<byte[]>();

        // Reserve IDs (1-based). We'll assign as we add to `objects`.
        int infoId = 0, catalogId = 0;
        int pagesId = ReserveObject(objects);
        var pageIds = new List<int>();
        var formFieldIds = new List<int>();

        // Collect fonts used across pages
        var fontObjectIds = new Dictionary<PdfStandardFont, int>();
        int EnsureFont(PdfStandardFont font) {
            if (!fontObjectIds.TryGetValue(font, out int id)) {
                id = AddObject(objects, PdfStandardFontDictionaryBuilder.BuildStandardType1FontObject(font));
                fontObjectIds[font] = id;
            }
            return id;
        }

        // Create content streams and page objects
        int totalPages = layout.Pages.Count;
        var pageNumberInfos = BuildPageNumberInfos(layout.Pages);
        for (int pageIndex = 0; pageIndex < layout.Pages.Count; pageIndex++) {
            var page = layout.Pages[pageIndex];
            // Make a resources dict that references the fonts we declared
            var pageOpts = page.Options ?? opts;
            var pageNumberInfo = pageNumberInfos[pageIndex];
            int headerFooterVariantPageNumber = pageNumberInfo.VariantPageNumber;
            int headerFooterPageNumber = pageNumberInfo.PageNumber;
            int headerFooterTotalPages = pageNumberInfo.TotalPages;
            var pageFontResources = new Dictionary<PdfStandardFont, string>();
            string EnsurePageFontResource(PdfStandardFont font, string preferredAlias) {
                if (pageFontResources.TryGetValue(font, out string? existingAlias)) {
                    return existingAlias;
                }

                string alias = preferredAlias;
                int aliasIndex = pageFontResources.Count + 1;
                while (pageFontResources.Values.Contains(alias, StringComparer.Ordinal)) {
                    alias = "F" + aliasIndex.ToString(CultureInfo.InvariantCulture);
                    aliasIndex++;
                }

                pageFontResources[font] = alias;
                EnsureFont(font);
                return alias;
            }

            var normalFont = ChooseNormal(pageOpts.DefaultFont);
            EnsurePageFontResource(normalFont, "F1");
            if (page.UsedBold) {
                var boldFont = ChooseBold(normalFont);
                EnsurePageFontResource(boldFont, "F2");
            }
            if (page.UsedItalic) {
                var italicFont = ChooseItalic(normalFont);
                EnsurePageFontResource(italicFont, "F3");
            }
            if (page.UsedBoldItalic) {
                var biFont = ChooseBoldItalic(normalFont);
                EnsurePageFontResource(biFont, "F4");
            }
            string? headerFontAlias = null;
            if (pageOpts.HasHeaderContentForPage(headerFooterVariantPageNumber)) {
                headerFontAlias = EnsurePageFontResource(pageOpts.HeaderFont, "F5");
            }
            string? footerFontAlias = null;
            if (pageOpts.HasFooterContentForPage(headerFooterVariantPageNumber)) {
                footerFontAlias = EnsurePageFontResource(pageOpts.FooterFont, "F6");
            }

            var fontResources = new List<(string Name, int Id)>();
            foreach (var kvp in pageFontResources.OrderBy(kvp => kvp.Value, StringComparer.Ordinal)) {
                fontResources.Add((kvp.Value, EnsureFont(kvp.Key)));
            }

            var graphicsStates = new List<(string Name, int Id)>();
            if (page.GraphicsStates.Count > 0) {
                foreach (var state in page.GraphicsStates) {
                    int gsId = AddObject(objects, PdfVisualResourceDictionaryBuilder.BuildExtGStateObject(state.FillOpacity, state.StrokeOpacity));
                    graphicsStates.Add(("/" + state.Name, gsId));
                }
            }

            var shadings = new List<(string Name, int Id)>();
            if (page.Shadings.Count > 0) {
                foreach (var shading in page.Shadings) {
                    int shadingId = AddObject(objects, PdfVisualResourceDictionaryBuilder.BuildAxialShadingObject(
                        shading.X0,
                        shading.Y0,
                        shading.X1,
                        shading.Y1,
                        shading.StartColor,
                        shading.EndColor));
                    shadings.Add(("/" + shading.Name, shadingId));
                }
            }

            // Content stream (append image draw commands at end)
            string contentStr = page.Content;
            if (pageOpts.HasHeaderContentForPage(headerFooterVariantPageNumber)) {
                string headerContent = BuildHeader(pageOpts, headerFooterVariantPageNumber, headerFooterPageNumber, headerFooterTotalPages, pageOpts.HeaderFont, headerFontAlias!);
                contentStr = headerContent + contentStr;
            }
            var xobjects = new List<(string Name, int Id)>();
            if (page.Images.Count > 0) {
                for (int i = 0; i < page.Images.Count; i++) {
                    var img = page.Images[i];
                    string name = "/Im" + (i + 1).ToString(CultureInfo.InvariantCulture);
                    if (!TryBuildImageStream(img, out var imageStream, out string? unsupportedReason)) {
                        throw new NotSupportedException(unsupportedReason ?? "Image format is not supported.");
                    }

                    int? softMaskId = null;
                    if (imageStream.SoftMask != null) {
                        string softMaskDictionary = PdfImageXObjectDictionaryBuilder.BuildStreamDictionary(imageStream.SoftMask);
                        softMaskId = AddStreamObject(objects, softMaskDictionary, imageStream.SoftMask.Data);
                    }

                    string imageDictionary = PdfImageXObjectDictionaryBuilder.BuildStreamDictionary(imageStream, softMaskId);
                    int imgId = AddStreamObject(objects, imageDictionary, imageStream.Data);
                    img.ObjectId = imgId;
                    img.Name = name;
                    xobjects.Add((name, imgId));
                }
                // Append draw commands
                var sbImgs = new StringBuilder();
                foreach (var img in page.Images) {
                    if (img.ClipPath != null) {
                        new ContentStreamBuilder(sbImgs)
                            .SaveState();
                        AppendClipPath(sbImgs, img.ClipPath, img.ClipX, img.ClipY, img.ClipHeight);
                    }

                    new ContentStreamBuilder(sbImgs)
                        .SaveState()
                        .TransformMatrix(img.W, 0, 0, img.H, img.X, img.Y)
                        .XObject(img.Name)
                        .RestoreState();

                    if (img.ClipPath != null) {
                        new ContentStreamBuilder(sbImgs)
                            .RestoreState();
                    }
                }
                contentStr += sbImgs.ToString();
            }
            if (pageOpts.HasFooterContentForPage(headerFooterVariantPageNumber)) {
                string footer = BuildFooter(pageOpts, headerFooterVariantPageNumber, headerFooterPageNumber, headerFooterTotalPages, pageOpts.FooterFont, footerFontAlias!);
                contentStr += footer;
            }
            int contentId = AddStreamObject(objects, Encoding.ASCII.GetBytes(contentStr));
            // Annotations (links and form widgets)
            var pageAnnotIds = new List<int>();
            if (page.Annotations.Count > 0) {
                foreach (var a in page.Annotations) {
                    string annot;
                    if (!string.IsNullOrEmpty(a.Uri)) {
                        annot = PdfAnnotationDictionaryBuilder.BuildUriLinkAnnotation(a.X1, a.Y1, a.X2, a.Y2, a.Uri!, a.Contents);
                    } else if (!string.IsNullOrEmpty(a.DestinationName)) {
                        annot = PdfAnnotationDictionaryBuilder.BuildGoToNamedDestinationLinkAnnotation(a.X1, a.Y1, a.X2, a.Y2, a.DestinationName!, a.Contents);
                    } else {
                        throw new ArgumentException("PDF link annotations require a URI or named destination target.");
                    }

                    int annId = AddObject(objects, annot);
                    pageAnnotIds.Add(annId);
                }
            }
            if (page.FormFields.Count > 0) {
                int helveticaFontId = EnsureFont(PdfStandardFont.Helvetica);
                foreach (var field in page.FormFields) {
                    double appearanceWidth = field.X2 - field.X1;
                    double appearanceHeight = field.Y2 - field.Y1;
                    string appearanceContent = PdfAcroFormDictionaryBuilder.BuildTextFieldAppearanceContent(appearanceWidth, appearanceHeight, field.Value, field.FontSize);
                    byte[] appearanceBytes = PdfEncoding.Latin1GetBytes(appearanceContent);
                    string appearanceDictionary = PdfAcroFormDictionaryBuilder.BuildTextFieldAppearanceStreamDictionary(appearanceWidth, appearanceHeight, helveticaFontId, appearanceBytes.Length);
                    int appearanceId = AddStreamObject(objects, appearanceDictionary, appearanceBytes);
                    string formField = PdfAnnotationDictionaryBuilder.BuildTextFieldWidgetAnnotation(field.X1, field.Y1, field.X2, field.Y2, field.Name, field.Value, field.FontSize, appearanceId);
                    int formFieldId = AddObject(objects, formField);
                    pageAnnotIds.Add(formFieldId);
                    formFieldIds.Add(formFieldId);
                }
            }
            // Page object
            int pageId = AddObject(objects,
                PdfPageDictionaryBuilder.BuildGeneratedPageDictionary(
                    pagesId,
                    pageOpts.PageWidth,
                    pageOpts.PageHeight,
                    contentId,
                    fontResources,
                    xobjects,
                    graphicsStates,
                    shadings,
                    pageAnnotIds));
            pageIds.Add(pageId);
        }

        // Pages tree
        ReplaceObject(objects, pagesId, PdfPageTreeBuilder.BuildPagesDictionary(pageIds));

        int outlinesId = BuildOutlines(objects, layout.Pages, pageIds);
        int namedDestinationsId = BuildNamedDestinations(objects, layout.Pages, pageIds);
        int acroFormId = 0;
        if (formFieldIds.Count > 0) {
            int helveticaFontId = EnsureFont(PdfStandardFont.Helvetica);
            acroFormId = AddObject(objects, PdfAcroFormDictionaryBuilder.BuildAcroFormDictionary(formFieldIds, helveticaFontId));
        }

        // Catalog
        catalogId = AddObject(objects, PdfCatalogDictionaryBuilder.BuildGeneratedCatalogDictionary(pagesId, outlinesId, namedDestinationsId, acroFormId));

        infoId = AddObject(objects, PdfInfoDictionaryBuilder.Build(title, author, subject, keywords));

        return PdfFileAssembler.Assemble(objects, catalogId, infoId);
    }

    private static List<PageNumberInfo> BuildPageNumberInfos(IReadOnlyList<LayoutResult.Page> pages) {
        var seen = new Dictionary<int, int>();
        var pending = new List<(int VariantPageNumber, int PageNumber, int SequenceId)>(pages.Count);
        int nextSequenceId = 0;
        int currentSequenceId = -1;
        int currentVisiblePageNumber = 0;

        foreach (var page in pages) {
            seen.TryGetValue(page.PageGroupId, out int pageNumber);
            pageNumber++;
            seen[page.PageGroupId] = pageNumber;

            bool firstPageOfGroup = pageNumber == 1;
            if (pending.Count == 0 || (firstPageOfGroup && page.Options.HasExplicitPageNumberStart)) {
                currentSequenceId = nextSequenceId++;
                currentVisiblePageNumber = page.Options.HasExplicitPageNumberStart ? page.Options.PageNumberStart : 1;
            } else {
                currentVisiblePageNumber++;
            }

            pending.Add((pageNumber, currentVisiblePageNumber, currentSequenceId));
        }

        var totals = new Dictionary<int, int>();
        foreach (var item in pending) {
            totals[item.SequenceId] = item.PageNumber;
        }

        var infos = new List<PageNumberInfo>(pages.Count);
        foreach (var item in pending) {
            infos.Add(new PageNumberInfo(item.VariantPageNumber, item.PageNumber, totals[item.SequenceId]));
        }

        return infos;
    }

    private static void ValidateNamedDestinationLinks(IReadOnlyList<LayoutResult.Page> pages) {
        var destinations = new HashSet<string>(StringComparer.Ordinal);
        foreach (var page in pages) {
            foreach (var destination in page.NamedDestinations) {
                if (string.IsNullOrWhiteSpace(destination.Name)) {
                    continue;
                }

                if (!destinations.Add(destination.Name)) {
                    throw new ArgumentException("PDF bookmark names must be unique.");
                }
            }
        }

        foreach (var page in pages) {
            foreach (var annotation in page.Annotations) {
                if (string.IsNullOrWhiteSpace(annotation.DestinationName)) {
                    continue;
                }

                if (!destinations.Contains(annotation.DestinationName!)) {
                    throw new ArgumentException($"PDF bookmark link target '{annotation.DestinationName}' was not found.");
                }
            }
        }
    }

    private static int BuildOutlines(List<byte[]> objects, IReadOnlyList<LayoutResult.Page> pages, List<int> pageIds) {
        var root = new OutlineNode { Level = 0 };
        var stack = new Stack<OutlineNode>();
        stack.Push(root);

        for (int pageIndex = 0; pageIndex < pages.Count; pageIndex++) {
            foreach (var bookmark in pages[pageIndex].Bookmarks) {
                if (string.IsNullOrWhiteSpace(bookmark.Title)) {
                    continue;
                }

                int level = Math.Max(1, bookmark.Level);
                while (stack.Count > 1 && stack.Peek().Level >= level) {
                    stack.Pop();
                }

                var parent = stack.Peek();
                var node = new OutlineNode {
                    Level = level,
                    PageIndex = pageIndex,
                    Title = bookmark.Title,
                    Y = bookmark.Y,
                    Parent = parent
                };
                parent.Children.Add(node);
                stack.Push(node);
            }
        }

        if (root.Children.Count == 0) {
            return 0;
        }

        int rootId = ReserveObject(objects);
        foreach (var node in EnumerateOutlines(root.Children)) {
            node.Id = ReserveObject(objects);
        }

        foreach (var node in EnumerateOutlines(root.Children)) {
            int parentId = node.Parent == null || node.Parent == root ? rootId : node.Parent.Id;
            int index = node.Parent?.Children.IndexOf(node) ?? -1;
            int previousId = index > 0 && node.Parent != null ? node.Parent.Children[index - 1].Id : 0;
            int nextId = node.Parent != null && index >= 0 && index < node.Parent.Children.Count - 1 ? node.Parent.Children[index + 1].Id : 0;
            int firstChildId = node.Children.Count > 0 ? node.Children[0].Id : 0;
            int lastChildId = node.Children.Count > 0 ? node.Children[node.Children.Count - 1].Id : 0;
            int descendantCount = CountOutlines(node.Children);
            int pageId = pageIds[node.PageIndex];

            ReplaceObject(objects, node.Id, PdfOutlineDictionaryBuilder.BuildOutlineItem(
                node.Title,
                parentId,
                previousId,
                nextId,
                firstChildId,
                lastChildId,
                descendantCount,
                pageId,
                node.Y));
        }

        ReplaceObject(objects, rootId, PdfOutlineDictionaryBuilder.BuildOutlineRoot(
            root.Children[0].Id,
            root.Children[root.Children.Count - 1].Id,
            CountOutlines(root.Children)));

        return rootId;
    }

    private static int BuildNamedDestinations(List<byte[]> objects, IReadOnlyList<LayoutResult.Page> pages, List<int> pageIds) {
        var destinations = new List<(string Name, int PageIndex, double Y)>();
        var seen = new HashSet<string>(StringComparer.Ordinal);
        for (int pageIndex = 0; pageIndex < pages.Count; pageIndex++) {
            foreach (var destination in pages[pageIndex].NamedDestinations) {
                if (string.IsNullOrWhiteSpace(destination.Name)) {
                    continue;
                }

                if (!seen.Add(destination.Name)) {
                    throw new ArgumentException("PDF bookmark names must be unique.");
                }

                destinations.Add((destination.Name, pageIndex, destination.Y));
            }
        }

        if (destinations.Count == 0) {
            return 0;
        }

        destinations.Sort((left, right) => StringComparer.Ordinal.Compare(left.Name, right.Name));
        var sb = new StringBuilder();
        sb.Append("<< /Names [");
        for (int i = 0; i < destinations.Count; i++) {
            var destination = destinations[i];
            int pageId = pageIds[destination.PageIndex];
            sb.Append(PdfString(destination.Name))
                .Append(" [")
                .Append(PdfSyntaxEscaper.IndirectReference(pageId))
                .Append(" /XYZ 0 ")
                .Append(destination.Y.ToString("0.###", CultureInfo.InvariantCulture))
                .Append(" 0]");
            if (i < destinations.Count - 1) {
                sb.Append(' ');
            }
        }

        sb.Append("] >>\n");
        return AddObject(objects, sb.ToString());
    }

    private static IEnumerable<OutlineNode> EnumerateOutlines(IEnumerable<OutlineNode> nodes) {
        foreach (var node in nodes) {
            yield return node;
            foreach (var child in EnumerateOutlines(node.Children)) {
                yield return child;
            }
        }
    }

    private static int CountOutlines(IEnumerable<OutlineNode> nodes) {
        int count = 0;
        foreach (var node in nodes) {
            count++;
            count += CountOutlines(node.Children);
        }

        return count;
    }

}


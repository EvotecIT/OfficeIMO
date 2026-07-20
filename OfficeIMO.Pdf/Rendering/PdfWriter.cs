using System.Globalization;
using System.IO;
using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    public static byte[] Write(PdfDocument doc, IEnumerable<IPdfBlock> blocks, PdfOptions opts, string? title, string? author, string? subject, string? keywords) =>
        WriteCore(doc, blocks, opts, title, author, subject, keywords, outputStream: null, out _, out _, out _)!;

    internal static (byte[] Bytes, PdfGeneratedDocumentComplianceEvidence ComplianceEvidence) WriteComplianceArtifact(
        PdfDocument doc,
        IEnumerable<IPdfBlock> blocks,
        PdfOptions opts,
        string? title,
        string? author,
        string? subject,
        string? keywords) {
        byte[] bytes = WriteCore(
            doc,
            blocks,
            opts,
            title,
            author,
            subject,
            keywords,
            outputStream: null,
            out _,
            out PdfGeneratedDocumentComplianceEvidence complianceEvidence,
            out _)!;
        return (bytes, complianceEvidence);
    }

    public static long Write(Stream destination, PdfDocument doc, IEnumerable<IPdfBlock> blocks, PdfOptions opts, string? title, string? author, string? subject, string? keywords) {
        return Write(destination, doc, blocks, opts, title, author, subject, keywords, out _);
    }

    internal static long Write(
        Stream destination,
        PdfDocument doc,
        IEnumerable<IPdfBlock> blocks,
        PdfOptions opts,
        string? title,
        string? author,
        string? subject,
        string? keywords,
        out int pageCount) {
        Guard.NotNull(destination, nameof(destination));
        WriteCore(doc, blocks, opts, title, author, subject, keywords, destination, out long bytesWritten, out _, out pageCount);
        return bytesWritten;
    }

    private static byte[]? WriteCore(
        PdfDocument doc,
        IEnumerable<IPdfBlock> blocks,
        PdfOptions opts,
        string? title,
        string? author,
        string? subject,
        string? keywords,
        Stream? outputStream,
        out long bytesWritten,
        out PdfGeneratedDocumentComplianceEvidence complianceEvidence,
        out int pageCount) {
        PdfComplianceValidator.ValidateGenerationOptions(opts);
        opts.ResetEmbeddedFontProgramUsage();

        // Layout blocks into pages and create per-page content streams.
        using var generatedSectionLayout = doc.BeginGeneratedSectionLayout();
        using var layout = LayoutBlocks(blocks, opts);
        pageCount = layout.Pages.Count;
        ValidateNamedDestinationLinks(layout.Pages);
        ValidateUriActionLinks(layout.Pages, opts);
        ValidateGeneratedFormFieldNames(layout.Pages);
        complianceEvidence = CollectGeneratedComplianceEvidence(layout, opts);
        PdfComplianceValidator.ValidateGeneratedDocument(opts, title, complianceEvidence);

        // Build PDF objects as byte arrays, then assemble with xref.
        using var objects = new PdfObjectStore(opts.ObjectBufferMemoryLimitBytes);

        // Reserve IDs (1-based). We'll assign as we add to `objects`.
        int infoId = 0, catalogId = 0;
        int pagesId = ReserveObject(objects);
        bool markInfo = opts.TaggedStructureMode == PdfTaggedStructureMode.CatalogMarkers;
        int structTreeRootId = markInfo ? ReserveObject(objects) : 0;
        var pageIds = new List<int>();
        var formFieldIds = new List<int>();

        // Collect fonts used across pages
        var fontObjectIds = new Dictionary<PdfOptions, Dictionary<PdfStandardFont, int>>();
        var namedFontObjectIds = new Dictionary<PdfOptions, Dictionary<PdfNamedFontFace, int>>();
        var formHelveticaFontIds = new Dictionary<PdfOptions, int>();
        var pendingFontObjects = new List<(int ObjectId, PdfStandardFont Font, PdfOptions Options)>();
        var pendingNamedFontObjects = new List<(int ObjectId, PdfNamedFontFace Font, PdfOptions Options)>();
        bool requiresPdf16FileVersion = false;
        int EnsureFont(PdfStandardFont font, PdfOptions fontOptions) {
            if (!fontObjectIds.TryGetValue(fontOptions, out Dictionary<PdfStandardFont, int>? optionFontObjectIds)) {
                optionFontObjectIds = new Dictionary<PdfStandardFont, int>();
                fontObjectIds[fontOptions] = optionFontObjectIds;
            }

            if (!optionFontObjectIds.TryGetValue(font, out int id)) {
                id = ReserveObject(objects);
                optionFontObjectIds[font] = id;
                pendingFontObjects.Add((id, font, fontOptions));
            }
            return id;
        }

        int EnsureNamedFont(PdfNamedFontFace font, PdfOptions fontOptions) {
            if (!namedFontObjectIds.TryGetValue(fontOptions, out Dictionary<PdfNamedFontFace, int>? optionFontObjectIds)) {
                optionFontObjectIds = new Dictionary<PdfNamedFontFace, int>();
                namedFontObjectIds[fontOptions] = optionFontObjectIds;
            }

            if (!optionFontObjectIds.TryGetValue(font, out int id)) {
                id = ReserveObject(objects);
                optionFontObjectIds[font] = id;
                pendingNamedFontObjects.Add((id, font, fontOptions));
            }

            return id;
        }

        void MaterializePendingFontObjects() {
            foreach (var pendingFont in pendingFontObjects) {
                if (pendingFont.Options.TryGetEmbeddedStandardFontProgramForGeneration(pendingFont.Font, out PdfEmbeddedFont? _, out PdfTrueTypeFontProgram? fontProgram) &&
                    fontProgram != null) {
                    byte[] fontData = fontProgram.BuildSubsetFontFile();
                    string fontFileExtraEntries = "/Length1 " + fontData.Length.ToString(CultureInfo.InvariantCulture);
                    int fontFileId = pendingFont.Options.CompressEmbeddedFonts
                        ? AddFlateStreamObject(objects, fontData, fontFileExtraEntries)
                        : AddStreamObject(
                            objects,
                            "<< /Length " + fontData.Length.ToString(CultureInfo.InvariantCulture) + " " + fontFileExtraEntries + " >>",
                            fontData);
                    int descriptorId = AddObject(objects, PdfStandardFontDictionaryBuilder.BuildTrueTypeFontDescriptorObject(fontProgram, fontFileId));
                    int descendantFontId = AddObject(objects, PdfStandardFontDictionaryBuilder.BuildCidFontType2DescendantObject(fontProgram, descriptorId));
                    int toUnicodeObjectId = AddStreamObject(objects, PdfToUnicodeCMapBuilder.BuildIdentityGlyphToUnicodeCMap(fontProgram));
                    ReplaceObject(objects, pendingFont.ObjectId, PdfStandardFontDictionaryBuilder.BuildEmbeddedType0FontObject(fontProgram, descendantFontId, toUnicodeObjectId));
                } else if (pendingFont.Options.TryGetEmbeddedStandardOpenTypeCffFontProgramForGeneration(pendingFont.Font, out PdfEmbeddedFont? _, out PdfOpenTypeCffFontProgram? cffFontProgram) &&
                    cffFontProgram != null) {
                    requiresPdf16FileVersion = true;
                    PdfOpenTypeCffCompactFontFile compactFontFile = cffFontProgram.BuildCompactOpenTypeFontFilePlan();
                    pendingFont.Options.AddFontDiagnostics(
                        pendingFont.Font,
                        PdfFontDiagnostics.AnalyzeOpenTypeCffCompactEmbedding(cffFontProgram, "embedded-font:" + pendingFont.Font, compactFontFile));
                    byte[] fontData = compactFontFile.Data;
                    string fontFileExtraEntries = "/Subtype /OpenType /Length1 " + fontData.Length.ToString(CultureInfo.InvariantCulture);
                    int fontFileId = pendingFont.Options.CompressEmbeddedFonts
                        ? AddFlateStreamObject(objects, fontData, fontFileExtraEntries)
                        : AddStreamObject(
                            objects,
                            "<< /Length " + fontData.Length.ToString(CultureInfo.InvariantCulture) + " " + fontFileExtraEntries + " >>",
                            fontData);
                    int descriptorId = AddObject(objects, PdfStandardFontDictionaryBuilder.BuildOpenTypeCffFontDescriptorObject(cffFontProgram, fontFileId));
                    int descendantFontId = AddObject(objects, PdfStandardFontDictionaryBuilder.BuildCidFontType0DescendantObject(cffFontProgram, descriptorId));
                    int toUnicodeObjectId = AddStreamObject(objects, PdfToUnicodeCMapBuilder.BuildIdentityGlyphToUnicodeCMap(cffFontProgram));
                    ReplaceObject(objects, pendingFont.ObjectId, PdfStandardFontDictionaryBuilder.BuildEmbeddedType0FontObject(cffFontProgram, descendantFontId, toUnicodeObjectId));
                } else {
                    int toUnicodeObjectId = pendingFont.Options.IncludeStandardFontToUnicodeMaps
                        ? AddStreamObject(objects, PdfToUnicodeCMapBuilder.BuildWinAnsiToUnicodeCMap())
                        : 0;
                    ReplaceObject(objects, pendingFont.ObjectId, PdfStandardFontDictionaryBuilder.BuildStandardType1FontObject(pendingFont.Font, toUnicodeObjectId));
                }
            }

            foreach (var pendingFont in pendingNamedFontObjects) {
                if (pendingFont.Options.TryGetNamedFontProgramForGeneration(pendingFont.Font, out PdfTrueTypeFontProgram? fontProgram) &&
                    fontProgram != null) {
                    byte[] fontData = fontProgram.BuildSubsetFontFile();
                    string fontFileExtraEntries = "/Length1 " + fontData.Length.ToString(CultureInfo.InvariantCulture);
                    int fontFileId = pendingFont.Options.CompressEmbeddedFonts
                        ? AddFlateStreamObject(objects, fontData, fontFileExtraEntries)
                        : AddStreamObject(
                            objects,
                            "<< /Length " + fontData.Length.ToString(CultureInfo.InvariantCulture) + " " + fontFileExtraEntries + " >>",
                            fontData);
                    int descriptorId = AddObject(objects, PdfStandardFontDictionaryBuilder.BuildTrueTypeFontDescriptorObject(fontProgram, fontFileId));
                    int descendantFontId = AddObject(objects, PdfStandardFontDictionaryBuilder.BuildCidFontType2DescendantObject(fontProgram, descriptorId));
                    int toUnicodeObjectId = AddStreamObject(objects, PdfToUnicodeCMapBuilder.BuildIdentityGlyphToUnicodeCMap(fontProgram));
                    ReplaceObject(objects, pendingFont.ObjectId, PdfStandardFontDictionaryBuilder.BuildEmbeddedType0FontObject(fontProgram, descendantFontId, toUnicodeObjectId));
                } else if (pendingFont.Options.TryGetNamedOpenTypeCffFontProgramForGeneration(pendingFont.Font, out PdfOpenTypeCffFontProgram? cffFontProgram) &&
                           cffFontProgram != null) {
                    requiresPdf16FileVersion = true;
                    PdfOpenTypeCffCompactFontFile compactFontFile = cffFontProgram.BuildCompactOpenTypeFontFilePlan();
                    pendingFont.Options.AddFontDiagnostics(
                        PdfStandardFont.Helvetica,
                        PdfFontDiagnostics.AnalyzeOpenTypeCffCompactEmbedding(cffFontProgram, "named-font:" + pendingFont.Font.FaceKey, compactFontFile));
                    byte[] fontData = compactFontFile.Data;
                    string fontFileExtraEntries = "/Subtype /OpenType /Length1 " + fontData.Length.ToString(CultureInfo.InvariantCulture);
                    int fontFileId = pendingFont.Options.CompressEmbeddedFonts
                        ? AddFlateStreamObject(objects, fontData, fontFileExtraEntries)
                        : AddStreamObject(
                            objects,
                            "<< /Length " + fontData.Length.ToString(CultureInfo.InvariantCulture) + " " + fontFileExtraEntries + " >>",
                            fontData);
                    int descriptorId = AddObject(objects, PdfStandardFontDictionaryBuilder.BuildOpenTypeCffFontDescriptorObject(cffFontProgram, fontFileId));
                    int descendantFontId = AddObject(objects, PdfStandardFontDictionaryBuilder.BuildCidFontType0DescendantObject(cffFontProgram, descriptorId));
                    int toUnicodeObjectId = AddStreamObject(objects, PdfToUnicodeCMapBuilder.BuildIdentityGlyphToUnicodeCMap(cffFontProgram));
                    ReplaceObject(objects, pendingFont.ObjectId, PdfStandardFontDictionaryBuilder.BuildEmbeddedType0FontObject(cffFontProgram, descendantFontId, toUnicodeObjectId));
                } else {
                    throw new InvalidOperationException("Named font resource '" + pendingFont.Font.FamilyName + "' could not be materialized.");
                }
            }
        }

        int EnsureFormHelveticaFont(PdfOptions formOptions) {
            if (!formHelveticaFontIds.TryGetValue(formOptions, out int formHelveticaFontId)) {
                formHelveticaFontId = ShouldUseEmbeddedFormHelveticaFont(formOptions)
                    ? EnsureFont(PdfStandardFont.Helvetica, formOptions)
                    : AddObject(objects, PdfStandardFontDictionaryBuilder.BuildStandardType1FontObject(PdfStandardFont.Helvetica));
                formHelveticaFontIds[formOptions] = formHelveticaFontId;
            }

            return formHelveticaFontId;
        }

        bool ShouldUseEmbeddedFormHelveticaFont(PdfOptions formOptions) =>
            (formOptions.TryGetEmbeddedStandardFontProgram(PdfStandardFont.Helvetica, out PdfTrueTypeFontProgram? fontProgram) &&
            fontProgram != null) ||
            (formOptions.TryGetEmbeddedStandardOpenTypeCffFontProgram(PdfStandardFont.Helvetica, out PdfOpenTypeCffFontProgram? cffFontProgram) &&
            cffFontProgram != null);

        // Create content streams and page objects
        int totalPages = pageCount;
        var pageNumberInfos = BuildPageNumberInfos(layout.Pages);
        int nextStructParentIndex = 0;
        var imageXObjectIds = new Dictionary<string, int>(StringComparer.Ordinal);
        var optimizedImageCache = new Dictionary<string, OfficeImageOptimizationResult>(StringComparer.Ordinal);

        int EnsureImageXObject(PdfImageStream imageStream) {
            string cacheKey = BuildImageXObjectCacheKey(imageStream);
            if (imageXObjectIds.TryGetValue(cacheKey, out int existingImageId)) {
                return existingImageId;
            }

            int? softMaskId = null;
            if (imageStream.SoftMask != null) {
                string softMaskDictionary = PdfImageXObjectDictionaryBuilder.BuildStreamDictionary(imageStream.SoftMask);
                softMaskId = AddStreamObject(objects, softMaskDictionary, imageStream.SoftMask.Data);
            }

            string imageDictionary = PdfImageXObjectDictionaryBuilder.BuildStreamDictionary(imageStream, softMaskId);
            int imageId = AddStreamObject(objects, imageDictionary, imageStream.Data);
            imageXObjectIds[cacheKey] = imageId;
            return imageId;
        }

        var layerDefinitions = layout.Pages
            .SelectMany(page => page.Layers)
            .Distinct()
            .OrderBy(definition => definition.Id)
            .ToList();
        var optionalContentGroupIds = new Dictionary<PdfLayerDefinition, int>();
        foreach (PdfLayerDefinition definition in layerDefinitions) {
            optionalContentGroupIds[definition] = AddObject(objects, PdfOptionalContentDictionaryBuilder.BuildGroup(definition));
        }
        int optionalContentPropertiesId = layerDefinitions.Count == 0
            ? 0
            : AddObject(objects, PdfOptionalContentDictionaryBuilder.BuildProperties(layerDefinitions, optionalContentGroupIds));

        for (int pageIndex = 0; pageIndex < layout.Pages.Count; pageIndex++) {
            var page = layout.Pages[pageIndex];
            // Make a resources dict that references the fonts we declared
            var pageOpts = page.Options ?? opts;
            var pageNumberInfo = pageNumberInfos[pageIndex];
            int headerFooterVariantPageNumber = pageNumberInfo.VariantPageNumber;
            int headerFooterPageNumber = pageNumberInfo.PageNumber;
            int headerFooterTotalPages = pageNumberInfo.TotalPages;
            string pageLayoutContent = layout.ReadContent(page.Content);
            var pageFontResources = new Dictionary<PdfStandardFont, string>();
            var pageNamedFontResources = new Dictionary<PdfNamedFontFace, string>();
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
                EnsureFont(font, pageOpts);
                return alias;
            }
            bool LayoutUsesFontResource(string resourceName) {
                string qualifiedName = "/" + resourceName;
                if (UsesPdfResource(pageLayoutContent, qualifiedName)) {
                    return true;
                }

                foreach (PageEffectGroup effect in page.EffectGroups) {
                    if (UsesPdfResource(layout.ReadContent(effect.Content), qualifiedName)) {
                        return true;
                    }
                }

                return false;
            }

            var normalFont = ChooseNormal(pageOpts.DefaultFont);
            if (page.UsedFonts.Count > 0) {
                EnsurePageFontResource(normalFont, "F1");
                if (page.UsedBold) {
                    EnsurePageFontResource(ChooseBold(normalFont), "F2");
                }
                if (page.UsedItalic) {
                    EnsurePageFontResource(ChooseItalic(normalFont), "F3");
                }
                if (page.UsedBoldItalic) {
                    EnsurePageFontResource(ChooseBoldItalic(normalFont), "F4");
                }
            }
            if (LayoutUsesFontResource("F1")) {
                EnsurePageFontResource(normalFont, "F1");
            }
            if (LayoutUsesFontResource("F2")) {
                EnsurePageFontResource(ChooseBold(normalFont), "F2");
            }
            if (LayoutUsesFontResource("F3")) {
                EnsurePageFontResource(ChooseItalic(normalFont), "F3");
            }
            if (LayoutUsesFontResource("F4")) {
                EnsurePageFontResource(ChooseBoldItalic(normalFont), "F4");
            }
            foreach (PdfStandardFont usedFont in page.UsedFonts) {
                EnsurePageFontResource(usedFont, GetStandardFontResourceName(usedFont, normalFont));
            }
            foreach (PdfNamedFontFace usedFont in page.UsedNamedFonts.OrderBy(font => font.ResourceName, StringComparer.Ordinal)) {
                pageNamedFontResources[usedFont] = usedFont.ResourceName;
                EnsureNamedFont(usedFont, pageOpts);
            }
            PdfTextWatermark? textWatermark = pageOpts.GetTextWatermarkForPage(headerFooterVariantPageNumber);
            string? watermarkFontAlias = null;
            string? textWatermarkGraphicsStateName = null;
            if (textWatermark != null && textWatermark.Opacity > 0D) {
                watermarkFontAlias = EnsurePageFontResource(GetTextWatermarkFont(textWatermark), "FW");
                EnsureTextWatermarkFontResources(textWatermark, pageOpts, EnsurePageFontResource);
                if (textWatermark.Opacity < 1D) {
                    textWatermarkGraphicsStateName = EnsureHeaderFooterGraphicsState(page, textWatermark.Opacity, textWatermark.Opacity);
                }
            }
            PdfPageBackgroundImage? pageBackgroundImage = pageOpts.PageBackgroundImageSnapshot;
            if (pageBackgroundImage != null && pageBackgroundImage.Opacity > 0D) {
                AddPageBackgroundImage(page, pageOpts, pageBackgroundImage);
            }
            PdfImageWatermark? imageWatermark = pageOpts.GetImageWatermarkForPage(headerFooterVariantPageNumber);
            if (imageWatermark != null && imageWatermark.Opacity > 0D) {
                AddImageWatermark(page, pageOpts, imageWatermark);
            }
            PdfPageBorder? pageBorder = pageOpts.PageBorderSnapshot;
            string? pageBorderGraphicsStateName = null;
            if (pageBorder != null && pageBorder.Opacity > 0D && pageBorder.Opacity < 1D) {
                pageBorderGraphicsStateName = EnsureHeaderFooterGraphicsState(page, 1D, pageBorder.Opacity);
            }
            string pageBackgroundShapeContent = BuildPageBackgroundShapes(page, pageOpts.PageBackgroundShapeSnapshots);
            void EnsurePageNamedFontResource(PdfNamedFontFace font) {
                pageNamedFontResources[font] = font.ResourceName;
                EnsureNamedFont(font, pageOpts);
            }

            string? headerFontAlias = null;
            if (pageOpts.HasHeaderTextContentForPage(headerFooterVariantPageNumber)) {
                if (TryResolvePageTextNamedFont(pageOpts, pageOpts.HeaderFontFamily, pageOpts.HeaderFont, out PdfNamedFontFace headerNamedFont)) {
                    EnsurePageNamedFontResource(headerNamedFont);
                    headerFontAlias = headerNamedFont.ResourceName;
                } else {
                    headerFontAlias = EnsurePageFontResource(pageOpts.HeaderFont, "F5");
                }
                EnsurePageTextFontResources(pageOpts, headerFooterVariantPageNumber, headerFooterPageNumber, headerFooterTotalPages, totalPages, pageOpts.HeaderFont, pageOpts.HeaderFontSize, isHeader: true, EnsurePageFontResource, EnsurePageNamedFontResource);
            }
            string? footerFontAlias = null;
            if (pageOpts.HasFooterTextContentForPage(headerFooterVariantPageNumber)) {
                if (TryResolvePageTextNamedFont(pageOpts, pageOpts.FooterFontFamily, pageOpts.FooterFont, out PdfNamedFontFace footerNamedFont)) {
                    EnsurePageNamedFontResource(footerNamedFont);
                    footerFontAlias = footerNamedFont.ResourceName;
                } else {
                    footerFontAlias = EnsurePageFontResource(pageOpts.FooterFont, "F6");
                }
                EnsurePageTextFontResources(pageOpts, headerFooterVariantPageNumber, headerFooterPageNumber, headerFooterTotalPages, totalPages, pageOpts.FooterFont, pageOpts.FooterFontSize, isHeader: false, EnsurePageFontResource, EnsurePageNamedFontResource);
            }

            string headerFooterShapeContent = BuildHeaderFooterShapes(
                page,
                pageOpts,
                headerFooterVariantPageNumber,
                headerFooterPageNumber,
                headerFooterTotalPages,
                totalPages);

            var fontResources = new List<(string Name, int Id)>();
            foreach (var kvp in pageFontResources.OrderBy(kvp => kvp.Value, StringComparer.Ordinal)) {
                fontResources.Add((kvp.Value, EnsureFont(kvp.Key, pageOpts)));
            }
            foreach (var kvp in pageNamedFontResources.OrderBy(kvp => kvp.Value, StringComparer.Ordinal)) {
                fontResources.Add((kvp.Value, EnsureNamedFont(kvp.Key, pageOpts)));
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
                    string shadingObject = shading.IsRadial
                        ? PdfVisualResourceDictionaryBuilder.BuildRadialShadingObject(
                            shading.X0,
                            shading.Y0,
                            shading.R0,
                            shading.X1,
                            shading.Y1,
                            shading.R1,
                            shading.Stops)
                        : PdfVisualResourceDictionaryBuilder.BuildAxialShadingObject(
                            shading.X0,
                            shading.Y0,
                            shading.X1,
                            shading.Y1,
                            shading.Stops);
                    int shadingId = AddObject(objects, shadingObject);
                    shadings.Add(("/" + shading.Name, shadingId));
                }
            }

            // Content stream (append image draw commands at end)
            AddHeaderFooterImages(
                page,
                pageOpts,
                headerFooterVariantPageNumber,
                headerFooterPageNumber,
                headerFooterTotalPages,
                totalPages);
            if (markInfo) {
                AssignFigureMarkedContentIds(page);
                AssignStructParentIndex(page, ref nextStructParentIndex);
            }

            var xobjects = new List<(string Name, int Id)>();
            if (page.Images.Count > 0) {
                var pageImageResourceNames = new Dictionary<int, string>();
                for (int i = 0; i < page.Images.Count; i++) {
                    var img = page.Images[i];
                    ApplyPlacementAwareImageOptimization(img, pageOpts, optimizedImageCache);
                    if (!TryBuildImageStream(img, out var imageStream, out string? unsupportedReason)) {
                        throw new NotSupportedException(unsupportedReason ?? "Image format is not supported.");
                    }

                    int imgId = EnsureImageXObject(imageStream);
                    if (!pageImageResourceNames.TryGetValue(imgId, out string? name)) {
                        name = "/Im" + (pageImageResourceNames.Count + 1).ToString(CultureInfo.InvariantCulture);
                        pageImageResourceNames[imgId] = name;
                        xobjects.Add((name, imgId));
                    }

                    img.ObjectId = imgId;
                    img.Name = name;
                }
            }

            if (page.EffectGroups.Count > 0) {
                for (int effectIndex = 0; effectIndex < page.EffectGroups.Count; effectIndex++) {
                    PageEffectGroup effect = page.EffectGroups[effectIndex];
                    string effectContent = ReplaceInlineImageDrawTokens(layout.ReadContent(effect.Content), page.Images);
                    effectContent = ReplaceInlineEffectGroupTokens(effectContent, page.EffectGroups, effectIndex);
                    byte[] effectBytes = PdfEncoding.Latin1GetBytes(effectContent);
                    string dictionary = PdfTransparencyGroupDictionaryBuilder.BuildStreamDictionary(
                        pageOpts.PageWidth,
                        pageOpts.PageHeight,
                        effectBytes.Length,
                        FilterPdfResources(effectContent, fontResources),
                        FilterPdfResources(effectContent, xobjects),
                        FilterPdfResources(effectContent, graphicsStates),
                        FilterPdfResources(effectContent, shadings));
                    int effectId = AddStreamObject(objects, dictionary, effectBytes);
                    effect.Name = "/Fx" + (effectIndex + 1).ToString(CultureInfo.InvariantCulture);
                    effect.ObjectId = effectId;
                    xobjects.Add((effect.Name, effectId));
                }
            }

            string pageBackgroundContent = BuildPageBackground(page, pageOpts, pageBackgroundShapeContent, textWatermark, watermarkFontAlias, pageFontResources, textWatermarkGraphicsStateName, pageBorder, pageBorderGraphicsStateName, markInfo);
            string contentStr = pageBackgroundContent + WrapArtifactContent(headerFooterShapeContent, markInfo);
            if (pageOpts.HasHeaderTextContentForPage(headerFooterVariantPageNumber)) {
                string headerContent = BuildHeader(pageOpts, headerFooterVariantPageNumber, headerFooterPageNumber, headerFooterTotalPages, totalPages, pageOpts.HeaderFont, headerFontAlias!, pageFontResources, pageNamedFontResources);
                contentStr += WrapArtifactContent(headerContent, markInfo);
            }
            string pageContent = ReplaceInlineImageDrawTokens(pageLayoutContent, page.Images);
            contentStr += ReplaceInlineEffectGroupTokens(pageContent, page.EffectGroups, page.EffectGroups.Count);
            if (page.Images.Count > 0) {
                var sbImgs = new StringBuilder();
                foreach (var img in page.Images) {
                    if (img.IsBackgroundDecoration || !string.IsNullOrEmpty(img.InlineDrawToken)) {
                        continue;
                    }

                    AppendPageImageDraw(sbImgs, img);
                    if (img.DebugBox) {
                        DrawRowRect(sbImgs, new PdfColor(1D, 0D, 1D), 0.6D, img.X, img.Y, img.W, img.H, markInfo);
                    }
                }

                contentStr += sbImgs.ToString();
            }
            if (pageOpts.HasFooterTextContentForPage(headerFooterVariantPageNumber)) {
                string footer = BuildFooter(pageOpts, headerFooterVariantPageNumber, headerFooterPageNumber, headerFooterTotalPages, totalPages, pageOpts.FooterFont, footerFontAlias!, pageFontResources, pageNamedFontResources);
                contentStr += WrapArtifactContent(footer, markInfo);
            }
            bool flattenVisualAnnotations = pageOpts.FlattenVisualAnnotations;
            if (flattenVisualAnnotations) {
                contentStr += BuildFlattenedVisualAnnotationContent(
                    page,
                    pageOpts,
                    objects,
                    xobjects,
                    EnsureFont,
                    EnsureFormHelveticaFont,
                    markInfo);
            }

            byte[] contentBytes = Encoding.ASCII.GetBytes(contentStr);
            int contentId = pageOpts.CompressContentStreams
                ? AddFlateStreamObject(objects, contentBytes)
                : AddStreamObject(objects, contentBytes);
            // Annotations (links and form widgets)
            var pageAnnotIds = new List<int>();
            if (page.Annotations.Count > 0) {
                foreach (var a in page.Annotations) {
                    if (markInfo && !a.StructParentIndex.HasValue) {
                        a.StructParentIndex = nextStructParentIndex++;
                    }

                    string annot;
                    if (!string.IsNullOrEmpty(a.Uri)) {
                        annot = PdfAnnotationDictionaryBuilder.BuildUriLinkAnnotation(a.X1, a.Y1, a.X2, a.Y2, a.Uri!, a.Contents, a.StructParentIndex);
                    } else if (!string.IsNullOrEmpty(a.DestinationName)) {
                        annot = PdfAnnotationDictionaryBuilder.BuildGoToNamedDestinationLinkAnnotation(a.X1, a.Y1, a.X2, a.Y2, a.DestinationName!, a.Contents, a.StructParentIndex);
                    } else {
                        throw new ArgumentException("PDF link annotations require a URI or named destination target.");
                    }

                    int annId = AddObject(objects, annot);
                    a.ObjectId = annId;
                    if (markInfo && a.StructParentIndex.HasValue) {
                        if (!a.StructElementIndex.HasValue && a.LinkedImage?.StructElementIndex.HasValue == true) {
                            a.StructElementIndex = a.LinkedImage.StructElementIndex;
                        }

                        if (a.StructElementIndex.HasValue &&
                            a.StructElementIndex.Value >= 0 &&
                            a.StructElementIndex.Value < page.StructElements.Count) {
                            AttachAnnotationToStructElement(page.StructElements[a.StructElementIndex.Value], annId, a.StructParentIndex.Value);
                        } else {
                            page.StructElements.Add(new PageStructElement {
                                StructureType = "Link",
                                AnnotationObjectId = annId,
                                AnnotationStructParentIndex = a.StructParentIndex
                            });
                        }
                    }

                    pageAnnotIds.Add(annId);
                }
            }
            if (page.TextAnnotations.Count > 0) {
                foreach (var annotation in page.TextAnnotations) {
                    AnnotationStructureReference? annotationStructureReference = RegisterAnnotationStructureReference(page, markInfo, ref nextStructParentIndex, "Annot");
                    string annot = PdfAnnotationDictionaryBuilder.BuildTextAnnotation(
                        annotation.X1,
                        annotation.Y1,
                        annotation.X2,
                        annotation.Y2,
                        annotation.Contents,
                        annotation.Icon,
                        annotation.Color,
                        annotation.Open,
                        annotationStructureReference?.StructParentIndex);
                    int annId = AddObject(objects, annot);
                    annotation.ObjectId = annId;
                    CompleteAnnotationStructureReference(page, annotationStructureReference, annId);
                    pageAnnotIds.Add(annId);
                }
            }
            if (!flattenVisualAnnotations && page.FreeTextAnnotations.Count > 0) {
                foreach (var annotation in page.FreeTextAnnotations) {
                    AnnotationStructureReference? annotationStructureReference = RegisterAnnotationStructureReference(page, markInfo, ref nextStructParentIndex, "Annot");
                    double appearanceWidth = annotation.X2 - annotation.X1;
                    double appearanceHeight = annotation.Y2 - annotation.Y1;
                    string appearanceContent = BuildFreeTextAnnotationAppearanceContent(
                        annotation,
                        appearanceWidth,
                        appearanceHeight,
                        pageOpts,
                        EnsureFont,
                        out IReadOnlyList<(string Name, int Id)> appearanceFontResources);
                    byte[] appearanceBytes = PdfEncoding.Latin1GetBytes(appearanceContent);
                    string appearanceDictionary = PdfAnnotationDictionaryBuilder.BuildAppearanceStreamDictionary(appearanceWidth, appearanceHeight, appearanceBytes.Length, appearanceFontResources);
                    int appearanceId = AddStreamObject(objects, appearanceDictionary, appearanceBytes);
                    string annot = PdfAnnotationDictionaryBuilder.BuildFreeTextAnnotation(
                        annotation.X1,
                        annotation.Y1,
                        annotation.X2,
                        annotation.Y2,
                        annotation.Contents,
                        annotation.FontSize,
                        annotation.TextColor,
                        annotation.BorderColor,
                        annotation.BorderWidth,
                        annotation.FillColor,
                        appearanceId,
                        annotationStructureReference?.StructParentIndex);
                    int annId = AddObject(objects, annot);
                    annotation.ObjectId = annId;
                    CompleteAnnotationStructureReference(page, annotationStructureReference, annId);
                    pageAnnotIds.Add(annId);
                }
            }
            if (!flattenVisualAnnotations && page.HighlightAnnotations.Count > 0) {
                foreach (var annotation in page.HighlightAnnotations) {
                    AnnotationStructureReference? annotationStructureReference = RegisterAnnotationStructureReference(page, markInfo, ref nextStructParentIndex, "Annot");
                    double appearanceWidth = annotation.X2 - annotation.X1;
                    double appearanceHeight = annotation.Y2 - annotation.Y1;
                    string appearanceContent = PdfAnnotationDictionaryBuilder.BuildHighlightAppearanceContent(appearanceWidth, appearanceHeight, annotation.Color);
                    byte[] appearanceBytes = PdfEncoding.Latin1GetBytes(appearanceContent);
                    string appearanceDictionary = PdfAnnotationDictionaryBuilder.BuildAppearanceStreamDictionary(appearanceWidth, appearanceHeight, appearanceBytes.Length, usesHighlightBlendMode: true);
                    int appearanceId = AddStreamObject(objects, appearanceDictionary, appearanceBytes);
                    string annot = PdfAnnotationDictionaryBuilder.BuildHighlightAnnotation(
                        annotation.X1,
                        annotation.Y1,
                        annotation.X2,
                        annotation.Y2,
                        annotation.Contents,
                        annotation.Color,
                        appearanceId,
                        annotationStructureReference?.StructParentIndex);
                    int annId = AddObject(objects, annot);
                    annotation.ObjectId = annId;
                    CompleteAnnotationStructureReference(page, annotationStructureReference, annId);
                    pageAnnotIds.Add(annId);
                }
            }
            if (page.FormFields.Count > 0) {
                foreach (var field in page.FormFields) {
                    string formField;
                    double appearanceWidth = field.X2 - field.X1;
                    double appearanceHeight = field.Y2 - field.Y1;
                    if (field.Kind == FormFieldAnnotationKind.RadioButtonGroup) {
                        int parentFieldId = ReserveObject(objects);
                        string offAppearance = PdfAcroFormDictionaryBuilder.BuildRadioButtonAppearanceContent(field.ButtonSize, field.ButtonSize, selected: false, field.Style);
                        byte[] offAppearanceBytes = PdfEncoding.Latin1GetBytes(offAppearance);
                        string offAppearanceDictionary = PdfAcroFormDictionaryBuilder.BuildCheckBoxAppearanceStreamDictionary(field.ButtonSize, field.ButtonSize, offAppearanceBytes.Length);
                        int offAppearanceId = AddStreamObject(objects, offAppearanceDictionary, offAppearanceBytes);

                        string selectedAppearance = PdfAcroFormDictionaryBuilder.BuildRadioButtonAppearanceContent(field.ButtonSize, field.ButtonSize, selected: true, field.Style);
                        byte[] selectedAppearanceBytes = PdfEncoding.Latin1GetBytes(selectedAppearance);
                        string selectedAppearanceDictionary = PdfAcroFormDictionaryBuilder.BuildCheckBoxAppearanceStreamDictionary(field.ButtonSize, field.ButtonSize, selectedAppearanceBytes.Length);
                        int selectedAppearanceId = AddStreamObject(objects, selectedAppearanceDictionary, selectedAppearanceBytes);

                        var widgetObjectIds = new List<int>(field.Options.Count);
                        for (int optionIndex = 0; optionIndex < field.Options.Count; optionIndex++) {
                            AnnotationStructureReference? widgetStructureReference = RegisterAnnotationStructureReference(page, markInfo, ref nextStructParentIndex, "Form");
                            double widgetTop = field.Y2 - optionIndex * (field.ButtonSize + field.ButtonGap);
                            double widgetBottom = widgetTop - field.ButtonSize;
                            string widget = PdfAnnotationDictionaryBuilder.BuildRadioButtonWidgetAnnotation(
                                field.X1,
                                widgetBottom,
                                field.X1 + field.ButtonSize,
                                widgetTop,
                                parentFieldId,
                                field.Options[optionIndex],
                                field.Value,
                                offAppearanceId,
                                selectedAppearanceId,
                                field.Style,
                                widgetStructureReference?.StructParentIndex);
                            int widgetObjectId = AddObject(objects, widget);
                            CompleteAnnotationStructureReference(page, widgetStructureReference, widgetObjectId);
                            widgetObjectIds.Add(widgetObjectId);
                            pageAnnotIds.Add(widgetObjectId);
                        }

                        ReplaceObject(objects, parentFieldId, PdfAnnotationDictionaryBuilder.BuildRadioButtonFieldDictionary(field.Name, field.Options, field.Value, widgetObjectIds, field.Style));
                        formFieldIds.Add(parentFieldId);
                        continue;
                    }

                    AnnotationStructureReference? formWidgetStructureReference = RegisterAnnotationStructureReference(page, markInfo, ref nextStructParentIndex, "Form");
                    if (field.Kind == FormFieldAnnotationKind.CheckBox) {
                        string offAppearance = PdfAcroFormDictionaryBuilder.BuildCheckBoxAppearanceContent(appearanceWidth, appearanceHeight, selected: false, field.Style);
                        byte[] offAppearanceBytes = PdfEncoding.Latin1GetBytes(offAppearance);
                        string offAppearanceDictionary = PdfAcroFormDictionaryBuilder.BuildCheckBoxAppearanceStreamDictionary(appearanceWidth, appearanceHeight, offAppearanceBytes.Length);
                        int offAppearanceId = AddStreamObject(objects, offAppearanceDictionary, offAppearanceBytes);

                        string checkedAppearance = PdfAcroFormDictionaryBuilder.BuildCheckBoxAppearanceContent(appearanceWidth, appearanceHeight, selected: true, field.Style);
                        byte[] checkedAppearanceBytes = PdfEncoding.Latin1GetBytes(checkedAppearance);
                        string checkedAppearanceDictionary = PdfAcroFormDictionaryBuilder.BuildCheckBoxAppearanceStreamDictionary(appearanceWidth, appearanceHeight, checkedAppearanceBytes.Length);
                        int checkedAppearanceId = AddStreamObject(objects, checkedAppearanceDictionary, checkedAppearanceBytes);

                        formField = PdfAnnotationDictionaryBuilder.BuildCheckBoxWidgetAnnotation(field.X1, field.Y1, field.X2, field.Y2, field.Name, field.IsChecked, field.CheckedValueName, offAppearanceId, checkedAppearanceId, field.Style, formWidgetStructureReference?.StructParentIndex);
                    } else if (field.Kind == FormFieldAnnotationKind.Choice) {
                        string appearanceValue = field.Values.Count > 1 ? string.Join(", ", field.Values) : field.Value;
                        string appearanceContent = BuildFormFieldTextAppearanceContent(
                            appearanceWidth,
                            appearanceHeight,
                            appearanceValue,
                            field.FontSize,
                            field.Style,
                            pageOpts,
                            EnsureFont,
                            out IReadOnlyList<(string Name, int Id)> appearanceFontResources);
                        byte[] appearanceBytes = PdfEncoding.Latin1GetBytes(appearanceContent);
                        string appearanceDictionary = PdfAcroFormDictionaryBuilder.BuildTextFieldAppearanceStreamDictionary(appearanceWidth, appearanceHeight, appearanceFontResources, appearanceBytes.Length);
                        int appearanceId = AddStreamObject(objects, appearanceDictionary, appearanceBytes);
                        formField = PdfAnnotationDictionaryBuilder.BuildChoiceFieldWidgetAnnotation(field.X1, field.Y1, field.X2, field.Y2, field.Name, field.Options, field.Values.Count == 0 ? new[] { field.Value } : field.Values, field.FontSize, appearanceId, field.IsComboBox, field.AllowsMultipleSelection, field.Style, formWidgetStructureReference?.StructParentIndex);
                    } else {
                        string appearanceContent = BuildFormFieldTextAppearanceContent(
                            appearanceWidth,
                            appearanceHeight,
                            field.Value,
                            field.FontSize,
                            field.Style,
                            pageOpts,
                            EnsureFont,
                            out IReadOnlyList<(string Name, int Id)> appearanceFontResources);
                        byte[] appearanceBytes = PdfEncoding.Latin1GetBytes(appearanceContent);
                        string appearanceDictionary = PdfAcroFormDictionaryBuilder.BuildTextFieldAppearanceStreamDictionary(appearanceWidth, appearanceHeight, appearanceFontResources, appearanceBytes.Length);
                        int appearanceId = AddStreamObject(objects, appearanceDictionary, appearanceBytes);
                        formField = PdfAnnotationDictionaryBuilder.BuildTextFieldWidgetAnnotation(field.X1, field.Y1, field.X2, field.Y2, field.Name, field.Value, field.FontSize, appearanceId, field.Style, formWidgetStructureReference?.StructParentIndex);
                    }

                    int formFieldId = AddObject(objects, formField);
                    CompleteAnnotationStructureReference(page, formWidgetStructureReference, formFieldId);
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
                    FilterPdfResources(contentStr, fontResources),
                    FilterPdfResources(contentStr, xobjects),
                    FilterPdfResources(contentStr, graphicsStates),
                    FilterPdfResources(contentStr, shadings),
                    pageAnnotIds,
                    page.StructParentIndex,
                    useStructureTabOrder: markInfo,
                    properties: page.Layers
                        .Distinct()
                        .Select(definition => ("/" + definition.ResourceName, optionalContentGroupIds[definition]))
                        .ToList()));
            pageIds.Add(pageId);
        }

        // Pages tree
        ReplaceObject(objects, pagesId, PdfPageTreeBuilder.BuildPagesDictionary(pageIds));
        if (markInfo) {
            BuildGeneratedStructTree(objects, layout.Pages, pageIds, structTreeRootId, opts.Language);
        }

        int outlinesId = BuildOutlines(objects, layout.Pages, pageIds, opts.OutlineExpansionLevelSnapshot);
        int namedDestinationsId = BuildNamedDestinations(objects, layout.Pages, pageIds);
        int acroFormId = 0;
        if (formFieldIds.Count > 0) {
            acroFormId = AddObject(objects, PdfAcroFormDictionaryBuilder.BuildAcroFormDictionary(formFieldIds, EnsureFormHelveticaFont(opts), opts.AcroFormDefaultTextAlignmentSnapshot));
        }

        int metadataId = 0;
        PdfAIdentification? pdfAIdentification = opts.PdfAIdentificationSnapshot;
        PdfUaIdentification? pdfUaIdentification = opts.PdfUaIdentificationSnapshot;
        PdfElectronicInvoiceMetadata? electronicInvoiceMetadata = opts.ElectronicInvoiceMetadataSnapshot;
        if (opts.IncludeXmpMetadata || pdfAIdentification != null || pdfUaIdentification != null || electronicInvoiceMetadata != null) {
            byte[] xmpMetadata = PdfXmpMetadataBuilder.Build(title, author, subject, keywords, pdfAIdentification, pdfUaIdentification, electronicInvoiceMetadata);
            metadataId = AddStreamObject(
                objects,
                "<< /Type /Metadata /Subtype /XML /Length " + xmpMetadata.Length.ToString(CultureInfo.InvariantCulture) + " >>",
                xmpMetadata);
        }

        int outputIntentId = 0;
        PdfOutputIntent? outputIntent = opts.OutputIntentSnapshot;
        if (outputIntent != null) {
            byte[] iccProfile = outputIntent.IccProfileSnapshot;
            int iccProfileId = AddStreamObject(objects, PdfOutputIntentDictionaryBuilder.BuildIccProfileStreamDictionary(outputIntent, iccProfile.Length), iccProfile);
            outputIntentId = AddObject(objects, PdfOutputIntentDictionaryBuilder.BuildOutputIntentObject(outputIntent, iccProfileId));
        }

        int embeddedFilesNameTreeId = 0;
        var associatedFileIds = new List<int>();
        IReadOnlyList<PdfEmbeddedFile> embeddedFiles = opts.EmbeddedFileSnapshots;
        if (embeddedFiles.Count > 0) {
            var nameTreeEntries = new List<(string FileName, int FileSpecId)>(embeddedFiles.Count);
            foreach (PdfEmbeddedFile embeddedFile in embeddedFiles.OrderBy(file => file.FileName, StringComparer.Ordinal)) {
                byte[] fileBytes = embeddedFile.DataSnapshot;
                int embeddedFileId = AddStreamObject(
                    objects,
                    PdfEmbeddedFileDictionaryBuilder.BuildEmbeddedFileStreamDictionary(embeddedFile, fileBytes),
                    fileBytes);
                int fileSpecId = AddObject(objects, PdfEmbeddedFileDictionaryBuilder.BuildFileSpecificationObject(embeddedFile, embeddedFileId));
                nameTreeEntries.Add((embeddedFile.FileName, fileSpecId));
                associatedFileIds.Add(fileSpecId);
            }

            embeddedFilesNameTreeId = AddObject(objects, PdfEmbeddedFileDictionaryBuilder.BuildEmbeddedFilesNameTree(nameTreeEntries));
        }

        int portfolioId = 0;
        PdfPortfolioOptions? portfolio = opts.PortfolioSnapshot;
        if (portfolio != null) {
            portfolioId = AddObject(objects, PdfPortfolioDictionaryBuilder.Build(portfolio, embeddedFiles));
        }

        int pageLabelsId = 0;
        if (opts.IncludePageLabels) {
            IReadOnlyList<PdfPageLabelRange> pageLabelRanges = opts.PageLabelRangeSnapshots;
            if (pageLabelRanges.Count > 0) {
                ValidatePageLabelRanges(pageLabelRanges, layout.Pages.Count);
                pageLabelsId = AddObject(objects, PdfPageLabelDictionaryBuilder.BuildGeneratedPageLabelsDictionary(pageLabelRanges));
            } else {
                pageLabelsId = AddObject(objects, PdfPageLabelDictionaryBuilder.BuildGeneratedPageLabelsDictionary(
                    opts.PageNumberStyle,
                    opts.PageNumberStart,
                    opts.PageLabelPrefix));
            }
        }

        int viewerPreferencesId = 0;
        PdfViewerPreferencesOptions? viewerPreferences = opts.ViewerPreferencesSnapshot;
        if (viewerPreferences != null && viewerPreferences.HasAny) {
            viewerPreferencesId = AddObject(objects, PdfViewerPreferenceDictionaryBuilder.BuildGeneratedViewerPreferencesDictionary(viewerPreferences, layout.Pages.Count));
        }

        string? openAction = null;
        PdfOpenActionOptions? openActionOptions = opts.OpenActionSnapshot;
        if (openActionOptions != null) {
            ValidateOpenAction(openActionOptions, layout.Pages.Count);
            int targetPageIndex = openActionOptions.PageNumber - 1;
            var destination = ResolveOpenActionDestinationCoordinates(openActionOptions, layout.Pages[targetPageIndex]);
            openAction = PdfCatalogDictionaryBuilder.BuildGeneratedOpenActionDestination(
                pageIds[targetPageIndex],
                destination.Top,
                openActionOptions.DestinationMode,
                destination.Left,
                destination.Bottom,
                destination.Right);
        }

        string? pageMode = opts.CatalogPageModeSnapshot.HasValue
            ? PdfCatalogDictionaryBuilder.GetPageModeName(opts.CatalogPageModeSnapshot.Value)
            : null;
        string? pageLayout = opts.CatalogPageLayoutSnapshot.HasValue
            ? PdfCatalogDictionaryBuilder.GetPageLayoutName(opts.CatalogPageLayoutSnapshot.Value)
            : null;

        // Catalog
        catalogId = AddObject(objects, PdfCatalogDictionaryBuilder.BuildGeneratedCatalogDictionary(
            pagesId,
            outlinesId,
            namedDestinationsId,
            acroFormId,
            metadataId,
            outputIntentId,
            opts.Language,
            embeddedFilesNameTreeId,
            associatedFileIds,
            pageLabelsId,
            viewerPreferencesId,
            structTreeRootId,
            markInfo,
            openAction,
            pageMode,
            pageLayout,
            opts.CatalogUriBaseSnapshot,
            portfolioId,
            optionalContentPropertiesId));

        infoId = AddObject(objects, PdfInfoDictionaryBuilder.Build(title, author, subject, keywords));
        MaterializePendingFontObjects();

        PdfFileVersion effectiveFileVersion = requiresPdf16FileVersion
            ? PdfFileAssembler.RequireAtLeast(opts.FileVersion, PdfFileVersion.Pdf16)
            : opts.FileVersion;
        if (portfolioId > 0) {
            effectiveFileVersion = PdfFileAssembler.RequireAtLeast(effectiveFileVersion, PdfFileVersion.Pdf17);
        }
        if (optionalContentPropertiesId > 0) {
            effectiveFileVersion = PdfFileAssembler.RequireAtLeast(effectiveFileVersion, PdfFileVersion.Pdf15);
        }
        if (outputStream != null) {
            bytesWritten = PdfFileAssembler.Assemble(outputStream, objects, catalogId, infoId, effectiveFileVersion, opts.EncryptionSnapshot, opts.ObjectBufferMemoryLimitBytes);
            return null;
        }

        byte[] bytes = PdfFileAssembler.Assemble(objects, catalogId, infoId, effectiveFileVersion, opts.EncryptionSnapshot, opts.ObjectBufferMemoryLimitBytes);
        bytesWritten = bytes.LongLength;
        return bytes;
    }

    private static string ReplaceInlineImageDrawTokens(string content, IReadOnlyList<PageImage> images) {
        if (string.IsNullOrEmpty(content) || images.Count == 0) {
            return content;
        }

        string result = content;
        foreach (PageImage image in images) {
            if (string.IsNullOrEmpty(image.InlineDrawToken)) {
                continue;
            }

            var imageDraw = new StringBuilder();
            AppendPageImageDraw(imageDraw, image);
            result = result.Replace(image.InlineDrawToken!, imageDraw.ToString());
        }

        return result;
    }

    private static string ReplaceInlineEffectGroupTokens(string content, IReadOnlyList<PageEffectGroup> effects, int availableCount) {
        if (string.IsNullOrEmpty(content) || effects.Count == 0 || availableCount <= 0) return content;
        string result = content;
        int count = Math.Min(availableCount, effects.Count);
        for (int index = 0; index < count; index++) {
            PageEffectGroup effect = effects[index];
            if (effect.ObjectId <= 0 || string.IsNullOrEmpty(effect.Name) || string.IsNullOrEmpty(effect.Token)) continue;
            var invocation = new StringBuilder();
            var stream = new ContentStreamBuilder(invocation).SaveState();
            if (!string.IsNullOrEmpty(effect.GraphicsStateName)) stream.GraphicsState(effect.GraphicsStateName!);
            stream.TransformMatrix(effect.Transform)
                .XObject(effect.Name)
                .RestoreState();
            result = result.Replace(effect.Token, invocation.ToString());
        }
        return result;
    }

    private static List<(string Name, int Id)> FilterPdfResources(
        string content,
        List<(string Name, int Id)> resources) {
        var used = new List<(string Name, int Id)>();
        if (string.IsNullOrEmpty(content) || resources.Count == 0) return used;
        for (int index = 0; index < resources.Count; index++) {
            if (UsesPdfResource(content, resources[index].Name)) used.Add(resources[index]);
        }
        return used;
    }

    private static bool UsesPdfResource(string content, string name) {
        int searchIndex = 0;
        while (searchIndex < content.Length) {
            int index = content.IndexOf(name, searchIndex, StringComparison.Ordinal);
            if (index < 0) return false;
            int next = index + name.Length;
            if (next >= content.Length || char.IsWhiteSpace(content[next])) return true;
            searchIndex = next;
        }
        return false;
    }

    private static string BuildPageBackground(LayoutResult.Page page, PdfOptions options, string pageBackgroundShapeContent, PdfTextWatermark? watermark, string? watermarkFontAlias, System.Collections.Generic.IReadOnlyDictionary<PdfStandardFont, string> fontResources, string? textWatermarkGraphicsStateName, PdfPageBorder? pageBorder, string? pageBorderGraphicsStateName, bool markDecorativeArtifacts) {
        var sb = new StringBuilder();
        if (options.BackgroundColor.HasValue) {
            var backgroundColor = new StringBuilder();
            new ContentStreamBuilder(backgroundColor)
                .SaveState()
                .FillColor(options.BackgroundColor.Value)
                .Rectangle(0, 0, options.PageWidth, options.PageHeight)
                .FillPath()
                .RestoreState();
            sb.Append(WrapArtifactContent(backgroundColor.ToString(), markDecorativeArtifacts));
        }

        sb.Append(WrapArtifactContent(pageBackgroundShapeContent, markDecorativeArtifacts));

        foreach (PageImage image in page.Images) {
            if (image.IsBackgroundDecoration) {
                AppendPageImageDraw(sb, image);
            }
        }

        if (watermark != null && watermark.Opacity > 0D && !string.IsNullOrEmpty(watermarkFontAlias)) {
            var watermarkContent = new StringBuilder();
            AppendTextWatermark(watermarkContent, options, watermark, watermarkFontAlias!, fontResources, textWatermarkGraphicsStateName);
            sb.Append(WrapArtifactContent(watermarkContent.ToString(), markDecorativeArtifacts));
        }

        if (pageBorder != null && pageBorder.Opacity > 0D) {
            var pageBorderContent = new StringBuilder();
            AppendPageBorder(pageBorderContent, options, pageBorder, pageBorderGraphicsStateName);
            sb.Append(WrapArtifactContent(pageBorderContent.ToString(), markDecorativeArtifacts));
        }

        return sb.ToString();
    }

    private static string WrapArtifactContent(string content, bool enabled) {
        if (!enabled || string.IsNullOrEmpty(content)) {
            return content;
        }

        return "/Artifact BMC\n" + content + "EMC\n";
    }

    private static void AssignFigureMarkedContentIds(LayoutResult.Page page) {
        foreach (PageImage image in page.Images) {
            if (image.SuppressAccessibilityWrapper || image.IsBackgroundDecoration || string.IsNullOrWhiteSpace(image.AlternativeText) || image.MarkedContentId.HasValue || image.StructElementIndex.HasValue) {
                continue;
            }

            int markedContentId = page.NextMarkedContentId++;
            int structElementIndex = page.StructElements.Count;
            image.MarkedContentId = markedContentId;
            image.StructElementIndex = structElementIndex;
            page.StructElements.Add(new PageStructElement {
                MarkedContentId = markedContentId,
                StructureType = "Figure",
                AlternativeText = image.AlternativeText!
            });
        }
    }

    private static void AttachAnnotationToStructElement(PageStructElement structElement, int annotationObjectId, int annotationStructParentIndex) {
        if (!structElement.AnnotationObjectId.HasValue) {
            structElement.AnnotationObjectId = annotationObjectId;
            structElement.AnnotationStructParentIndex = annotationStructParentIndex;
            return;
        }

        (structElement.AdditionalAnnotationObjectIds ??= new System.Collections.Generic.List<int>()).Add(annotationObjectId);
        (structElement.AdditionalAnnotationStructParentIndexes ??= new System.Collections.Generic.List<int>()).Add(annotationStructParentIndex);
    }

    private static void AssignStructParentIndex(LayoutResult.Page page, ref int nextStructParentIndex) {
        if (page.StructElements.Count > 0 && !page.StructParentIndex.HasValue) {
            page.StructParentIndex = nextStructParentIndex++;
        }
    }

    private static AnnotationStructureReference? RegisterAnnotationStructureReference(LayoutResult.Page page, bool markInfo, ref int nextStructParentIndex, string structureType) {
        if (!markInfo) {
            return null;
        }

        var reference = new AnnotationStructureReference {
            StructParentIndex = nextStructParentIndex++,
            StructElementIndex = page.StructElements.Count
        };
        page.StructElements.Add(new PageStructElement {
            StructureType = structureType,
            AnnotationStructParentIndex = reference.StructParentIndex
        });
        return reference;
    }

    private static void CompleteAnnotationStructureReference(LayoutResult.Page page, AnnotationStructureReference? reference, int annotationObjectId) {
        if (reference == null) {
            return;
        }

        reference.ObjectId = annotationObjectId;
        if (reference.StructElementIndex >= 0 && reference.StructElementIndex < page.StructElements.Count) {
            page.StructElements[reference.StructElementIndex].AnnotationObjectId = annotationObjectId;
        }
    }

    private static void BuildGeneratedStructTree(IList<byte[]> objects, IReadOnlyList<LayoutResult.Page> pages, List<int> pageIds, int structTreeRootId, string? documentLanguage) {
        if (!pages.Any(page => page.StructElements.Count > 0)) {
            ReplaceObject(objects, structTreeRootId, PdfStructTreeRootDictionaryBuilder.BuildEmptyStructTreeRootDictionary());
            return;
        }

        int documentStructElementId = ReserveObject(objects);
        var documentChildElementIds = new List<int>();
        var parentTreeEntries = new List<PdfStructTreeRootDictionaryBuilder.ParentTreeEntry>();
        for (int pageIndex = 0; pageIndex < pages.Count; pageIndex++) {
            LayoutResult.Page page = pages[pageIndex];
            for (int elementIndex = 0; elementIndex < page.StructElements.Count; elementIndex++) {
                page.StructElements[elementIndex].ObjectId = ReserveObject(objects);
            }
        }

        for (int pageIndex = 0; pageIndex < pages.Count; pageIndex++) {
            LayoutResult.Page page = pages[pageIndex];
            if (page.StructElements.Count == 0) {
                continue;
            }

            for (int elementIndex = 0; elementIndex < page.StructElements.Count; elementIndex++) {
                PageStructElement element = page.StructElements[elementIndex];
                int parentObjectId = element.ParentElement != null
                    ? element.ParentElement.ObjectId
                    : element.ParentElementIndex.HasValue &&
                    element.ParentElementIndex.Value >= 0 &&
                    element.ParentElementIndex.Value < page.StructElements.Count
                        ? page.StructElements[element.ParentElementIndex.Value].ObjectId
                        : documentStructElementId;
                string structElement;
                if (element.AnnotationObjectId.HasValue) {
                    structElement = PdfStructTreeRootDictionaryBuilder.BuildAnnotationStructElement(
                        parentObjectId,
                        pageIds[pageIndex],
                        element.AnnotationObjectId.Value,
                        element.MarkedContentId,
                        element.AdditionalMarkedContentIds,
                        element.AdditionalAnnotationObjectIds,
                        element.StructureType,
                        element.AlternativeText);
                } else if (element.MarkedContentId.HasValue) {
                    structElement = string.Equals(element.StructureType, "Figure", StringComparison.Ordinal)
                        ? PdfStructTreeRootDictionaryBuilder.BuildFigureStructElement(
                            parentObjectId,
                            pageIds[pageIndex],
                            element.MarkedContentId.Value,
                            element.AlternativeText)
                        : PdfStructTreeRootDictionaryBuilder.BuildTextStructElement(
                            parentObjectId,
                            pageIds[pageIndex],
                            element.StructureType,
                            element.MarkedContentId.Value,
                            element.TableHeaderScope,
                            element.TableColumnSpan,
                            element.TableRowSpan,
                            element.AdditionalMarkedContentIds);
                } else {
                    var elementChildIds = new List<int>();
                    for (int childIndex = 0; childIndex < page.StructElements.Count; childIndex++) {
                        if (page.StructElements[childIndex].ParentElementIndex == elementIndex) {
                            elementChildIds.Add(page.StructElements[childIndex].ObjectId);
                        }
                    }

                    for (int childPageIndex = 0; childPageIndex < pages.Count; childPageIndex++) {
                        LayoutResult.Page childPage = pages[childPageIndex];
                        for (int childIndex = 0; childIndex < childPage.StructElements.Count; childIndex++) {
                            if (ReferenceEquals(childPage.StructElements[childIndex].ParentElement, element)) {
                                elementChildIds.Add(childPage.StructElements[childIndex].ObjectId);
                            }
                        }
                    }

                    structElement = PdfStructTreeRootDictionaryBuilder.BuildContainerStructElement(
                        parentObjectId,
                        pageIds[pageIndex],
                        element.StructureType,
                        elementChildIds,
                        element.TableHeaderScope,
                        element.TableColumnSpan,
                        element.TableRowSpan,
                        element.AlternativeText);
                }

                ReplaceObject(objects, element.ObjectId, structElement);
            }

            var pageMarkedContentElements = new List<(int MarkedContentId, int ObjectId)>();
            foreach (PageStructElement element in page.StructElements.Where(element => element.MarkedContentId.HasValue)) {
                pageMarkedContentElements.Add((element.MarkedContentId!.Value, element.ObjectId));
                if (element.AdditionalMarkedContentIds != null) {
                    for (int additionalIndex = 0; additionalIndex < element.AdditionalMarkedContentIds.Count; additionalIndex++) {
                        pageMarkedContentElements.Add((element.AdditionalMarkedContentIds[additionalIndex], element.ObjectId));
                    }
                }
            }

            var pageElementIds = new List<int>();
            foreach ((int MarkedContentId, int ObjectId) mapping in pageMarkedContentElements.OrderBy(mapping => mapping.MarkedContentId)) {
                pageElementIds.Add(mapping.ObjectId);
            }

            for (int elementIndex = 0; elementIndex < page.StructElements.Count; elementIndex++) {
                PageStructElement element = page.StructElements[elementIndex];
                if (!element.ParentElementIndex.HasValue && element.ParentElement == null) {
                    documentChildElementIds.Add(element.ObjectId);
                }
            }

            if (page.StructParentIndex.HasValue && pageElementIds.Count > 0) {
                parentTreeEntries.Add(PdfStructTreeRootDictionaryBuilder.ParentTreeEntry.ForMarkedContentPage(page.StructParentIndex.Value, pageElementIds));
            }

            foreach (PageStructElement element in page.StructElements.Where(element => element.AnnotationObjectId.HasValue && element.AnnotationStructParentIndex.HasValue).OrderBy(element => element.AnnotationStructParentIndex!.Value)) {
                parentTreeEntries.Add(PdfStructTreeRootDictionaryBuilder.ParentTreeEntry.ForObjectReference(element.AnnotationStructParentIndex!.Value, element.ObjectId));
                if (element.AdditionalAnnotationStructParentIndexes != null) {
                    for (int additionalIndex = 0; additionalIndex < element.AdditionalAnnotationStructParentIndexes.Count; additionalIndex++) {
                        parentTreeEntries.Add(PdfStructTreeRootDictionaryBuilder.ParentTreeEntry.ForObjectReference(element.AdditionalAnnotationStructParentIndexes[additionalIndex], element.ObjectId));
                    }
                }
            }
        }

        if (documentChildElementIds.Count == 0) {
            ReplaceObject(objects, structTreeRootId, PdfStructTreeRootDictionaryBuilder.BuildEmptyStructTreeRootDictionary());
            ReplaceObject(objects, documentStructElementId, PdfStructTreeRootDictionaryBuilder.BuildDocumentStructElement(structTreeRootId, documentChildElementIds, documentLanguage));
            return;
        }

        ReplaceObject(objects, documentStructElementId, PdfStructTreeRootDictionaryBuilder.BuildDocumentStructElement(structTreeRootId, documentChildElementIds, documentLanguage));
        int parentTreeId = AddObject(objects, PdfStructTreeRootDictionaryBuilder.BuildParentTree(parentTreeEntries));
        int parentTreeNextKey = parentTreeEntries.Count == 0
            ? 0
            : parentTreeEntries.Max(entry => entry.StructParentIndex) + 1;
        ReplaceObject(objects, structTreeRootId, PdfStructTreeRootDictionaryBuilder.BuildStructTreeRootDictionary(new[] { documentStructElementId }, parentTreeId, parentTreeNextKey));
    }

    private static string BuildPageBackgroundShapes(LayoutResult.Page page, System.Collections.Generic.IReadOnlyList<PdfPageBackgroundShape> shapes) {
        if (shapes.Count == 0) {
            return string.Empty;
        }

        var sb = new StringBuilder();
        foreach (PdfPageBackgroundShape backgroundShape in shapes) {
            OfficeIMO.Drawing.OfficeShape shape = backgroundShape.Shape;
            DrawHeaderFooterShapeGeometryAt(sb, page, shape, backgroundShape.X, backgroundShape.Y);
        }

        return sb.ToString();
    }

    private static void AppendPageBorder(StringBuilder sb, PdfOptions options, PdfPageBorder border, string? graphicsStateName) {
        double inset = border.Inset;
        double pathX = inset;
        double pathY = inset;
        double pathWidth = options.PageWidth - (inset * 2D);
        double pathHeight = options.PageHeight - (inset * 2D);
        if (pathWidth <= 0D || pathHeight <= 0D || double.IsNaN(pathWidth) || double.IsInfinity(pathWidth) || double.IsNaN(pathHeight) || double.IsInfinity(pathHeight)) {
            throw new ArgumentException("PDF page border inset must leave a positive border rectangle.");
        }

        var content = new ContentStreamBuilder(sb)
            .SaveState();
        if (!string.IsNullOrEmpty(graphicsStateName)) {
            content.GraphicsState(graphicsStateName!);
        }

        content
            .StrokeColor(border.Color)
            .LineWidth(border.Width);
        ApplyStrokeStyle(content, border.DashStyle, border.Width, strokeLineCap: null, strokeLineJoin: null);
        content
            .Rectangle(pathX, pathY, pathWidth, pathHeight)
            .StrokePath()
            .RestoreState();
    }

    private static void AddPageBackgroundImage(LayoutResult.Page page, PdfOptions options, PdfPageBackgroundImage image) {
        double imageWidth = image.ImageInfo.Width > 0 ? image.ImageInfo.Width : options.PageWidth;
        double imageHeight = image.ImageInfo.Height > 0 ? image.ImageInfo.Height : options.PageHeight;
        OfficeImageRenderPlan renderPlan = OfficeImageRenderPlan.CreateBottomLeft(
            imageWidth,
            imageHeight,
            0D,
            0D,
            options.PageWidth,
            options.PageHeight,
            image.Fit);
        string? stateName = image.Opacity < 1D
            ? EnsureHeaderFooterGraphicsState(page, image.Opacity, image.Opacity)
            : null;

        page.Images.Add(new PageImage {
            Data = image.DataSnapshot,
            Info = image.ImageInfo,
            X = renderPlan.ImagePlacement.X,
            Y = renderPlan.ImagePlacement.Y,
            W = renderPlan.ImagePlacement.Width,
            H = renderPlan.ImagePlacement.Height,
            IsBackgroundDecoration = true,
            Opacity = image.Opacity,
            GraphicsStateName = stateName
        });
    }

    private static void AddImageWatermark(LayoutResult.Page page, PdfOptions options, PdfImageWatermark watermark) {
        double x = (options.PageWidth - watermark.Width) / 2D;
        double y = (options.PageHeight - watermark.Height) / 2D;
        string? stateName = watermark.Opacity < 1D
            ? EnsureHeaderFooterGraphicsState(page, watermark.Opacity, watermark.Opacity)
            : null;

        page.Images.Add(new PageImage {
            Data = watermark.DataSnapshot,
            Info = watermark.ImageInfo,
            X = x,
            Y = y,
            W = watermark.Width,
            H = watermark.Height,
            IsBackgroundDecoration = true,
            Opacity = watermark.Opacity,
            RotationAngle = watermark.RotationAngle,
            GraphicsStateName = stateName
        });
    }

    private static void AppendPageImageDraw(StringBuilder sb, PageImage img) {
        if (img.ClipPath != null) {
            new ContentStreamBuilder(sb)
                .SaveState();
            AppendClipPath(sb, img.ClipPath, img.ClipX, img.ClipY, img.ClipHeight);
        }

        bool hasAlternativeText = !img.SuppressAccessibilityWrapper && !string.IsNullOrWhiteSpace(img.AlternativeText);
        if (hasAlternativeText) {
            sb.Append("/Figure << /Alt ")
                .Append(PdfSyntaxEscaper.TextString(img.AlternativeText!));
            if (img.MarkedContentId.HasValue) {
                sb.Append(" /MCID ")
                    .Append(img.MarkedContentId.Value.ToString(CultureInfo.InvariantCulture));
            }

            sb.Append(" >> BDC\n");
        } else if (!img.SuppressAccessibilityWrapper && img.IsDecorativeArtifact) {
            sb.Append("/Artifact BMC\n");
        }

        OfficeTransform imageTransform = new OfficeImageProjection(
            new OfficeImagePlacement(img.X, img.Y, img.W, img.H),
            rotationDegrees: img.RotationAngle,
            flipHorizontal: img.HorizontalFlip,
            flipVertical: img.VerticalFlip)
            .CreateUnitSquareTransform();

        var content = new ContentStreamBuilder(sb)
            .SaveState();
        if (!string.IsNullOrEmpty(img.GraphicsStateName)) {
            content.GraphicsState(img.GraphicsStateName!);
        }

        content
            .TransformMatrix(imageTransform);
        if (img.SourceCrop?.HasCrop == true) {
            double clipWidth = 1D - img.SourceCrop.Left - img.SourceCrop.Right;
            double clipHeight = 1D - img.SourceCrop.Top - img.SourceCrop.Bottom;
            content.Rectangle(img.SourceCrop.Left, img.SourceCrop.Bottom, clipWidth, clipHeight)
                .ClipPath()
                .EndPath();
        }

        content.XObject(img.Name)
            .RestoreState();

        if (hasAlternativeText || !img.SuppressAccessibilityWrapper && img.IsDecorativeArtifact) {
            sb.Append("EMC\n");
        }

        if (img.ClipPath != null) {
            new ContentStreamBuilder(sb)
                .RestoreState();
        }
    }

    private static void EnsureTextWatermarkFontResources(PdfTextWatermark watermark, PdfOptions options, Func<PdfStandardFont, string, string> ensureFontResource) {
        PdfStandardFont baseFont = ChooseNormal(watermark.Font);
        PdfStandardFont normalFont = ChooseNormal(options.DefaultFont);
        System.Collections.Generic.IReadOnlyList<TextRun> runs = BuildTextWatermarkRuns(watermark, options);
        foreach (TextRun run in runs) {
            PdfStandardFont runFont = ResolvePageTextRunFont(run, baseFont);
            ensureFontResource(runFont, GetStandardFontResourceName(runFont, normalFont));
        }
    }

    private static void AppendTextWatermark(StringBuilder sb, PdfOptions options, PdfTextWatermark watermark, string fontAlias, System.Collections.Generic.IReadOnlyDictionary<PdfStandardFont, string> fontResources, string? graphicsStateName) {
        PdfStandardFont baseFont = ChooseNormal(watermark.Font);
        System.Collections.Generic.IReadOnlyList<TextRun> runs = BuildTextWatermarkRuns(watermark, options);
        double textWidth = MeasureTextWatermarkRuns(runs, baseFont, watermark.FontSize, options);
        double angle = watermark.RotationAngle * System.Math.PI / 180D;
        double cos = System.Math.Cos(angle);
        double sin = System.Math.Sin(angle);
        double centerX = options.PageWidth / 2D;
        double centerY = options.PageHeight / 2D;
        double originX = centerX - cos * textWidth / 2D + sin * watermark.FontSize / 2D;
        double originY = centerY - sin * textWidth / 2D - cos * watermark.FontSize / 2D;

        var content = new ContentStreamBuilder(sb)
            .SaveState();
        if (!string.IsNullOrEmpty(graphicsStateName)) {
            content.GraphicsState(graphicsStateName!);
        }

        content
            .BeginText()
            .Font(fontAlias, watermark.FontSize)
            .FillColor(watermark.Color)
            .TextMatrix(cos, sin, -sin, cos, originX, originY);
        foreach (TextRun run in runs) {
            string text = run.Text ?? string.Empty;
            if (text.Length == 0) {
                continue;
            }

            PdfStandardFont runFont = ResolvePageTextRunFont(run, baseFont);
            string runFontResource = ResolvePageTextFontResource(fontResources, runFont);
            double runFontSize = run.FontSize ?? watermark.FontSize;
            content
                .Font(runFontResource, runFontSize)
                .ShowText(EncodeTextShowCommand(text, runFont, options), runFontSize);
        }

        content.EndText()
            .RestoreState();
    }

    private static System.Collections.Generic.IReadOnlyList<TextRun> BuildTextWatermarkRuns(PdfTextWatermark watermark, PdfOptions options) {
        PdfStandardFont baseFont = ChooseNormal(watermark.Font);
        var run = new TextRun(
            watermark.Text,
            bold: watermark.Bold,
            underline: false,
            color: watermark.Color,
            italic: watermark.Italic,
            strike: false,
            fontSize: watermark.FontSize,
            font: baseFont);
        return NormalizeFallbackRuns(new[] { run }, baseFont, options);
    }

    private static double MeasureTextWatermarkRuns(System.Collections.Generic.IReadOnlyList<TextRun> runs, PdfStandardFont baseFont, double fontSize, PdfOptions options) {
        double width = 0D;
        foreach (TextRun run in runs) {
            width += MeasureRichText(run.Text ?? string.Empty, ResolvePageTextRunFont(run, baseFont), run.FontSize ?? fontSize, run.Baseline, options);
        }

        return width;
    }

    private static PdfStandardFont GetTextWatermarkFont(PdfTextWatermark watermark) {
        PdfStandardFont normal = ChooseNormal(watermark.Font);
        if (watermark.Bold && watermark.Italic) {
            return ChooseBoldItalic(normal);
        }

        if (watermark.Bold) {
            return ChooseBold(normal);
        }

        if (watermark.Italic) {
            return ChooseItalic(normal);
        }

        return normal;
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

    private static void ValidateGeneratedFormFieldNames(IReadOnlyList<LayoutResult.Page> pages) {
        var names = new HashSet<string>(StringComparer.Ordinal);
        foreach (var page in pages) {
            foreach (var field in page.FormFields) {
                if (!names.Add(field.Name)) {
                    throw new ArgumentException("PDF generated form field names must be unique: " + field.Name);
                }
            }
        }
    }

    private static int BuildOutlines(IList<byte[]> objects, IReadOnlyList<LayoutResult.Page> pages, List<int> pageIds, int outlineExpansionLevel) {
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
            int descendantCount = IsOutlineExpanded(node, outlineExpansionLevel)
                ? CountVisibleOutlines(node.Children, outlineExpansionLevel)
                : -CountOutlines(node.Children);
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
            CountVisibleOutlines(root.Children, outlineExpansionLevel)));

        return rootId;
    }

    private static int BuildNamedDestinations(IList<byte[]> objects, IReadOnlyList<LayoutResult.Page> pages, List<int> pageIds) {
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

    private static int CountVisibleOutlines(IEnumerable<OutlineNode> nodes, int outlineExpansionLevel) {
        int count = 0;
        foreach (var node in nodes) {
            count++;
            if (IsOutlineExpanded(node, outlineExpansionLevel)) {
                count += CountVisibleOutlines(node.Children, outlineExpansionLevel);
            }
        }

        return count;
    }

    private static bool IsOutlineExpanded(OutlineNode node, int outlineExpansionLevel) =>
        node.Children.Count > 0 && node.Level <= outlineExpansionLevel;

}

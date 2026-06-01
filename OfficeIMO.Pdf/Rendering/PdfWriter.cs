using System.Globalization;

namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    public static byte[] Write(PdfDoc doc, IEnumerable<IPdfBlock> blocks, PdfOptions opts, string? title, string? author, string? subject, string? keywords) {
        PdfComplianceValidator.ValidateGenerationOptions(opts);

        // Layout blocks into pages and create per-page content streams.
        var layout = LayoutBlocks(blocks, opts);
        ValidateNamedDestinationLinks(layout.Pages);
        ValidateGeneratedFormFieldNames(layout.Pages);

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
                if (opts.TryGetEmbeddedStandardFont(font, out PdfEmbeddedFont? embeddedFont) && embeddedFont != null) {
                    PdfTrueTypeFontProgram fontProgram = PdfTrueTypeFontProgram.Parse(embeddedFont.DataSnapshot, embeddedFont.FontName);
                    byte[] fontData = embeddedFont.DataSnapshot;
                    string fontFileExtraEntries = "/Length1 " + fontData.Length.ToString(CultureInfo.InvariantCulture);
                    int fontFileId = opts.CompressEmbeddedFonts
                        ? AddFlateStreamObject(objects, fontData, fontFileExtraEntries)
                        : AddStreamObject(
                            objects,
                            "<< /Length " + fontData.Length.ToString(CultureInfo.InvariantCulture) + " " + fontFileExtraEntries + " >>",
                            fontData);
                    int descriptorId = AddObject(objects, PdfStandardFontDictionaryBuilder.BuildTrueTypeFontDescriptorObject(fontProgram, fontFileId));
                    int toUnicodeObjectId = AddStreamObject(objects, PdfToUnicodeCMapBuilder.BuildWinAnsiToUnicodeCMap());
                    id = AddObject(objects, PdfStandardFontDictionaryBuilder.BuildEmbeddedTrueTypeFontObject(fontProgram, descriptorId, toUnicodeObjectId));
                } else {
                    int toUnicodeObjectId = opts.IncludeStandardFontToUnicodeMaps
                        ? AddStreamObject(objects, PdfToUnicodeCMapBuilder.BuildWinAnsiToUnicodeCMap())
                        : 0;
                    id = AddObject(objects, PdfStandardFontDictionaryBuilder.BuildStandardType1FontObject(font, toUnicodeObjectId));
                }

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
            foreach (PdfStandardFont usedFont in page.UsedFonts) {
                EnsurePageFontResource(usedFont, GetStandardFontResourceName(usedFont, normalFont));
            }
            PdfTextWatermark? textWatermark = pageOpts.TextWatermarkSnapshot;
            string? watermarkFontAlias = null;
            string? textWatermarkGraphicsStateName = null;
            if (textWatermark != null && textWatermark.Opacity > 0D) {
                watermarkFontAlias = EnsurePageFontResource(GetTextWatermarkFont(textWatermark), "FW");
                if (textWatermark.Opacity < 1D) {
                    textWatermarkGraphicsStateName = EnsureHeaderFooterGraphicsState(page, textWatermark.Opacity, textWatermark.Opacity);
                }
            }
            PdfPageBackgroundImage? pageBackgroundImage = pageOpts.PageBackgroundImageSnapshot;
            if (pageBackgroundImage != null && pageBackgroundImage.Opacity > 0D) {
                AddPageBackgroundImage(page, pageOpts, pageBackgroundImage);
            }
            PdfImageWatermark? imageWatermark = pageOpts.ImageWatermarkSnapshot;
            if (imageWatermark != null && imageWatermark.Opacity > 0D) {
                AddImageWatermark(page, pageOpts, imageWatermark);
            }
            PdfPageBorder? pageBorder = pageOpts.PageBorderSnapshot;
            string? pageBorderGraphicsStateName = null;
            if (pageBorder != null && pageBorder.Opacity > 0D && pageBorder.Opacity < 1D) {
                pageBorderGraphicsStateName = EnsureHeaderFooterGraphicsState(page, 1D, pageBorder.Opacity);
            }
            string pageBackgroundShapeContent = BuildPageBackgroundShapes(page, pageOpts.PageBackgroundShapeSnapshots);
            string? headerFontAlias = null;
            if (pageOpts.HasHeaderTextContentForPage(headerFooterVariantPageNumber)) {
                headerFontAlias = EnsurePageFontResource(pageOpts.HeaderFont, "F5");
            }
            string? footerFontAlias = null;
            if (pageOpts.HasFooterTextContentForPage(headerFooterVariantPageNumber)) {
                footerFontAlias = EnsurePageFontResource(pageOpts.FooterFont, "F6");
            }

            string headerFooterShapeContent = BuildHeaderFooterShapes(page, pageOpts, headerFooterVariantPageNumber);

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
            AddHeaderFooterImages(page, pageOpts, headerFooterVariantPageNumber);
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
            }

            string pageBackgroundContent = BuildPageBackground(page, pageOpts, pageBackgroundShapeContent, textWatermark, watermarkFontAlias, textWatermarkGraphicsStateName, pageBorder, pageBorderGraphicsStateName);
            string contentStr = pageBackgroundContent + headerFooterShapeContent;
            if (pageOpts.HasHeaderTextContentForPage(headerFooterVariantPageNumber)) {
                string headerContent = BuildHeader(pageOpts, headerFooterVariantPageNumber, headerFooterPageNumber, headerFooterTotalPages, totalPages, pageOpts.HeaderFont, headerFontAlias!);
                contentStr += headerContent;
            }
            contentStr += page.Content;
            if (page.Images.Count > 0) {
                var sbImgs = new StringBuilder();
                foreach (var img in page.Images) {
                    if (img.IsBackgroundDecoration) {
                        continue;
                    }

                    AppendPageImageDraw(sbImgs, img);
                }

                contentStr += sbImgs.ToString();
            }
            if (pageOpts.HasFooterTextContentForPage(headerFooterVariantPageNumber)) {
                string footer = BuildFooter(pageOpts, headerFooterVariantPageNumber, headerFooterPageNumber, headerFooterTotalPages, totalPages, pageOpts.FooterFont, footerFontAlias!);
                contentStr += footer;
            }
            byte[] contentBytes = Encoding.ASCII.GetBytes(contentStr);
            int contentId = pageOpts.CompressContentStreams
                ? AddFlateStreamObject(objects, contentBytes)
                : AddStreamObject(objects, contentBytes);
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
                                field.Style);
                            int widgetObjectId = AddObject(objects, widget);
                            widgetObjectIds.Add(widgetObjectId);
                            pageAnnotIds.Add(widgetObjectId);
                        }

                        ReplaceObject(objects, parentFieldId, PdfAnnotationDictionaryBuilder.BuildRadioButtonFieldDictionary(field.Name, field.Options, field.Value, widgetObjectIds));
                        formFieldIds.Add(parentFieldId);
                        continue;
                    }

                    if (field.Kind == FormFieldAnnotationKind.CheckBox) {
                        string offAppearance = PdfAcroFormDictionaryBuilder.BuildCheckBoxAppearanceContent(appearanceWidth, appearanceHeight, selected: false, field.Style);
                        byte[] offAppearanceBytes = PdfEncoding.Latin1GetBytes(offAppearance);
                        string offAppearanceDictionary = PdfAcroFormDictionaryBuilder.BuildCheckBoxAppearanceStreamDictionary(appearanceWidth, appearanceHeight, offAppearanceBytes.Length);
                        int offAppearanceId = AddStreamObject(objects, offAppearanceDictionary, offAppearanceBytes);

                        string checkedAppearance = PdfAcroFormDictionaryBuilder.BuildCheckBoxAppearanceContent(appearanceWidth, appearanceHeight, selected: true, field.Style);
                        byte[] checkedAppearanceBytes = PdfEncoding.Latin1GetBytes(checkedAppearance);
                        string checkedAppearanceDictionary = PdfAcroFormDictionaryBuilder.BuildCheckBoxAppearanceStreamDictionary(appearanceWidth, appearanceHeight, checkedAppearanceBytes.Length);
                        int checkedAppearanceId = AddStreamObject(objects, checkedAppearanceDictionary, checkedAppearanceBytes);

                        formField = PdfAnnotationDictionaryBuilder.BuildCheckBoxWidgetAnnotation(field.X1, field.Y1, field.X2, field.Y2, field.Name, field.IsChecked, field.CheckedValueName, offAppearanceId, checkedAppearanceId, field.Style);
                    } else if (field.Kind == FormFieldAnnotationKind.Choice) {
                        string appearanceValue = field.Values.Count > 1 ? string.Join(", ", field.Values) : field.Value;
                        string appearanceContent = PdfAcroFormDictionaryBuilder.BuildTextFieldAppearanceContent(appearanceWidth, appearanceHeight, appearanceValue, field.FontSize, field.Style);
                        byte[] appearanceBytes = PdfEncoding.Latin1GetBytes(appearanceContent);
                        string appearanceDictionary = PdfAcroFormDictionaryBuilder.BuildTextFieldAppearanceStreamDictionary(appearanceWidth, appearanceHeight, helveticaFontId, appearanceBytes.Length);
                        int appearanceId = AddStreamObject(objects, appearanceDictionary, appearanceBytes);
                        formField = PdfAnnotationDictionaryBuilder.BuildChoiceFieldWidgetAnnotation(field.X1, field.Y1, field.X2, field.Y2, field.Name, field.Options, field.Values.Count == 0 ? new[] { field.Value } : field.Values, field.FontSize, appearanceId, field.IsComboBox, field.AllowsMultipleSelection, field.Style);
                    } else {
                        string appearanceContent = PdfAcroFormDictionaryBuilder.BuildTextFieldAppearanceContent(appearanceWidth, appearanceHeight, field.Value, field.FontSize, field.Style);
                        byte[] appearanceBytes = PdfEncoding.Latin1GetBytes(appearanceContent);
                        string appearanceDictionary = PdfAcroFormDictionaryBuilder.BuildTextFieldAppearanceStreamDictionary(appearanceWidth, appearanceHeight, helveticaFontId, appearanceBytes.Length);
                        int appearanceId = AddStreamObject(objects, appearanceDictionary, appearanceBytes);
                        formField = PdfAnnotationDictionaryBuilder.BuildTextFieldWidgetAnnotation(field.X1, field.Y1, field.X2, field.Y2, field.Name, field.Value, field.FontSize, appearanceId, field.Style);
                    }

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

        int metadataId = 0;
        if (opts.IncludeXmpMetadata) {
            byte[] xmpMetadata = PdfXmpMetadataBuilder.Build(title, author, subject, keywords);
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
                    PdfEmbeddedFileDictionaryBuilder.BuildEmbeddedFileStreamDictionary(embeddedFile, fileBytes.Length),
                    fileBytes);
                int fileSpecId = AddObject(objects, PdfEmbeddedFileDictionaryBuilder.BuildFileSpecificationObject(embeddedFile, embeddedFileId));
                nameTreeEntries.Add((embeddedFile.FileName, fileSpecId));
                associatedFileIds.Add(fileSpecId);
            }

            embeddedFilesNameTreeId = AddObject(objects, PdfEmbeddedFileDictionaryBuilder.BuildEmbeddedFilesNameTree(nameTreeEntries));
        }

        int pageLabelsId = 0;
        if (opts.IncludePageLabels) {
            pageLabelsId = AddObject(objects, PdfPageLabelDictionaryBuilder.BuildGeneratedPageLabelsDictionary(
                opts.PageNumberStyle,
                opts.PageNumberStart,
                opts.PageLabelPrefix));
        }

        int viewerPreferencesId = 0;
        PdfViewerPreferencesOptions? viewerPreferences = opts.ViewerPreferencesSnapshot;
        if (viewerPreferences != null && viewerPreferences.HasAny) {
            viewerPreferencesId = AddObject(objects, PdfViewerPreferenceDictionaryBuilder.BuildGeneratedViewerPreferencesDictionary(viewerPreferences));
        }

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
            viewerPreferencesId));

        infoId = AddObject(objects, PdfInfoDictionaryBuilder.Build(title, author, subject, keywords));

        return PdfFileAssembler.Assemble(objects, catalogId, infoId);
    }

    private static string BuildPageBackground(LayoutResult.Page page, PdfOptions options, string pageBackgroundShapeContent, PdfTextWatermark? watermark, string? watermarkFontAlias, string? textWatermarkGraphicsStateName, PdfPageBorder? pageBorder, string? pageBorderGraphicsStateName) {
        var sb = new StringBuilder();
        if (options.BackgroundColor.HasValue) {
            new ContentStreamBuilder(sb)
                .SaveState()
                .FillColor(options.BackgroundColor.Value)
                .Rectangle(0, 0, options.PageWidth, options.PageHeight)
                .FillPath()
                .RestoreState();
        }

        sb.Append(pageBackgroundShapeContent);

        foreach (PageImage image in page.Images) {
            if (image.IsBackgroundDecoration) {
                AppendPageImageDraw(sb, image);
            }
        }

        if (watermark != null && watermark.Opacity > 0D && !string.IsNullOrEmpty(watermarkFontAlias)) {
            AppendTextWatermark(sb, options, watermark, watermarkFontAlias!, textWatermarkGraphicsStateName);
        }

        if (pageBorder != null && pageBorder.Opacity > 0D) {
            AppendPageBorder(sb, options, pageBorder, pageBorderGraphicsStateName);
        }

        return sb.ToString();
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
        (double x, double y, double width, double height) = FitImageToBox(image.ImageInfo, image.Fit, 0D, 0D, options.PageWidth, options.PageHeight);
        string? stateName = image.Opacity < 1D
            ? EnsureHeaderFooterGraphicsState(page, image.Opacity, image.Opacity)
            : null;

        page.Images.Add(new PageImage {
            Data = image.DataSnapshot,
            Info = image.ImageInfo,
            X = x,
            Y = y,
            W = width,
            H = height,
            IsBackgroundDecoration = true,
            Opacity = image.Opacity,
            GraphicsStateName = stateName
        });
    }

    private static (double X, double Y, double Width, double Height) FitImageToBox(OfficeIMO.Drawing.OfficeImageInfo imageInfo, OfficeIMO.Drawing.OfficeImageFit fit, double boxX, double boxY, double boxWidth, double boxHeight) {
        if (fit == OfficeIMO.Drawing.OfficeImageFit.Stretch) {
            return (boxX, boxY, boxWidth, boxHeight);
        }

        double imageWidth = imageInfo.Width > 0 ? imageInfo.Width : boxWidth;
        double imageHeight = imageInfo.Height > 0 ? imageInfo.Height : boxHeight;
        double scaleX = boxWidth / imageWidth;
        double scaleY = boxHeight / imageHeight;
        double scale = fit == OfficeIMO.Drawing.OfficeImageFit.Contain ? System.Math.Min(scaleX, scaleY) : System.Math.Max(scaleX, scaleY);
        double width = imageWidth * scale;
        double height = imageHeight * scale;
        double x = boxX + (boxWidth - width) / 2D;
        double y = boxY + (boxHeight - height) / 2D;
        return (x, y, width, height);
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

        double angle = img.RotationAngle * System.Math.PI / 180D;
        double cos = System.Math.Cos(angle);
        double sin = System.Math.Sin(angle);
        double a = img.W * cos;
        double b = img.W * sin;
        double c = -img.H * sin;
        double d = img.H * cos;
        double centerX = img.X + img.W / 2D;
        double centerY = img.Y + img.H / 2D;
        double e = centerX - (a + c) / 2D;
        double f = centerY - (b + d) / 2D;

        var content = new ContentStreamBuilder(sb)
            .SaveState();
        if (!string.IsNullOrEmpty(img.GraphicsStateName)) {
            content.GraphicsState(img.GraphicsStateName!);
        }

        content
            .TransformMatrix(a, b, c, d, e, f)
            .XObject(img.Name)
            .RestoreState();

        if (img.ClipPath != null) {
            new ContentStreamBuilder(sb)
                .RestoreState();
        }
    }

    private static void AppendTextWatermark(StringBuilder sb, PdfOptions options, PdfTextWatermark watermark, string fontAlias, string? graphicsStateName) {
        PdfStandardFont font = GetTextWatermarkFont(watermark);
        double textWidth = EstimateSimpleTextWidth(watermark.Text, font, watermark.FontSize);
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
            .TextMatrix(cos, sin, -sin, cos, originX, originY)
            .ShowHexText(EncodeWinAnsiHex(watermark.Text))
            .EndText()
            .RestoreState();
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


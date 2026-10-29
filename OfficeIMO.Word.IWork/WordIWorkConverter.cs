using System.Threading;
using System.Globalization;
using OfficeIMO.Drawing;
using OfficeIMO.IWork;
using OfficeIMO.IWork.Internal;
using OfficeIMO.Word.IWork;
using OpenXmlParagraph = DocumentFormat.OpenXml.Wordprocessing.Paragraph;
using OpenXmlRun = DocumentFormat.OpenXml.Wordprocessing.Run;
using OpenXmlParagraphProperties = DocumentFormat.OpenXml.Wordprocessing.ParagraphProperties;
using DrawingWordprocessing = DocumentFormat.OpenXml.Drawing.Wordprocessing;

namespace OfficeIMO.Word.IWork;

/// <summary>Projects Apple Pages sources into editable OfficeIMO Word documents.</summary>
public static partial class WordIWorkConverter {
    private static PagesToWordResult ProjectPages(IWorkSourceDocument source,
        IWorkConversionOptions? options = null) {
        CancellationToken cancellationToken = source.CancellationToken;
        cancellationToken.ThrowIfCancellationRequested();
        IWorkConversionOptions settings = (options ?? new IWorkConversionOptions()).Clone();
        IWorkConversionMode mode = settings.Mode;
        IWorkPreviewAsset? preview = mode == IWorkConversionMode.VisualOnly
            ? source.PreferredRasterPreview
            : null;
        if (mode == IWorkConversionMode.VisualOnly && preview == null) {
            throw new NotSupportedException("The Pages source has no embedded raster preview.");
        }

        IWorkPagesProjection projection = source.ReadPages();
        HashSet<ulong> standaloneInlineTables = FindStandaloneInlineTableBlocks(projection,
            out HashSet<ulong> inlineDrawables, cancellationToken);
        string? destinationLimitation = mode == IWorkConversionMode.VisualOnly
            ? null
            : FindWordProjectionLimitation(projection, settings.AllowPartialEditableReconstruction,
                standaloneInlineTables, inlineDrawables, cancellationToken);
        bool hasEditableContent = projection.HasEditableContent
            || settings.AllowPartialEditableReconstruction && projection.HasRecoverableContent;
        bool editable = mode != IWorkConversionMode.VisualOnly && hasEditableContent
            && destinationLimitation == null;
        IReadOnlyList<IWorkDiagnostic> destinationDiagnostics = !hasEditableContent
                || destinationLimitation == null
            ? Array.Empty<IWorkDiagnostic>()
            : new[] { new IWorkDiagnostic(IWorkDiagnosticSeverity.Warning,
                "IWORK_PAGES_WORD_DESTINATION_UNSUPPORTED", destinationLimitation) };
        if (editable && settings.AllowPartialEditableReconstruction && RequiresWordRounding(projection, inlineDrawables, cancellationToken)) {
            destinationDiagnostics = destinationDiagnostics.Concat(new[] {
                new IWorkDiagnostic(IWorkDiagnosticSeverity.Warning, "IWORK_PAGES_DOCX_PRECISION",
                    "Source geometry, spacing, or font sizes were rounded to the nearest supported DOCX measurement unit; original values remain available on the projection.")
            }).ToArray();
        }
        if (editable && settings.AllowPartialEditableReconstruction && projection.Tables.Any(table =>
                HasTableGeometry(table) && !CanApplyInlineTableGeometry(table, projection.PageLayout,
                    standaloneInlineTables, cancellationToken))) {
            destinationDiagnostics = destinationDiagnostics.Concat(new[] {
                new IWorkDiagnostic(IWorkDiagnosticSeverity.Warning, "IWORK_PAGES_TABLE_LAYOUT_APPROXIMATED",
                    "Positioned source tables were retained as editable flowing DOCX tables; source geometry remains available on the projection.",
                    lossKind: global::OfficeIMO.OfficeConversionLossKind.Approximation)
            }).ToArray();
        }
        int reconstructedSectionCount = 1 + projection.Body.Paragraphs.Count(paragraph =>
            paragraph.BreakKind == IWorkParagraphBreakKind.Section);
        if (editable) {
            if (projection.Tables.Any(table => table.ResolveMergedBorders(cancellationToken).Any(merge => !merge.IsUniform))) {
                destinationDiagnostics = destinationDiagnostics.Concat(new[] {
                    new IWorkDiagnostic(IWorkDiagnosticSeverity.Warning, "IWORK_PAGES_MERGED_TABLE_BORDERS_APPROXIMATED",
                        "Merged DOCX cells use the first recovered segment of each outer edge; varying native perimeter segments remain on the source projection.",
                        lossKind: global::OfficeIMO.OfficeConversionLossKind.Approximation)
                }).ToArray();
            }
            if (AllPagesText(projection).SelectMany(content => content.Paragraphs).Any(paragraph => paragraph.ListLayout != null
                    && !CanApplyQualifiedListLayout(paragraph))) {
                destinationDiagnostics = destinationDiagnostics.Concat(new[] {
                    new IWorkDiagnostic(IWorkDiagnosticSeverity.Warning, "IWORK_PAGES_LIST_LAYOUT_APPROXIMATED",
                        "DOCX uses destination list-marker placement and size; native list geometry remains available on the projection.",
                        lossKind: global::OfficeIMO.OfficeConversionLossKind.Approximation)
                }).ToArray();
            }
            destinationDiagnostics = destinationDiagnostics.Concat(IWorkNumericDisplayDiagnostics.ForTextTables(
                projection.Tables.SelectMany(table => table.Cells), "PAGES", "DOCX", cancellationToken)).ToArray();
            if (projection.Tables.Any(table => table.HiddenRows.Count > 0 || table.HiddenColumns.Count > 0)) {
                destinationDiagnostics = destinationDiagnostics.Concat(new[] {
                    new IWorkDiagnostic(IWorkDiagnosticSeverity.Warning, "IWORK_PAGES_TABLE_VISIBILITY_OMITTED",
                        "DOCX tables retain source-hidden rows and columns as visible content; their visibility is retained on the iWork projection.",
                        lossKind: global::OfficeIMO.OfficeConversionLossKind.Omission)
                }).ToArray();
            }
        }
        if (editable && projection.Sections.Count > reconstructedSectionCount) {
            destinationDiagnostics = destinationDiagnostics.Concat(new[] {
                new IWorkDiagnostic(IWorkDiagnosticSeverity.Warning, "IWORK_PAGES_SECTION_CONTENT_OMITTED",
                    (projection.Sections.Count - reconstructedSectionCount) + " source section(s) have no recovered DOCX section boundary; their headers and footers were not reconstructed.",
                    lossKind: global::OfficeIMO.OfficeConversionLossKind.Omission)
            }).ToArray();
        }
        if (editable && settings.FindUnacceptedDestinationOmission(destinationDiagnostics) is { } omission) {
            destinationLimitation = omission;
            editable = false;
        }
        if (editable && settings.AllowPartialEditableReconstruction &&
            (!projection.HasEditableContent || destinationDiagnostics.Count > 0 || projection.Diagnostics.Any(diagnostic =>
                diagnostic.Severity != IWorkDiagnosticSeverity.Information))) {
            destinationDiagnostics = destinationDiagnostics.Concat(new[] {
                new IWorkDiagnostic(IWorkDiagnosticSeverity.Warning,
                    "IWORK_PARTIAL_EDITABLE_RECONSTRUCTION",
                    "Recovered editable content was retained under the explicit partial-reconstruction policy; source and destination diagnostics describe incomplete details.")
            }).ToArray();
        }
        if (!editable && mode == IWorkConversionMode.EditableOnly) {
            throw new InvalidDataException(destinationLimitation
                ?? "The Pages source has no supported editable content.");
        }

        preview ??= editable ? null : source.PreferredRasterPreview;
        if (!editable && preview == null) {
            throw new NotSupportedException("The Pages source has no supported editable content or embedded raster preview.");
        }

        if (!editable) settings.ValidateVisualPreview(preview);

        WordDocument document = WordDocument.Create();
        try {
            if (projection.PageLayout is { } pageLayout && CanApplyPageLayout(pageLayout, settings.AllowPartialEditableReconstruction)) {
                ApplyPageLayout(document.Sections[0], pageLayout);
            }
            if (editable) {
                if (PagesDecimalTabSymbols(projection, cancellationToken).SingleOrDefault() is string decimalSymbol)
                    document.Settings.DecimalSymbol = decimalSymbol;
                var nativeLists = new IWorkNativeListCatalog(document);
                var cellComments = new List<WordCellComment>();
                (double contentWidth, double contentHeight) = ContentBox(document.Sections[0]);
                var semanticSections = new List<WordSection> { document.Sections[0] };
                var pageHosts = new Dictionary<int, WordParagraph>();
                var pageTableAnchors = new Dictionary<int, WordTable>();
                int currentPageIndex = 1;
                var drawableLookup = projection.Drawables.ToDictionary(DrawableIdentifier);
                AddRichText(projection.Body, value => {
                        WordParagraph paragraph = document.AddParagraph(value);
                        if (!pageHosts.ContainsKey(currentPageIndex)) {
                            pageHosts.Add(currentPageIndex, paragraph);
                        }
                        return paragraph;
                    }, nativeLists, () => {
                        WordParagraph pageBreak = document.AddPageBreak();
                        currentPageIndex++;
                        return pageBreak;
                    },
                    breakKind => {
                        WordSection section = document.AddSection(breakKind == IWorkParagraphBreakKind.Layout
                            ? WordSectionBreakType.Continuous
                            : WordSectionBreakType.NextPage);
                        if (breakKind == IWorkParagraphBreakKind.Section) {
                            semanticSections.Add(section);
                            currentPageIndex++;
                        }
                    }, cancellationToken: cancellationToken, addInlineObject: (paragraph, run) => {
                        ulong identifier = run.InlineObject!.Drawable.RecordIdentifier;
                        if (AddInlineObject(document, paragraph, drawableLookup[identifier],
                                nativeLists, contentWidth, contentHeight, cellComments, standaloneInlineTables,
                                projection.PageLayout, cancellationToken)
                            && pageHosts.TryGetValue(currentPageIndex, out WordParagraph? host) && host == paragraph) {
                            pageHosts.Remove(currentPageIndex);
                        }
                    });
                if (projection.PageLayout != null) {
                    foreach (WordSection section in document.Sections) ApplyPageLayout(section, projection.PageLayout);
                }
                for (int drawableIndex = 0; drawableIndex < projection.Drawables.Count; drawableIndex++) {
                    cancellationToken.ThrowIfCancellationRequested();
                    IWorkPagesDrawable sourceDrawable = projection.Drawables[drawableIndex];
                    if (inlineDrawables.Contains(DrawableIdentifier(sourceDrawable))) continue;
                    WordParagraph? pageHost = sourceDrawable.PageIndex.HasValue
                        && pageHosts.TryGetValue(sourceDrawable.PageIndex.Value, out WordParagraph? host)
                            ? host
                            : null;
                    uint zOrder = checked(251658240U + (uint)drawableIndex);
                    switch (sourceDrawable.Kind) {
                        case IWorkPagesDrawableKind.TextBox:
                            AddRichTextBox(document, sourceDrawable.TextBox!, nativeLists, pageHost, cancellationToken).ZOrder = zOrder;
                            break;
                        case IWorkPagesDrawableKind.Image:
                            AddImage(document, sourceDrawable.Image!, contentWidth, contentHeight, pageHost).ZOrder = zOrder;
                            break;
                        case IWorkPagesDrawableKind.Table:
                            WordTable? priorTable = sourceDrawable.PageIndex.HasValue
                                && pageTableAnchors.TryGetValue(sourceDrawable.PageIndex.Value,
                                    out WordTable? anchor)
                                    ? anchor
                                    : null;
                            WordTable? insertedTable = AddTable(document,
                                sourceDrawable.Table!, nativeLists, pageHost, priorTable, cellComments, cancellationToken);
                            if (sourceDrawable.PageIndex.HasValue && insertedTable != null) {
                                pageTableAnchors[sourceDrawable.PageIndex.Value] = insertedTable;
                            }
                            break;
                    }
                }
                document.AddCellComments(cellComments, cancellationToken);
                bool hasAnyEvenPageTemplate = projection.Sections.Any(section => section.DifferentOddAndEvenPages);
                for (int sectionIndex = 0; sectionIndex < projection.Sections.Count; sectionIndex++) {
                    cancellationToken.ThrowIfCancellationRequested();
                    IWorkPagesSection sourceSection = projection.Sections[sectionIndex];
                    if (sectionIndex >= semanticSections.Count) break;
                    WordSection targetSection = semanticSections[sectionIndex];
                    AddSectionHeadersAndFooters(targetSection, sourceSection, nativeLists,
                        hasAnyEvenPageTemplate, cancellationToken);
                }
            } else {
                byte[] bytes = preview!.GetBytes();
                using var image = new MemoryStream(bytes, writable: false);
                WordSection section = document.Sections[0];
                (double contentWidth, double contentHeight) = ContentBox(section);
                (double width, double height) = PreviewSize(preview, contentWidth, contentHeight);
                document.AddParagraph().AddImage(image, PreviewFileName(preview), OfficeIMO.OfficeLength.Pixels(Helpers.ConvertPointsToPixels(width)), OfficeIMO.OfficeLength.Pixels(Helpers.ConvertPointsToPixels(height)), options: new OfficeIMO.Word.WordImageOptions { Description = "Visual fallback from the source Pages package" });
            }

            IWorkProjectionKind kind = editable
                ? IWorkProjectionKind.EditableReconstruction
                : IWorkProjectionKind.VisualFallback;
            cancellationToken.ThrowIfCancellationRequested();
            return new PagesToWordResult(document, source, projection,
                projection.CreateConversionReport(kind, preview, destinationDiagnostics,
                    settings.AllowPartialEditableReconstruction, reconstructedSectionCount));
        } catch {
            document.Dispose();
            throw;
        }
    }

    private static (double Width, double Height) PreviewSize(IWorkPreviewAsset preview,
        double maximumWidth, double maximumHeight) {
        double width = preview.PixelWidth.GetValueOrDefault(800) * 72d / 96d;
        double height = preview.PixelHeight.GetValueOrDefault(1040) * 72d / 96d;
        double scale = Math.Min(1d, Math.Min(maximumWidth / width, maximumHeight / height));
        return (Math.Max(1, width * scale), Math.Max(1, height * scale));
    }

    private static (double Width, double Height) FitInside(double width, double height,
        double maximumWidth, double maximumHeight) {
        double safeWidth = Math.Max(1d, width);
        double safeHeight = Math.Max(1d, height);
        double scale = Math.Min(1d, Math.Min(maximumWidth / safeWidth, maximumHeight / safeHeight));
        return (safeWidth * scale, safeHeight * scale);
    }

    private static (double Width, double Height) ContentBox(WordSection section) {
        uint pageWidth = checked((uint?)(section.PageSettings.Width?.ToTwips())) ?? checked((uint)(WordPageSizes.Letter.Width.ToTwips()));
        uint pageHeight = checked((uint?)(section.PageSettings.Height?.ToTwips())) ?? checked((uint)(WordPageSizes.Letter.Height.ToTwips()));
        long horizontalMargins = (long)checked((uint)(section.Margins.Left.GetValueOrDefault().ToTwips())) + checked((uint)(section.Margins.Right.GetValueOrDefault().ToTwips()));
        long verticalMargins = (long)checked((int?)(section.Margins.Top?.ToTwips())).GetValueOrDefault()
            + checked((int?)(section.Margins.Bottom?.ToTwips())).GetValueOrDefault();
        return (Math.Max(1L, (long)pageWidth - Math.Max(0L, horizontalMargins)) / 20d,
            Math.Max(1L, (long)pageHeight - Math.Max(0L, verticalMargins)) / 20d);
    }

    private static string PreviewFileName(IWorkPreviewAsset preview) =>
        preview.MediaType == "image/png" ? "pages-preview.png" : "pages-preview.jpg";

    private static WordImage AddImage(WordDocument document, IWorkImageAsset source,
        double contentWidth, double contentHeight, WordParagraph? pageHost = null) {
        using var image = new MemoryStream(source.GetBytes(), writable: false);
        double width = source.Geometry?.WidthPoints
            ?? source.PixelWidth.GetValueOrDefault(640) * 72d / 96d;
        double height = source.Geometry?.HeightPoints
            ?? source.PixelHeight.GetValueOrDefault(480) * 72d / 96d;
        if (source.Geometry == null) {
            (width, height) = FitInside(width, height, contentWidth, contentHeight);
        }
        WordImage target = (pageHost ?? document.AddParagraph()).AddImage(image, source.FileName, OfficeIMO.OfficeLength.Pixels(Helpers.ConvertPointsToPixels(width)), OfficeIMO.OfficeLength.Pixels(Helpers.ConvertPointsToPixels(height)), options: new OfficeIMO.Word.WordImageOptions { Wrapping = WordImageTextWrapping.Square, Description = source.AccessibilityDescription ?? "Image imported from Pages" });
        if (source.Geometry is { } geometry) {
            target.horizontalPosition.RelativeFrom =
                DrawingWordprocessing.HorizontalRelativePositionValues.Page;
            target.horizontalPosition.PositionOffset = new DrawingWordprocessing.PositionOffset {
                Text = ToEmusInt32(geometry.LeftPoints).ToString(CultureInfo.InvariantCulture)
            };
            target.verticalPosition.RelativeFrom =
                DrawingWordprocessing.VerticalRelativePositionValues.Page;
            target.verticalPosition.PositionOffset = new DrawingWordprocessing.PositionOffset {
                Text = ToEmusInt32(geometry.TopPoints).ToString(CultureInfo.InvariantCulture)
            };
        }
        return target;
    }

    private static void AddSectionHeadersAndFooters(WordSection target,
        IWorkPagesSection source, IWorkNativeListCatalog nativeLists,
        bool hasAnyEvenPageTemplate, CancellationToken cancellationToken) {
        if (source.HasDefaultPageTemplate) {
            WordHeader header = target.GetOrCreateHeader(WordHeaderFooterType.Default);
            WordFooter footer = target.GetOrCreateFooter(WordHeaderFooterType.Default);
            foreach (IWorkTextContent content in source.DefaultPageHeaderContents) {
                AddRichText(content, header.AddParagraph, nativeLists, cancellationToken: cancellationToken);
            }
            foreach (IWorkTextContent content in source.DefaultPageFooterContents) {
                AddRichText(content, footer.AddParagraph, nativeLists, cancellationToken: cancellationToken);
            }
        }
        if (source.DifferentFirstPage || source.HideFirstPageHeadersAndFooters) {
            WordHeader header = target.GetOrCreateHeader(WordHeaderFooterType.First);
            WordFooter footer = target.GetOrCreateFooter(WordHeaderFooterType.First);
            foreach (IWorkTextContent content in source.HideFirstPageHeadersAndFooters
                         ? Array.Empty<IWorkTextContent>() : source.FirstPageHeaderContents) {
                AddRichText(content, header.AddParagraph, nativeLists, cancellationToken: cancellationToken);
            }
            foreach (IWorkTextContent content in source.HideFirstPageHeadersAndFooters
                         ? Array.Empty<IWorkTextContent>() : source.FirstPageFooterContents) {
                AddRichText(content, footer.AddParagraph, nativeLists, cancellationToken: cancellationToken);
            }
        }
        if (hasAnyEvenPageTemplate) {
            WordHeader header = target.GetOrCreateHeader(WordHeaderFooterType.Even);
            WordFooter footer = target.GetOrCreateFooter(WordHeaderFooterType.Even);
            IReadOnlyList<IWorkTextContent> headerContents = source.DifferentOddAndEvenPages
                ? source.EvenPageHeaderContents
                : source.DefaultPageHeaderContents;
            IReadOnlyList<IWorkTextContent> footerContents = source.DifferentOddAndEvenPages
                ? source.EvenPageFooterContents
                : source.DefaultPageFooterContents;
            foreach (IWorkTextContent content in headerContents) {
                AddRichText(content, header.AddParagraph, nativeLists, cancellationToken: cancellationToken);
            }
            foreach (IWorkTextContent content in footerContents) {
                AddRichText(content, footer.AddParagraph, nativeLists, cancellationToken: cancellationToken);
            }
        }
    }

    private static WordTextBox AddRichTextBox(WordDocument document, IWorkTextBox source,
        IWorkNativeListCatalog nativeLists, WordParagraph? pageHost, CancellationToken cancellationToken) {
        WordTextBox textBox = pageHost == null
            ? document.AddTextBox(string.Empty)
            : pageHost.AddTextBox(string.Empty, WordImageTextWrapping.Square);
        if (source.Geometry is { } geometry) {
            textBox.HorizontalPositionRelativeFrom = WordHorizontalRelativePosition.Page;
            textBox.VerticalPositionRelativeFrom = WordVerticalRelativePosition.Page;
            textBox.HorizontalPositionOffset = OfficeLength.Emus(ToEmusInt32(geometry.LeftPoints));
            textBox.VerticalPositionOffset = OfficeLength.Emus(ToEmusInt32(geometry.TopPoints));
            textBox.Width = OfficeLength.Emus(ToEmusInt64(geometry.WidthPoints));
            textBox.Height = OfficeLength.Emus(ToEmusInt64(geometry.HeightPoints));
        }
        textBox.Description = source.AccessibilityDescription;

        DocumentFormat.OpenXml.Wordprocessing.TextBoxContent? content = textBox.Content;
        if (content == null) return textBox;
        content.RemoveAllChildren<OpenXmlParagraph>();
        AddRichText(source.Content, value => {
            var paragraph = new OpenXmlParagraph();
            content.Append(paragraph);
            var result = new WordParagraph(document, paragraph, newRun: false);
            if (value.Length > 0) result.AddText(value);
            return result;
        }, nativeLists, cancellationToken: cancellationToken);
        if (!content.Elements<OpenXmlParagraph>().Any()) {
            content.Append(new OpenXmlParagraph(new OpenXmlRun()));
        }
        return textBox;
    }

    private static void AddRichText(IWorkTextContent content, Func<string, WordParagraph> addParagraph,
        IWorkNativeListCatalog nativeLists,
        Func<WordParagraph>? addPageBreak = null,
        Action<IWorkParagraphBreakKind>? addSectionBreak = null,
        bool forceBold = false, CancellationToken cancellationToken = default,
        Action<WordParagraph, IWorkTextRun>? addInlineObject = null, IWorkParagraphStyle? defaultStyle = null) {
        // Inline tables can recursively project another text storage. Its lists
        // must not replace the current list in the surrounding storage.
        nativeLists = nativeLists.CreateTextScope();
        ulong? previousListIdentifier = null;
        bool hasPreviousListParagraph = false;
        foreach (IWorkTextParagraph sourceParagraph in content.Paragraphs) {
            cancellationToken.ThrowIfCancellationRequested();
            WordParagraph paragraph = addParagraph(string.Empty);
            InitializeParagraphSpacing(paragraph);
            if (defaultStyle != null) {
                ApplyParagraphStyle(paragraph, defaultStyle, sourceParagraph.Text);
                ApplyParagraphTextStyle(paragraph, defaultStyle.TextStyle);
            }
            ApplyParagraphStyle(paragraph, sourceParagraph);
            if (forceBold) paragraph.Bold = true;
            ApplyParagraphTextStyle(paragraph, sourceParagraph.Style.TextStyle);
            if (sourceParagraph.ListLevel >= 0) {
                bool startsNewList = !hasPreviousListParagraph
                    || sourceParagraph.ListIdentifier != previousListIdentifier
                    || sourceParagraph.ListStartNumber.HasValue;
                bool qualifiedLayout = CanApplyQualifiedListLayout(sourceParagraph, out double markerSize);
                if (qualifiedLayout) {
                    double firstRunSize = sourceParagraph.Runs.FirstOrDefault()?.Style.FontSizePoints
                        ?? sourceParagraph.Style.TextStyle.FontSizePoints!.Value;
                    if (firstRunSize != sourceParagraph.Style.TextStyle.FontSizePoints) {
                        // Pages interprets relative marker sizes using the leading run and paragraph mark.
                        // Keep both at the first body run's size; the numbering level owns the marker size.
                        paragraph.FontSize = global::OfficeIMO.OfficeLength.Points(firstRunSize);
                        paragraph.CopyRunFormattingToParagraphMark();
                    }
                }
                nativeLists.Apply(paragraph, sourceParagraph.ListLevel,
                    sourceParagraph.ListLabel, qualifiedLayout
                        ? IWorkDecimalListLayout.MarkerFontName(sourceParagraph) : sourceParagraph.ListFontName,
                    sourceParagraph.ListMarkerKind, startsNewList, sourceParagraph.ListNumber,
                    qualifiedLayout ? checked((int?)(paragraph.IndentationBefore?.ToTwips())) : null,
                    qualifiedLayout ? checked((int?)(paragraph.IndentationHanging?.ToTwips())) : null,
                    qualifiedLayout ? checked((int)markerSize) : null);
                previousListIdentifier = sourceParagraph.ListIdentifier;
                hasPreviousListParagraph = true;
            } else {
                previousListIdentifier = null;
                hasPreviousListParagraph = false;
            }
            foreach (IWorkTextRun sourceRun in sourceParagraph.Runs) {
                cancellationToken.ThrowIfCancellationRequested();
                if (sourceRun.InlineObject != null) addInlineObject?.Invoke(paragraph, sourceRun);
                else AddStyledTextRun(paragraph, sourceRun, forceBold, defaultStyle?.TextStyle);
            }
            if (sourceParagraph.BreakKind == IWorkParagraphBreakKind.Page) addPageBreak?.Invoke();
            else if (sourceParagraph.BreakKind is IWorkParagraphBreakKind.Section
                     or IWorkParagraphBreakKind.Layout) addSectionBreak?.Invoke(sourceParagraph.BreakKind);
        }
    }

    private static void AddStyledTextRun(WordParagraph paragraph, IWorkTextRun sourceRun,
        bool forceBold = false, IWorkTextStyle? defaultStyle = null) {
        string[] lines = sourceRun.Text.Split('\n');
        for (int index = 0; index < lines.Length; index++) {
            if (index > 0) paragraph.AddBreak();
            if (lines[index].Length == 0) continue;

            WordParagraph run;
            if (sourceRun.Hyperlink != null
                && Uri.TryCreate(sourceRun.Hyperlink, UriKind.Absolute, out Uri? uri)) {
                paragraph.AddHyperLink(lines[index], uri);
                run = new WordParagraph(paragraph._document, paragraph._paragraph, paragraph._hyperlink!);
            } else {
                run = paragraph.AddText(lines[index]);
            }
            if (forceBold) run.Bold = true;
            if (defaultStyle != null) ApplyTextStyle(run, defaultStyle);
            ApplyTextStyle(run, sourceRun.Style);
        }
    }

    private static void ApplyParagraphStyle(WordParagraph paragraph, IWorkTextParagraph source) {
        ApplyParagraphStyle(paragraph, source.Style, source.Text);
        if (CanApplyQualifiedListLayout(source)) {
            // Native relative marker offsets are quantized to DOCX twentieth-points,
            // matching the independently exported decimal-list reference.
            double offset = Math.Round(source.ListLayout!.TextIndentEm
                * source.Style.TextStyle.FontSizePoints!.Value * 20d,
                MidpointRounding.AwayFromZero) / 20d;
            paragraph.IndentationBefore = global::OfficeIMO.OfficeLength.Points(Math.Round((source.ListLayout.MarkerIndentPoints
                + source.ListLayout.TextIndentEm * source.Style.TextStyle.FontSizePoints!.Value) * 20d,
                MidpointRounding.AwayFromZero) / 20d);
            paragraph.IndentationFirstLine= null;
            paragraph.IndentationHanging = global::OfficeIMO.OfficeLength.Points(offset);
        }
    }

    private static void InitializeParagraphSpacing(WordParagraph paragraph) {
        // Start with native defaults before layering table and rich-text styles.
        // Blank header/footer stories must not inherit destination template spacing.
        paragraph.LineSpacingBefore = global::OfficeIMO.OfficeLength.Zero;
        paragraph.LineSpacingAfter = global::OfficeIMO.OfficeLength.Zero;
    }

    private static void ApplyParagraphStyle(WordParagraph paragraph, IWorkParagraphStyle style, string text) {
        paragraph.BiDi = OfficeTextElements.ResolveBaseDirection(text)
            == OfficeTextDirection.RightToLeft;
        if (style.Alignment.HasValue) {
            paragraph.ParagraphAlignment = style.Alignment.Value switch {
                IWorkTextAlignment.Natural => WordParagraphAlignment.Start,
                IWorkTextAlignment.Left => WordParagraphAlignment.Left,
                IWorkTextAlignment.Center => WordParagraphAlignment.Center,
                IWorkTextAlignment.Right => WordParagraphAlignment.Right,
                IWorkTextAlignment.Justified => WordParagraphAlignment.Both,
                _ => throw new InvalidOperationException("Unsupported iWork paragraph alignment.")
            };
        }
        if (style.FirstLineIndentPoints.HasValue) if ((style.FirstLineIndentPoints) is { } wordLength30336) { paragraph.IndentationFirstLine = global::OfficeIMO.OfficeLength.Points(wordLength30336); }
        if (style.LeftIndentPoints.HasValue) if ((style.LeftIndentPoints) is { } wordLength30446) { paragraph.IndentationBefore = global::OfficeIMO.OfficeLength.Points(wordLength30446); }
        if (style.RightIndentPoints.HasValue) if ((style.RightIndentPoints) is { } wordLength30551) { paragraph.IndentationAfter = global::OfficeIMO.OfficeLength.Points(wordLength30551); }
        if (style.SpaceBeforePoints.HasValue) if ((style.SpaceBeforePoints) is { } wordLength30658) { paragraph.LineSpacingBefore = global::OfficeIMO.OfficeLength.Points(wordLength30658); }
        if (style.SpaceAfterPoints.HasValue) if ((style.SpaceAfterPoints) is { } wordLength30763) { paragraph.LineSpacingAfter = global::OfficeIMO.OfficeLength.Points(wordLength30763); }
        if (style.LineSpacingMultiplier is double multiplier) {
            paragraph.LineSpacing = WordLineSpacing.Multiple(multiplier);
        }
        if (style.TabStops != null) {
            paragraph.ClearTabStops();
            foreach (var tab in style.TabStops) paragraph.AddTabStop(
                checked((int)Math.Round(tab.PositionPoints * 20d, MidpointRounding.AwayFromZero)), tab.Alignment switch {
                    IWorkTabAlignment.Center => WordTabAlignment.Center,
                    IWorkTabAlignment.Right => WordTabAlignment.Right,
                    IWorkTabAlignment.Decimal => WordTabAlignment.Decimal,
                    _ => WordTabAlignment.Left
                });
        }
        if (style.PageBreakBefore.HasValue) paragraph.PageBreakBefore = style.PageBreakBefore.Value;
        if (style.KeepWithNext.HasValue) paragraph.KeepWithNext = style.KeepWithNext.Value;
        if (style.KeepLinesTogether.HasValue) paragraph.KeepLinesTogether = style.KeepLinesTogether.Value;
    }

    private static void ApplyParagraphTextStyle(WordParagraph paragraph, IWorkTextStyle style) {
        ApplyTextStyle(paragraph, style);
        paragraph.CopyRunFormattingToParagraphMark();
    }

    private static void ApplyTextStyle(WordParagraph run, IWorkTextStyle style) {
        if (style.Bold.HasValue) run.Bold = style.Bold.Value;
        if (style.Italic.HasValue) run.Italic = style.Italic.Value;
        if (style.Underline.HasValue) run.Underline = style.Underline.Value ? WordUnderlineStyle.Single : null;
        if (style.Strikethrough.HasValue) run.Strike = style.Strikethrough.Value;
        if (style.FontSizePoints.HasValue) run.FontSize = global::OfficeIMO.OfficeLength.Points(style.FontSizePoints.Value);
        if (style.CharacterSpacingPoints is double spacing) run.Spacing = ToCharacterSpacingTwips(spacing);
        if (!string.IsNullOrWhiteSpace(style.FontName)) run.FontFamily = style.FontName;
        if (style.Color != null) run.ColorHex = style.Color.RgbHex;
        if (style.BackgroundColor != null) run.RunShadingFillColorHex = style.BackgroundColor.RgbHex;
    }

    private static void ApplyPageLayout(WordSection section, IWorkPageLayout layout) {
        section.PageSettings.Orientation = layout.IsLandscape
            ? OfficePageOrientation.Landscape
            : OfficePageOrientation.Portrait;
        section.PageSettings.Width = global::OfficeIMO.OfficeLength.Points((ToTwips(layout.WidthPoints)) / 20D);
        section.PageSettings.Height = global::OfficeIMO.OfficeLength.Points((ToTwips(layout.HeightPoints)) / 20D);
        section.Margins.Left = global::OfficeIMO.OfficeLength.Points((ToTwips(layout.LeftMarginPoints)) / 20D);
        section.Margins.Right = global::OfficeIMO.OfficeLength.Points((ToTwips(layout.RightMarginPoints)) / 20D);
        section.Margins.Top = global::OfficeIMO.OfficeLength.Points((checked((int)ToTwips(layout.TopMarginPoints))) / 20D);
        section.Margins.Bottom = global::OfficeIMO.OfficeLength.Points((checked((int)ToTwips(layout.BottomMarginPoints))) / 20D);
        section.Margins.HeaderDistance = global::OfficeIMO.OfficeLength.Points((ToTwips(layout.HeaderMarginPoints)) / 20D);
        section.Margins.FooterDistance = global::OfficeIMO.OfficeLength.Points((ToTwips(layout.FooterMarginPoints)) / 20D);
    }

    private static uint ToTwips(double points) {
        double value = Math.Round(points * 20d, MidpointRounding.AwayFromZero);
        if (value < 0 || value > uint.MaxValue) throw new InvalidDataException("A Pages page measurement exceeds the DOCX range.");
        return (uint)value;
    }

    private static int ToCharacterSpacingTwips(double points) {
        double value = Math.Round(points * 20d, MidpointRounding.AwayFromZero);
        if (double.IsNaN(value) || double.IsInfinity(value) || value < int.MinValue || value > int.MaxValue)
            throw new InvalidDataException("Pages character spacing exceeds the DOCX range.");
        return (int)value;
    }

    private static int ToSignedTwips(double points) {
        double value = Math.Round(points * 20d, MidpointRounding.AwayFromZero);
        if (value < 0 || value > int.MaxValue) throw new InvalidDataException("A Pages table measurement exceeds the DOCX range.");
        return (int)value;
    }
}

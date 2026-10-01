using System.Threading;
using System.Globalization;
using OfficeIMO.Drawing;
using OfficeIMO.IWork;
using OfficeIMO.Word.IWork;
using OpenXmlParagraph = DocumentFormat.OpenXml.Wordprocessing.Paragraph;
using OpenXmlRun = DocumentFormat.OpenXml.Wordprocessing.Run;
using OpenXmlNumberingId = DocumentFormat.OpenXml.Wordprocessing.NumberingId;
using OpenXmlNumberingLevelReference = DocumentFormat.OpenXml.Wordprocessing.NumberingLevelReference;
using OpenXmlNumberingProperties = DocumentFormat.OpenXml.Wordprocessing.NumberingProperties;
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
        string? destinationLimitation = mode == IWorkConversionMode.VisualOnly
            ? null
            : FindWordProjectionLimitation(projection, settings.AllowPartialEditableReconstruction);
        bool hasEditableContent = projection.HasEditableContent
            || settings.AllowPartialEditableReconstruction && projection.HasRecoverableContent;
        bool editable = mode != IWorkConversionMode.VisualOnly && hasEditableContent
            && destinationLimitation == null;
        IReadOnlyList<IWorkDiagnostic> destinationDiagnostics = !hasEditableContent
                || destinationLimitation == null
            ? Array.Empty<IWorkDiagnostic>()
            : new[] { new IWorkDiagnostic(IWorkDiagnosticSeverity.Warning,
                "IWORK_PAGES_WORD_DESTINATION_UNSUPPORTED", destinationLimitation) };
        if (editable && settings.AllowPartialEditableReconstruction && RequiresWordRounding(projection)) {
            destinationDiagnostics = destinationDiagnostics.Concat(new[] {
                new IWorkDiagnostic(IWorkDiagnosticSeverity.Warning, "IWORK_PAGES_DOCX_PRECISION",
                    "Source geometry, spacing, or font sizes were rounded to the nearest supported DOCX measurement unit; original values remain available on the projection.")
            }).ToArray();
        }
        if (editable && settings.AllowPartialEditableReconstruction && projection.Tables.Any(table =>
                table.Geometry is { } geometry && (Math.Abs(geometry.LeftPoints) > 0.000001d
                    || Math.Abs(geometry.TopPoints) > 0.000001d || Math.Abs(geometry.WidthPoints) > 0.000001d
                    || Math.Abs(geometry.HeightPoints) > 0.000001d || Math.Abs(geometry.RotationDegrees) > 0.000001d))) {
            destinationDiagnostics = destinationDiagnostics.Concat(new[] {
                new IWorkDiagnostic(IWorkDiagnosticSeverity.Warning, "IWORK_PAGES_TABLE_LAYOUT_APPROXIMATED",
                    "Positioned source tables were retained as editable flowing DOCX tables; source geometry remains available on the projection.",
                    lossKind: global::OfficeIMO.OfficeConversionLossKind.Approximation)
            }).ToArray();
        }
        int reconstructedSectionCount = 1 + projection.Body.Paragraphs.Count(paragraph =>
            paragraph.BreakKind == IWorkParagraphBreakKind.Section);
        if (editable && projection.Tables.SelectMany(table => table.Cells).Any(cell => cell.NumberFormat != null)) {
            destinationDiagnostics = destinationDiagnostics.Concat(new[] {
                new IWorkDiagnostic(IWorkDiagnosticSeverity.Warning, "IWORK_PAGES_NUMBER_FORMAT_OMITTED",
                    "DOCX table cells retain raw cached values without the source numeric display formats. Semantic number formats remain available on the source projection.",
                    lossKind: global::OfficeIMO.OfficeConversionLossKind.Omission)
            }).ToArray();
        }
        if (editable && projection.Sections.Count > reconstructedSectionCount) {
            destinationDiagnostics = destinationDiagnostics.Concat(new[] {
                new IWorkDiagnostic(IWorkDiagnosticSeverity.Warning, "IWORK_PAGES_SECTION_CONTENT_OMITTED",
                    (projection.Sections.Count - reconstructedSectionCount) + " source section(s) have no recovered DOCX section boundary; their headers and footers were not reconstructed.",
                    lossKind: global::OfficeIMO.OfficeConversionLossKind.Omission)
            }).ToArray();
        }
        if (editable && settings.AllowPartialEditableReconstruction &&
            (!projection.HasEditableContent || destinationDiagnostics.Count > 0 || projection.Diagnostics.Any(diagnostic =>
                diagnostic.Severity != IWorkDiagnosticSeverity.Information))) {
            destinationDiagnostics = destinationDiagnostics.Concat(new[] {
                new IWorkDiagnostic(IWorkDiagnosticSeverity.Warning,
                    "IWORK_PARTIAL_EDITABLE_RECONSTRUCTION",
                    "Recovered editable content was retained under the explicit partial-reconstruction policy; source diagnostics describe incomplete details.")
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
                var nativeLists = new IWorkNativeListCatalog(document);
                (double contentWidth, double contentHeight) = ContentBox(document.Sections[0]);
                var semanticSections = new List<WordSection> { document.Sections[0] };
                var pageHosts = new Dictionary<int, WordParagraph>();
                var pageTableAnchors = new Dictionary<int, WordTable>();
                int currentPageIndex = 1;
                var inlineDrawables = new HashSet<ulong>(projection.Body.Paragraphs.SelectMany(paragraph => paragraph.Runs)
                    .Where(run => run.InlineObject != null).Select(run => run.InlineObject!.Drawable.RecordIdentifier));
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
                    }, cancellationToken: cancellationToken, addInlineObject: (paragraph, run) =>
                        AddInlineObject(document, paragraph, drawableLookup[run.InlineObject!.Drawable.RecordIdentifier],
                            nativeLists, contentWidth, contentHeight, cancellationToken));
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
                                sourceDrawable.Table!, nativeLists, pageHost, priorTable, cancellationToken);
                            if (sourceDrawable.PageIndex.HasValue && insertedTable != null) {
                                pageTableAnchors[sourceDrawable.PageIndex.Value] = insertedTable;
                            }
                            break;
                    }
                }
                bool hasAnyEvenPageTemplate = projection.Sections.Any(section => section.HasEvenPageTemplate);
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
                document.AddParagraph().AddImage(image, PreviewFileName(preview), width, height,
                    description: "Visual fallback from the source Pages package");
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
        uint pageWidth = section.PageSettings.Width ?? WordPageSizes.Letter.WidthTwips;
        uint pageHeight = section.PageSettings.Height ?? WordPageSizes.Letter.HeightTwips;
        long horizontalMargins = (long)section.Margins.Left + section.Margins.Right;
        long verticalMargins = (long)section.Margins.Top.GetValueOrDefault()
            + section.Margins.Bottom.GetValueOrDefault();
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
        WordImage target = (pageHost ?? document.AddParagraph()).InsertImage(image, source.FileName,
            width, height, WordImageTextWrapping.Square,
            source.AccessibilityDescription ?? "Image imported from Pages");
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
        if (source.HasFirstPageTemplate) {
            WordHeader header = target.GetOrCreateHeader(WordHeaderFooterType.First);
            WordFooter footer = target.GetOrCreateFooter(WordHeaderFooterType.First);
            foreach (IWorkTextContent content in source.FirstPageHeaderContents) {
                AddRichText(content, header.AddParagraph, nativeLists, cancellationToken: cancellationToken);
            }
            foreach (IWorkTextContent content in source.FirstPageFooterContents) {
                AddRichText(content, footer.AddParagraph, nativeLists, cancellationToken: cancellationToken);
            }
        }
        if (hasAnyEvenPageTemplate) {
            WordHeader header = target.GetOrCreateHeader(WordHeaderFooterType.Even);
            WordFooter footer = target.GetOrCreateFooter(WordHeaderFooterType.Even);
            IReadOnlyList<IWorkTextContent> headerContents = source.HasEvenPageTemplate
                ? source.EvenPageHeaderContents
                : source.DefaultPageHeaderContents;
            IReadOnlyList<IWorkTextContent> footerContents = source.HasEvenPageTemplate
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

    private static WordTable? AddTable(WordDocument document, IWorkTable source,
        IWorkNativeListCatalog nativeLists,
        WordParagraph? pageHost, WordTable? tableHost, CancellationToken cancellationToken) {
        if (source.RowCount == 0 || source.ColumnCount == 0) return null;
        WordTable table = pageHost == null
            ? document.AddTable(source.RowCount, source.ColumnCount, WordTableStyle.TableGrid)
            : document.CreateTable(source.RowCount, source.ColumnCount, WordTableStyle.TableGrid);
        table.Description = source.AccessibilityDescription;
        if (source.DefaultColumnWidth is > 0 || source.ColumnWidths.Count > 0) {
            List<int> widths = table.ColumnWidth;
            table.ColumnWidthType = WordTableWidthUnit.Dxa;
            for (int column = 1; column <= source.ColumnCount; column++) {
                cancellationToken.ThrowIfCancellationRequested();
                if (source.GetColumnWidth(column) is double width) widths[column - 1] = ToSignedTwips(width);
            }
            table.ColumnWidth = widths;
        }
        for (int row = 1; row <= source.RowCount; row++) {
            cancellationToken.ThrowIfCancellationRequested();
            if (source.GetRowHeight(row) is double height) {
                if (source.AutoResizeRows == true) table.Rows[row - 1].MinimumHeight = ToSignedTwips(height);
                else table.Rows[row - 1].Height = ToSignedTwips(height);
            }
        }
        foreach (IWorkTableCell sourceCell in source.Cells) {
            cancellationToken.ThrowIfCancellationRequested();
            WordTableCell target = table.Rows[sourceCell.Row - 1].Cells[sourceCell.Column - 1];
            if (sourceCell.Fill is { } fill) {
                if (fill.Color is { } color) target.ShadingFillColorHex = color.RgbHex;
                else target.ShadingPattern = WordShadingPattern.Nil;
            }
            bool header = sourceCell.Row <= source.HeaderRowCount
                || sourceCell.Column <= source.HeaderColumnCount
                || sourceCell.Row > source.RowCount - source.FooterRowCount;
            if (sourceCell.RichText is { Paragraphs.Count: > 0 } richText) {
                bool first = true;
                AddRichText(richText, _ => {
                    WordParagraph paragraph = target.AddParagraph(string.Empty,
                        removeExistingParagraphs: first);
                    first = false;
                    return paragraph;
                }, nativeLists, forceBold: header, cancellationToken: cancellationToken);
            } else {
                WordParagraph paragraph = target.AddParagraph(CellText(sourceCell),
                    removeExistingParagraphs: true);
                if (header) paragraph.Bold = true;
            }
        }
        foreach (IWorkTableMergeRange merge in source.MergedRanges) {
            cancellationToken.ThrowIfCancellationRequested();
            table.MergeCells(merge.FirstRow - 1, merge.FirstColumn - 1,
                merge.LastRow - merge.FirstRow + 1, merge.LastColumn - merge.FirstColumn + 1);
        }
        for (int row = 0; row < Math.Min(source.HeaderRowCount, table.Rows.Count); row++) {
            table.Rows[row].RepeatHeaderRowAtTheTopOfEachPage = true;
        }
        if (tableHost != null) tableHost._table.InsertAfterSelf(table._table);
        else if (pageHost != null) document.InsertTableAfter(pageHost, table);
        return table;
    }

    private static WordTextBox AddRichTextBox(WordDocument document, IWorkTextBox source,
        IWorkNativeListCatalog nativeLists, WordParagraph? pageHost, CancellationToken cancellationToken) {
        WordTextBox textBox = pageHost == null
            ? document.AddTextBox(string.Empty)
            : pageHost.AddTextBox(string.Empty, WordImageTextWrapping.Square);
        if (source.Geometry is { } geometry) {
            textBox.HorizontalPositionRelativeFrom = WordHorizontalRelativePosition.Page;
            textBox.VerticalPositionRelativeFrom = WordVerticalRelativePosition.Page;
            textBox.HorizontalPositionOffset = ToEmusInt32(geometry.LeftPoints);
            textBox.VerticalPositionOffset = ToEmusInt32(geometry.TopPoints);
            textBox.Width = ToEmusInt64(geometry.WidthPoints);
            textBox.Height = ToEmusInt64(geometry.HeightPoints);
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

    private static string CellText(IWorkTableCell cell) {
        return cell.Kind == IWorkCellKind.Formula && cell.Value != null
            ? cell.CachedDisplayText
            : cell.DisplayText;
    }

    private static void AddRichText(IWorkTextContent content, Func<string, WordParagraph> addParagraph,
        IWorkNativeListCatalog nativeLists,
        Func<WordParagraph>? addPageBreak = null,
        Action<IWorkParagraphBreakKind>? addSectionBreak = null,
        bool forceBold = false, CancellationToken cancellationToken = default,
        Action<WordParagraph, IWorkTextRun>? addInlineObject = null) {
        ulong? previousListIdentifier = null;
        bool hasPreviousListParagraph = false;
        foreach (IWorkTextParagraph sourceParagraph in content.Paragraphs) {
            cancellationToken.ThrowIfCancellationRequested();
            WordParagraph paragraph = addParagraph(string.Empty);
            ApplyParagraphStyle(paragraph, sourceParagraph);
            if (forceBold) paragraph.Bold = true;
            if (sourceParagraph.ListLevel >= 0) {
                bool startsNewList = !hasPreviousListParagraph
                    || sourceParagraph.ListIdentifier != previousListIdentifier;
                nativeLists.Apply(paragraph, sourceParagraph.ListLevel,
                    sourceParagraph.ListLabel, startsNewList);
                previousListIdentifier = sourceParagraph.ListIdentifier;
                hasPreviousListParagraph = true;
            } else {
                previousListIdentifier = null;
                hasPreviousListParagraph = false;
            }
            foreach (IWorkTextRun sourceRun in sourceParagraph.Runs) {
                cancellationToken.ThrowIfCancellationRequested();
                if (sourceRun.InlineObject != null) addInlineObject?.Invoke(paragraph, sourceRun);
                else AddStyledTextRun(paragraph, sourceRun, forceBold);
            }
            if (sourceParagraph.BreakKind == IWorkParagraphBreakKind.Page) addPageBreak?.Invoke();
            else if (sourceParagraph.BreakKind is IWorkParagraphBreakKind.Section
                     or IWorkParagraphBreakKind.Layout) addSectionBreak?.Invoke(sourceParagraph.BreakKind);
        }
    }

    private static void AddStyledTextRun(WordParagraph paragraph, IWorkTextRun sourceRun,
        bool forceBold = false) {
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
            ApplyTextStyle(run, sourceRun.Style);
            if (forceBold) run.Bold = true;
        }
    }

    private static void ApplyParagraphStyle(WordParagraph paragraph, IWorkTextParagraph source) {
        IWorkParagraphStyle style = source.Style;
        paragraph.BiDi = OfficeTextElements.ResolveBaseDirection(source.Text)
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
        paragraph.IndentationFirstLinePoints = style.FirstLineIndentPoints;
        paragraph.IndentationBeforePoints = style.LeftIndentPoints;
        paragraph.IndentationAfterPoints = style.RightIndentPoints;
        paragraph.LineSpacingBeforePoints = style.SpaceBeforePoints;
        paragraph.LineSpacingAfterPoints = style.SpaceAfterPoints;
        if (style.PageBreakBefore.HasValue) paragraph.PageBreakBefore = style.PageBreakBefore.Value;
        if (style.KeepWithNext.HasValue) paragraph.KeepWithNext = style.KeepWithNext.Value;
        if (style.KeepLinesTogether.HasValue) paragraph.KeepLinesTogether = style.KeepLinesTogether.Value;
    }

    private static void ApplyTextStyle(WordParagraph run, IWorkTextStyle style) {
        if (style.Bold.HasValue) run.Bold = style.Bold.Value;
        if (style.Italic.HasValue) run.Italic = style.Italic.Value;
        if (style.Underline.HasValue) run.Underline = style.Underline.Value ? WordUnderlineStyle.Single : null;
        if (style.Strikethrough.HasValue) run.Strike = style.Strikethrough.Value;
        if (style.FontSizePoints.HasValue) run.FontSizePoints = style.FontSizePoints.Value;
        if (!string.IsNullOrWhiteSpace(style.FontName)) run.FontFamily = style.FontName;
        if (style.Color != null) run.ColorHex = style.Color.RgbHex;
        if (style.BackgroundColor != null) run.RunShadingFillColorHex = style.BackgroundColor.RgbHex;
    }

    private static void ApplyPageLayout(WordSection section, IWorkPageLayout layout) {
        section.PageSettings.Orientation = layout.IsLandscape
            ? OfficePageOrientation.Landscape
            : OfficePageOrientation.Portrait;
        section.PageSettings.Width = ToTwips(layout.WidthPoints);
        section.PageSettings.Height = ToTwips(layout.HeightPoints);
        section.Margins.Left = ToTwips(layout.LeftMarginPoints);
        section.Margins.Right = ToTwips(layout.RightMarginPoints);
        section.Margins.Top = checked((int)ToTwips(layout.TopMarginPoints));
        section.Margins.Bottom = checked((int)ToTwips(layout.BottomMarginPoints));
        section.Margins.HeaderDistance = ToTwips(layout.HeaderMarginPoints);
        section.Margins.FooterDistance = ToTwips(layout.FooterMarginPoints);
    }

    private sealed class IWorkNativeListCatalog {
        private readonly WordDocument _document;
        private readonly HashSet<int> _observedLevels = new();
        private WordList? _current;

        internal IWorkNativeListCatalog(WordDocument document) {
            _document = document;
        }

        internal void Apply(WordParagraph paragraph, int level, string? label, bool startsNewList) {
            if (startsNewList || _current == null) {
                _current = WordList.AddCustomList(_document);
                _observedLevels.Clear();
            }
            WordList list = _current;
            WordListLevelKind levelKind = Classify(label);
            bool targetLevelExists = list.Numbering.Levels.Count > level;
            while (list.Numbering.Levels.Count <= level) {
                list.Numbering.AddLevel(new WordListLevel(levelKind));
            }
            bool firstObservation = _observedLevels.Add(level);
            if (firstObservation && targetLevelExists) {
                ReplacePlaceholderLevel(list, level, levelKind);
            }
            if (firstObservation && TryParseStart(label, levelKind, out int start) && start != 1) {
                list.Numbering.Levels[level].SetStartNumberingValue(start);
            }
            if (firstObservation && levelKind == WordListLevelKind.Bullet
                && !string.IsNullOrWhiteSpace(label)) {
                list.Numbering.Levels[level].LevelText = label!.Trim();
            }
            if (firstObservation && levelKind != WordListLevelKind.Bullet
                && IsParenthesized(label)) {
                list.Numbering.Levels[level].LevelText = "(%"
                    + (level + 1).ToString(System.Globalization.CultureInfo.InvariantCulture) + ")";
            }
            OpenXmlParagraphProperties properties = paragraph._paragraph.ParagraphProperties
                ?? paragraph._paragraph.PrependChild(new OpenXmlParagraphProperties());
            properties.NumberingProperties = new OpenXmlNumberingProperties(
                new OpenXmlNumberingLevelReference { Val = level },
                new OpenXmlNumberingId { Val = list.NumberId });
        }

        private static void ReplacePlaceholderLevel(WordList list, int level,
            WordListLevelKind levelKind) {
            WordListLevel current = list.Numbering.Levels[level];
            var replacement = new WordListLevel(levelKind);
            replacement.OpenXmlElement.LevelIndex = level;
            replacement.LevelText = replacement.LevelText.Replace(
                "%CurrentLevel", "%" + (level + 1).ToString(
                    System.Globalization.CultureInfo.InvariantCulture));
            current.OpenXmlElement.InsertAfterSelf(replacement.OpenXmlElement);
            current.Remove();
        }

        private static WordListLevelKind Classify(string? label) {
            if (string.IsNullOrWhiteSpace(label)) return WordListLevelKind.Bullet;
            string marker = label!.Trim();
            bool bracket = marker.EndsWith(")", StringComparison.Ordinal);
            bool dot = marker.EndsWith(".", StringComparison.Ordinal);
            string token = MarkerToken(marker);
            if (token.All(char.IsDigit)) {
                return bracket ? WordListLevelKind.DecimalBracket
                    : dot ? WordListLevelKind.DecimalDot : WordListLevelKind.Decimal;
            }
            bool roman = token.Length > 0
                && token.All(character => "ivxlcdmIVXLCDM".IndexOf(character) >= 0)
                && (token.Length > 1 || "ivxIVX".IndexOf(token[0]) >= 0);
            if (roman) {
                bool upper = token.All(char.IsUpper);
                return upper
                    ? bracket ? WordListLevelKind.UpperRomanBracket
                        : dot ? WordListLevelKind.UpperRomanDot : WordListLevelKind.UpperRoman
                    : bracket ? WordListLevelKind.LowerRomanBracket
                        : dot ? WordListLevelKind.LowerRomanDot : WordListLevelKind.LowerRoman;
            }
            if (token.Length > 0 && token.All(char.IsLetter)) {
                bool upper = token.All(char.IsUpper);
                return upper
                    ? bracket ? WordListLevelKind.UpperLetterBracket
                        : dot ? WordListLevelKind.UpperLetterDot : WordListLevelKind.UpperLetter
                    : bracket ? WordListLevelKind.LowerLetterBracket
                        : dot ? WordListLevelKind.LowerLetterDot : WordListLevelKind.LowerLetter;
            }
            return WordListLevelKind.Bullet;
        }

        private static bool TryParseStart(string? label, WordListLevelKind kind, out int start) {
            start = 1;
            if (string.IsNullOrWhiteSpace(label)) return false;
            string token = MarkerToken(label!.Trim());
            switch (kind) {
                case WordListLevelKind.Decimal:
                case WordListLevelKind.DecimalDot:
                case WordListLevelKind.DecimalBracket:
                    return int.TryParse(token, System.Globalization.NumberStyles.None,
                        System.Globalization.CultureInfo.InvariantCulture, out start) && start > 0;
                case WordListLevelKind.UpperLetter:
                case WordListLevelKind.UpperLetterDot:
                case WordListLevelKind.UpperLetterBracket:
                case WordListLevelKind.LowerLetter:
                case WordListLevelKind.LowerLetterDot:
                case WordListLevelKind.LowerLetterBracket:
                    start = 0;
                    foreach (char character in token) {
                        int digit = char.ToUpperInvariant(character) - 'A' + 1;
                        if (digit < 1 || digit > 26 || start > (int.MaxValue - digit) / 26) return false;
                        start = start * 26 + digit;
                    }
                    return start > 0;
                case WordListLevelKind.UpperRoman:
                case WordListLevelKind.UpperRomanDot:
                case WordListLevelKind.UpperRomanBracket:
                case WordListLevelKind.LowerRoman:
                case WordListLevelKind.LowerRomanDot:
                case WordListLevelKind.LowerRomanBracket:
                    return TryParseRoman(token, out start);
                default:
                    return false;
            }
        }

        private static bool TryParseRoman(string token, out int value) {
            value = 0;
            int previous = 0;
            for (int index = token.Length - 1; index >= 0; index--) {
                int current = char.ToUpperInvariant(token[index]) switch {
                    'I' => 1, 'V' => 5, 'X' => 10, 'L' => 50,
                    'C' => 100, 'D' => 500, 'M' => 1000, _ => 0
                };
                if (current == 0) return false;
                int delta = current < previous ? -current : current;
                if (delta > 0 && value > int.MaxValue - delta
                    || delta < 0 && value < int.MinValue - delta) return false;
                value += delta;
                if (current > previous) previous = current;
            }
            if (value <= 0) return false;
            bool upper = token.All(character => character is >= 'A' and <= 'Z');
            string canonical = FormatRoman(value);
            return string.Equals(token, upper ? canonical : canonical.ToLowerInvariant(),
                StringComparison.Ordinal);
        }

        private static string FormatRoman(int value) {
            var builder = new System.Text.StringBuilder();
            foreach ((int Number, string Token) part in new[] {
                         (1000, "M"), (900, "CM"), (500, "D"), (400, "CD"),
                         (100, "C"), (90, "XC"), (50, "L"), (40, "XL"),
                         (10, "X"), (9, "IX"), (5, "V"), (4, "IV"), (1, "I")
                     }) {
                while (value >= part.Number) {
                    builder.Append(part.Token);
                    value -= part.Number;
                }
            }
            return builder.ToString();
        }

        internal static bool CanPreserveStart(string? label) {
            if (!string.IsNullOrWhiteSpace(label)) {
                string token = MarkerToken(label!.Trim());
                if (token.All(char.IsLetter)
                    && token.Any(char.IsUpper) && token.Any(char.IsLower)) return false;
            }
            WordListLevelKind kind = Classify(label);
            return kind == WordListLevelKind.Bullet || TryParseStart(label, kind, out _);
        }

        private static bool IsParenthesized(string? marker) {
            if (string.IsNullOrWhiteSpace(marker)) return false;
            string trimmed = marker!.Trim();
            return trimmed.Length >= 2 && trimmed[0] == '('
                && trimmed[trimmed.Length - 1] == ')';
        }

        private static string MarkerToken(string marker) {
            string trimmed = marker.Trim();
            if (trimmed.Length >= 2 && trimmed[0] == '('
                && trimmed[trimmed.Length - 1] == ')') {
                return trimmed.Substring(1, trimmed.Length - 2).Trim();
            }
            return trimmed.TrimEnd('.', ')', ' ');
        }
    }

    private static uint ToTwips(double points) {
        double value = Math.Round(points * 20d, MidpointRounding.AwayFromZero);
        if (value < 0 || value > uint.MaxValue) throw new InvalidDataException("A Pages page measurement exceeds the DOCX range.");
        return (uint)value;
    }

    private static int ToSignedTwips(double points) {
        double value = Math.Round(points * 20d, MidpointRounding.AwayFromZero);
        if (value < 0 || value > int.MaxValue) throw new InvalidDataException("A Pages table measurement exceeds the DOCX range.");
        return (int)value;
    }
}

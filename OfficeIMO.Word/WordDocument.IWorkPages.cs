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

namespace OfficeIMO.Word;

public partial class WordDocument {
    /// <summary>Loads a Pages source into the normal editable Word model, using a visual preview only when requested or necessary.</summary>
    public static WordDocument LoadPages(string path, IWorkReadOptions? options = null) =>
        LoadPagesWithReport(path, options).Document;

    /// <summary>Loads a Pages stream into the normal editable Word model, using a visual preview only when requested or necessary.</summary>
    public static WordDocument LoadPages(Stream stream, IWorkReadOptions? options = null) =>
        LoadPagesWithReport(stream, options).Document;

    /// <summary>Loads a Pages source and returns its Word projection, bounded source model, and loss report.</summary>
    public static IWorkPagesLoadResult LoadPagesWithReport(string path, IWorkReadOptions? options = null) {
        if (path == null) throw new ArgumentNullException(nameof(path));
        return ProjectPages(IWorkSourceDocument.Open(path, IWorkDocumentKind.Pages, options));
    }

    /// <summary>Loads a Pages stream and returns its Word projection, bounded source model, and loss report.</summary>
    public static IWorkPagesLoadResult LoadPagesWithReport(Stream stream, IWorkReadOptions? options = null) {
        if (stream == null) throw new ArgumentNullException(nameof(stream));
        return ProjectPages(IWorkSourceDocument.Open(stream, IWorkDocumentKind.Pages, options));
    }

    private static IWorkPagesLoadResult ProjectPages(IWorkSourceDocument source) {
        IWorkImportMode mode = source.RequestedImportMode;
        IWorkPreviewAsset? preview = mode == IWorkImportMode.VisualOnly
            ? source.PreferredRasterPreview
            : null;
        if (mode == IWorkImportMode.VisualOnly && preview == null) {
            throw new NotSupportedException("The Pages source has no embedded raster preview.");
        }

        IWorkPagesProjection projection = source.ReadPages();
        string? destinationLimitation = mode == IWorkImportMode.VisualOnly
            ? null
            : FindWordProjectionLimitation(projection);
        bool editable = mode != IWorkImportMode.VisualOnly && projection.HasEditableContent
            && destinationLimitation == null;
        if (!editable && mode == IWorkImportMode.EditableOnly) {
            throw new InvalidDataException(destinationLimitation
                ?? "The Pages source has no supported editable content.");
        }

        preview ??= editable ? null : source.PreferredRasterPreview;
        if (!editable && preview == null) {
            throw new NotSupportedException("The Pages source has no supported editable content or embedded raster preview.");
        }

        WordDocument document = Create();
        try {
            if (projection.PageLayout is { } pageLayout && CanApplyPageLayout(pageLayout)) {
                ApplyPageLayout(document.Sections[0], pageLayout);
            }
            if (editable) {
                var nativeLists = new IWorkNativeListCatalog(document);
                (double contentWidth, double contentHeight) = ContentBox(document.Sections[0]);
                var semanticSections = new List<WordSection> { document.Sections[0] };
                AddRichText(projection.Body, document.AddParagraph, nativeLists, document.AddPageBreak,
                    breakKind => {
                        WordSection section = document.AddSection(breakKind == IWorkParagraphBreakKind.Layout
                            ? WordSectionBreakType.Continuous
                            : WordSectionBreakType.NextPage);
                        if (breakKind == IWorkParagraphBreakKind.Section) semanticSections.Add(section);
                    });
                if (projection.PageLayout != null) {
                    foreach (WordSection section in document.Sections) ApplyPageLayout(section, projection.PageLayout);
                }
                for (int drawableIndex = 0; drawableIndex < projection.Drawables.Count; drawableIndex++) {
                    IWorkPagesDrawable sourceDrawable = projection.Drawables[drawableIndex];
                    uint zOrder = checked(251658240U + (uint)drawableIndex);
                    switch (sourceDrawable.Kind) {
                        case IWorkPagesDrawableKind.TextBox:
                            AddRichTextBox(document, sourceDrawable.TextBox!, nativeLists).ZOrder = zOrder;
                            break;
                        case IWorkPagesDrawableKind.Image:
                            AddImage(document, sourceDrawable.Image!, contentWidth, contentHeight).ZOrder = zOrder;
                            break;
                        case IWorkPagesDrawableKind.Table:
                            AddTable(document, sourceDrawable.Table!);
                            break;
                    }
                }
                bool hasAnyEvenPageTemplate = projection.Sections.Any(section => section.HasEvenPageTemplate);
                for (int sectionIndex = 0; sectionIndex < projection.Sections.Count; sectionIndex++) {
                    IWorkPagesSection sourceSection = projection.Sections[sectionIndex];
                    WordSection targetSection = semanticSections[sectionIndex];
                    AddSectionHeadersAndFooters(targetSection, sourceSection, nativeLists,
                        hasAnyEvenPageTemplate);
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
            return new IWorkPagesLoadResult(document, source, projection, projection.CreateImportReport(kind, preview));
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
        double contentWidth, double contentHeight) {
        using var image = new MemoryStream(source.GetBytes(), writable: false);
        double width = source.Geometry?.WidthPoints
            ?? source.PixelWidth.GetValueOrDefault(640) * 72d / 96d;
        double height = source.Geometry?.HeightPoints
            ?? source.PixelHeight.GetValueOrDefault(480) * 72d / 96d;
        if (source.Geometry == null) {
            (width, height) = FitInside(width, height, contentWidth, contentHeight);
        }
        WordImage target = document.AddParagraph().InsertImage(image, source.FileName,
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
        bool hasAnyEvenPageTemplate) {
        if (source.HasDefaultPageTemplate) {
            WordHeader header = target.GetOrCreateHeader(WordHeaderFooterType.Default);
            WordFooter footer = target.GetOrCreateFooter(WordHeaderFooterType.Default);
            foreach (IWorkTextContent content in source.DefaultPageHeaderContents) {
                AddRichText(content, header.AddParagraph, nativeLists);
            }
            foreach (IWorkTextContent content in source.DefaultPageFooterContents) {
                AddRichText(content, footer.AddParagraph, nativeLists);
            }
        }
        if (source.HasFirstPageTemplate) {
            WordHeader header = target.GetOrCreateHeader(WordHeaderFooterType.First);
            WordFooter footer = target.GetOrCreateFooter(WordHeaderFooterType.First);
            foreach (IWorkTextContent content in source.FirstPageHeaderContents) {
                AddRichText(content, header.AddParagraph, nativeLists);
            }
            foreach (IWorkTextContent content in source.FirstPageFooterContents) {
                AddRichText(content, footer.AddParagraph, nativeLists);
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
                AddRichText(content, header.AddParagraph, nativeLists);
            }
            foreach (IWorkTextContent content in footerContents) {
                AddRichText(content, footer.AddParagraph, nativeLists);
            }
        }
    }

    private static void AddTable(WordDocument document, IWorkTable source) {
        if (source.RowCount == 0 || source.ColumnCount == 0) return;
        WordTable table = document.AddTable(source.RowCount, source.ColumnCount, WordTableStyle.TableGrid);
        table.Description = source.AccessibilityDescription;
        if (source.DefaultColumnWidth is > 0) {
            int width = ToSignedTwips(source.DefaultColumnWidth.Value);
            table.ColumnWidthType = WordTableWidthUnit.Dxa;
            table.ColumnWidth = Enumerable.Repeat(width, source.ColumnCount).ToList();
        }
        if (source.DefaultRowHeight is > 0) {
            int height = ToSignedTwips(source.DefaultRowHeight.Value);
            foreach (WordTableRow row in table.Rows) row.Height = height;
        }
        foreach (IWorkTableCell sourceCell in source.Cells) {
            WordTableCell target = table.Rows[sourceCell.Row - 1].Cells[sourceCell.Column - 1];
            WordParagraph paragraph = target.AddParagraph(CellText(sourceCell), removeExistingParagraphs: true);
            if (sourceCell.Row <= source.HeaderRowCount || sourceCell.Column <= source.HeaderColumnCount
                || sourceCell.Row > source.RowCount - source.FooterRowCount) {
                paragraph.Bold = true;
            }
        }
        foreach (IWorkTableMergeRange merge in source.MergedRanges) {
            table.MergeCells(merge.FirstRow - 1, merge.FirstColumn - 1,
                merge.LastRow - merge.FirstRow + 1, merge.LastColumn - merge.FirstColumn + 1);
        }
        for (int row = 0; row < Math.Min(source.HeaderRowCount, table.Rows.Count); row++) {
            table.Rows[row].RepeatHeaderRowAtTheTopOfEachPage = true;
        }
    }

    private static WordTextBox AddRichTextBox(WordDocument document, IWorkTextBox source,
        IWorkNativeListCatalog nativeLists) {
        WordTextBox textBox = document.AddTextBox(string.Empty);
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
        }, nativeLists);
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

    private static string? FindWordProjectionLimitation(IWorkPagesProjection projection) {
        const long MaximumDestinationTableCells = 1_000_000;
        long destinationTableCells = 0;
        if (projection.TextBoxObjects.Any(textBox => textBox.Hyperlink != null)
            || projection.Images.Any(image => image.Hyperlink != null)) {
            return "Pages contains a drawable hyperlink that cannot be represented by the DOCX owner.";
        }
        if (projection.Sections.SelectMany(section => section.HeaderContents)
            .Concat(projection.Sections.SelectMany(section => section.FooterContents))
            .Any(HasContainerScopedBreak)) {
            return "Pages headers or footers contain a section, layout, or page break that cannot be represented inside a DOCX header or footer.";
        }
        if (projection.TextBoxObjects.Any(textBox => HasContainerScopedBreak(textBox.Content))) {
            return "A Pages text box contains a section, layout, or page break that cannot be represented inside a DOCX text box.";
        }
        if (projection.Body.Paragraphs
                .Concat(projection.TextBoxObjects.SelectMany(textBox => textBox.Content.Paragraphs))
                .Concat(projection.Sections.SelectMany(section => section.HeaderContents)
                    .Concat(projection.Sections.SelectMany(section => section.FooterContents))
                    .SelectMany(content => content.Paragraphs))
                .SelectMany(paragraph => paragraph.Runs)
                .Any(run => run.Hyperlink != null
                    && !Uri.TryCreate(run.Hyperlink, UriKind.Absolute, out _))) {
            return "Pages contains a text hyperlink that cannot be represented by the DOCX owner.";
        }
        if (projection.Body.Paragraphs
                .Concat(projection.TextBoxObjects.SelectMany(textBox => textBox.Content.Paragraphs))
                .Concat(projection.Sections.SelectMany(section => section.HeaderContents)
                    .Concat(projection.Sections.SelectMany(section => section.FooterContents))
                    .SelectMany(content => content.Paragraphs))
                .Any(paragraph => paragraph.ListLevel > 8)) {
            return "Pages contains a list nesting level outside the DOCX numbering range.";
        }
        if (AllPagesText(projection).SelectMany(content => content.Paragraphs)
                .Any(paragraph => paragraph.ListLevel >= 0
                    && !IWorkNativeListCatalog.CanPreserveStart(paragraph.ListLabel))) {
            return "Pages contains an ordered-list marker that cannot be represented by DOCX numbering.";
        }
        if (projection.PageLayout is { } layout && !CanApplyPageLayout(layout)) {
            return "The Pages page layout exceeds the DOCX measurement range.";
        }
        foreach (IWorkTable table in projection.Tables) {
            long tableCells = (long)table.RowCount * table.ColumnCount;
            if (table.RowCount == 0 || table.ColumnCount == 0) {
                return $"Pages table '{table.Name}' has no rows or columns and cannot be represented by the DOCX table owner.";
            }
            if (table.ColumnCount > 63) {
                return $"Pages table '{table.Name}' exceeds Word's supported 63-column table layout.";
            }
            if (table.RowCount > 32_767 || tableCells > 100_000) {
                return $"Pages table '{table.Name}' is too large for bounded DOCX table reconstruction.";
            }
            if (destinationTableCells > MaximumDestinationTableCells - tableCells) {
                return "Pages tables exceed the bounded DOCX destination cell budget.";
            }
            if (table.Cells.Any(cell => cell.Kind == IWorkCellKind.Formula && cell.Value == null)) {
                return $"Pages table '{table.Name}' contains an uncached formula that the DOCX owner cannot evaluate.";
            }
            if (!FitsSignedTwips(table.DefaultRowHeight)
                || !FitsSignedTwips(table.DefaultColumnWidth)) {
                return $"Pages table '{table.Name}' has default sizing outside the DOCX measurement range.";
            }
            if (table.Geometry is { } geometry
                && (Math.Abs(geometry.LeftPoints) > 0.000001d
                    || Math.Abs(geometry.TopPoints) > 0.000001d
                    || Math.Abs(geometry.WidthPoints) > 0.000001d
                    || Math.Abs(geometry.HeightPoints) > 0.000001d
                    || Math.Abs(geometry.RotationDegrees) > 0.000001d)) {
                return $"Pages table '{table.Name}' has positioned, sized, or rotated drawable geometry that the DOCX table owner cannot preserve.";
            }
            destinationTableCells += tableCells;
        }
        foreach (IWorkTextBox textBox in projection.TextBoxObjects) {
            if (textBox.Geometry is { } geometry
                && (!FitsEmuOffset(geometry.LeftPoints) || !FitsEmuOffset(geometry.TopPoints)
                    || !FitsEmuExtent(geometry.WidthPoints) || !FitsEmuExtent(geometry.HeightPoints)
                    || Math.Abs(geometry.RotationDegrees) > 0.000001d)) {
                return "A Pages text box has unsupported rotation or geometry outside the DOCX measurement range.";
            }
        }
        foreach (IWorkImageAsset image in projection.Images) {
            if (image.Geometry is { } geometry
                && (geometry.WidthPoints <= 0 || geometry.HeightPoints <= 0
                    || !FitsEmuExtent(geometry.WidthPoints)
                    || !FitsEmuExtent(geometry.HeightPoints)
                    || !FitsEmuOffset(geometry.LeftPoints)
                    || !FitsEmuOffset(geometry.TopPoints)
                    || Math.Abs(geometry.RotationDegrees) > 0.000001d)) {
                return "A Pages image has unsupported placement, rotation, or extent for the DOCX image owner.";
            }
        }
        foreach (IWorkTextContent content in AllPagesText(projection)) {
            foreach (IWorkTextParagraph paragraph in content.Paragraphs) {
                IWorkParagraphStyle style = paragraph.Style;
                if (!FitsSignedTwips(style.FirstLineIndentPoints)
                    || !FitsSignedTwips(style.LeftIndentPoints)
                    || !FitsSignedTwips(style.RightIndentPoints)
                    || !FitsUnsignedNullableTwips(style.SpaceBeforePoints)
                    || !FitsUnsignedNullableTwips(style.SpaceAfterPoints)) {
                    return "Pages paragraph formatting exceeds the DOCX measurement range.";
                }
                foreach (IWorkTextRun run in paragraph.Runs) {
                    if (run.Style.Color is { Alpha: < byte.MaxValue }
                        || run.Style.BackgroundColor is { Alpha: < byte.MaxValue }) {
                        return "Pages contains transparent text colors that cannot be represented by the DOCX owner.";
                    }
                    if (run.Style.FontSizePoints is double fontSize
                        && (!IsFinite(fontSize) || fontSize < 0 || fontSize > int.MaxValue / 2d
                            || fontSize * 2d != Math.Round(fontSize * 2d,
                                MidpointRounding.AwayFromZero))) {
                        return "A Pages font size exceeds the DOCX measurement range or half-point precision.";
                    }
                }
            }
        }
        return null;
    }

    private static IEnumerable<IWorkTextContent> AllPagesText(IWorkPagesProjection projection) {
        yield return projection.Body;
        foreach (IWorkPagesSection section in projection.Sections) {
            foreach (IWorkTextContent header in section.HeaderContents) yield return header;
            foreach (IWorkTextContent footer in section.FooterContents) yield return footer;
        }
        foreach (IWorkTextBox textBox in projection.TextBoxObjects) yield return textBox.Content;
    }

    private static bool HasContainerScopedBreak(IWorkTextContent content) =>
        content.Paragraphs.Any(paragraph => paragraph.BreakKind is IWorkParagraphBreakKind.Section
            or IWorkParagraphBreakKind.Layout or IWorkParagraphBreakKind.Page);

    private static bool FitsUnsignedTwips(double points) =>
        IsFinite(points) && points >= 0 && points <= uint.MaxValue / 20d
        && IsExactDestinationUnit(points, 20d);

    private static bool CanApplyPageLayout(IWorkPageLayout layout) =>
        layout.WidthPoints > 0 && layout.HeightPoints > 0
        && layout.LeftMarginPoints + layout.RightMarginPoints < layout.WidthPoints
        && layout.TopMarginPoints + layout.BottomMarginPoints < layout.HeightPoints
        && FitsUnsignedTwips(layout.WidthPoints)
        && FitsUnsignedTwips(layout.HeightPoints)
        && FitsUnsignedTwips(layout.LeftMarginPoints)
        && FitsUnsignedTwips(layout.RightMarginPoints)
        && FitsSignedTwips(layout.TopMarginPoints)
        && FitsSignedTwips(layout.BottomMarginPoints)
        && FitsUnsignedTwips(layout.HeaderMarginPoints)
        && FitsUnsignedTwips(layout.FooterMarginPoints);

    private static bool FitsSignedTwips(double? points) => !points.HasValue
        || IsFinite(points.Value) && Math.Abs(points.Value) <= int.MaxValue / 20d
        && IsExactDestinationUnit(points.Value, 20d);

    private static bool FitsUnsignedNullableTwips(double? points) => !points.HasValue
        || IsFinite(points.Value) && points.Value >= 0 && points.Value <= uint.MaxValue / 20d
        && IsExactDestinationUnit(points.Value, 20d);

    private static bool IsFinite(double value) => !double.IsNaN(value) && !double.IsInfinity(value);

    private static bool FitsEmuOffset(double points) =>
        !double.IsNaN(points) && !double.IsInfinity(points)
        && points >= int.MinValue / 12700d && points <= int.MaxValue / 12700d
        && IsExactDestinationUnit(points, 12700d);

    private static bool FitsEmuExtent(double points) =>
        !double.IsNaN(points) && !double.IsInfinity(points)
        && points >= 0 && points <= long.MaxValue / 12700d
        && IsExactDestinationUnit(points, 12700d);

    private static bool IsExactDestinationUnit(double points, double unitsPerPoint) {
        double scaled = points * unitsPerPoint;
        double rounded = Math.Round(scaled, MidpointRounding.AwayFromZero);
        double sourceFloatTolerance = Math.Max(1e-6d, Math.Abs(scaled) * 1e-7d);
        return Math.Abs(scaled - rounded) <= sourceFloatTolerance;
    }

    private static int ToEmusInt32(double points) => checked((int)Math.Round(points * 12700d,
        MidpointRounding.AwayFromZero));

    private static long ToEmusInt64(double points) => checked((long)Math.Round(points * 12700d,
        MidpointRounding.AwayFromZero));

    private static void AddRichText(IWorkTextContent content, Func<string, WordParagraph> addParagraph,
        IWorkNativeListCatalog nativeLists,
        Func<WordParagraph>? addPageBreak = null,
        Action<IWorkParagraphBreakKind>? addSectionBreak = null) {
        ulong? previousListIdentifier = null;
        bool hasPreviousListParagraph = false;
        foreach (IWorkTextParagraph sourceParagraph in content.Paragraphs) {
            WordParagraph paragraph = addParagraph(string.Empty);
            ApplyParagraphStyle(paragraph, sourceParagraph);
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
                AddStyledTextRun(paragraph, sourceRun);
            }
            if (sourceParagraph.BreakKind == IWorkParagraphBreakKind.Page) addPageBreak?.Invoke();
            else if (sourceParagraph.BreakKind is IWorkParagraphBreakKind.Section
                     or IWorkParagraphBreakKind.Layout) addSectionBreak?.Invoke(sourceParagraph.BreakKind);
        }
    }

    private static void AddStyledTextRun(WordParagraph paragraph, IWorkTextRun sourceRun) {
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

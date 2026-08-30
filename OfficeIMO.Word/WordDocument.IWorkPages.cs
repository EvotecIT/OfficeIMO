using OfficeIMO.IWork;
using OfficeIMO.Word.IWork;
using OpenXmlParagraph = DocumentFormat.OpenXml.Wordprocessing.Paragraph;
using OpenXmlRun = DocumentFormat.OpenXml.Wordprocessing.Run;

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
            if (editable) {
                if (projection.PageLayout != null) ApplyPageLayout(document.Sections[0], projection.PageLayout);
                (double contentWidth, double contentHeight) = ContentBox(document.Sections[0]);
                var semanticSections = new List<WordSection> { document.Sections[0] };
                AddRichText(projection.Body, document.AddParagraph, document.AddPageBreak,
                    breakKind => {
                        WordSection section = document.AddSection(breakKind == IWorkParagraphBreakKind.Layout
                            ? WordSectionBreakType.Continuous
                            : WordSectionBreakType.NextPage);
                        if (breakKind == IWorkParagraphBreakKind.Section) semanticSections.Add(section);
                    });
                if (projection.PageLayout != null) {
                    foreach (WordSection section in document.Sections) ApplyPageLayout(section, projection.PageLayout);
                }
                foreach (IWorkTextBox textBox in projection.TextBoxObjects) AddRichTextBox(document, textBox);
                foreach (IWorkTable sourceTable in projection.Tables) AddTable(document, sourceTable);
                foreach (IWorkImageAsset sourceImage in projection.Images.Where(image =>
                             image.MediaType is "image/png" or "image/jpeg")) {
                    using var image = new MemoryStream(sourceImage.GetBytes(), writable: false);
                    double width = sourceImage.Geometry?.WidthPoints
                        ?? sourceImage.PixelWidth.GetValueOrDefault(640) * 72d / 96d;
                    double height = sourceImage.Geometry?.HeightPoints
                        ?? sourceImage.PixelHeight.GetValueOrDefault(480) * 72d / 96d;
                    (width, height) = FitInside(width, height, contentWidth, contentHeight);
                    document.AddParagraph().AddImage(image, sourceImage.FileName,
                        width, height, WordImageTextWrapping.Square,
                        sourceImage.AccessibilityDescription ?? "Image imported from Pages");
                }
                for (int sectionIndex = 0; sectionIndex < projection.Sections.Count; sectionIndex++) {
                    IWorkPagesSection sourceSection = projection.Sections[sectionIndex];
                    WordSection targetSection = semanticSections[sectionIndex];
                    targetSection.AddHeadersAndFooters();
                    foreach (IWorkTextContent header in sourceSection.HeaderContents) {
                        AddRichText(header, targetSection.Header.Default!.AddParagraph);
                    }
                    foreach (IWorkTextContent footer in sourceSection.FooterContents) {
                        AddRichText(footer, targetSection.Footer.Default!.AddParagraph);
                    }
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

    private static void AddTable(WordDocument document, IWorkTable source) {
        if (source.RowCount == 0 || source.ColumnCount == 0) return;
        WordTable table = document.AddTable(source.RowCount, source.ColumnCount, WordTableStyle.TableGrid);
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
        if (source.HeaderRowCount > 0 && table.Rows.Count > 0) {
            table.RepeatAsHeaderRowAtTheTopOfEachPage = true;
        }
    }

    private static void AddRichTextBox(WordDocument document, IWorkTextBox source) {
        WordTextBox textBox = document.AddTextBox(string.Empty);
        if (source.Geometry is { } geometry) {
            textBox.HorizontalPositionRelativeFrom = WordHorizontalRelativePosition.Page;
            textBox.VerticalPositionRelativeFrom = WordVerticalRelativePosition.Page;
            textBox.HorizontalPositionOffset = ToEmusInt32(geometry.LeftPoints);
            textBox.VerticalPositionOffset = ToEmusInt32(geometry.TopPoints);
            textBox.Width = ToEmusInt64(geometry.WidthPoints);
            textBox.Height = ToEmusInt64(geometry.HeightPoints);
        }

        DocumentFormat.OpenXml.Wordprocessing.TextBoxContent? content = textBox.Content;
        if (content == null) return;
        content.RemoveAllChildren<OpenXmlParagraph>();
        AddRichText(source.Content, value => {
            var paragraph = new OpenXmlParagraph();
            content.Append(paragraph);
            var result = new WordParagraph(document, paragraph, newRun: false);
            if (value.Length > 0) result.AddText(value);
            return result;
        });
        if (!content.Elements<OpenXmlParagraph>().Any()) {
            content.Append(new OpenXmlParagraph(new OpenXmlRun()));
        }
    }

    private static string CellText(IWorkTableCell cell) {
        return cell.Kind == IWorkCellKind.Formula && cell.Value != null
            ? cell.CachedDisplayText
            : cell.DisplayText;
    }

    private static string? FindWordProjectionLimitation(IWorkPagesProjection projection) {
        if (projection.TextBoxObjects.Any(textBox => textBox.Hyperlink != null)
            || projection.Images.Any(image => image.Hyperlink != null)) {
            return "Pages contains a drawable hyperlink that cannot be represented by the DOCX owner.";
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
        if (projection.PageLayout is { } layout
            && (layout.WidthPoints <= 0 || layout.HeightPoints <= 0
                || layout.WidthPoints > uint.MaxValue / 20d || layout.HeightPoints > uint.MaxValue / 20d
                || !FitsUnsignedTwips(layout.LeftMarginPoints)
                || !FitsUnsignedTwips(layout.RightMarginPoints)
                || !FitsSignedTwips(layout.TopMarginPoints)
                || !FitsSignedTwips(layout.BottomMarginPoints)
                || !FitsUnsignedTwips(layout.HeaderMarginPoints)
                || !FitsUnsignedTwips(layout.FooterMarginPoints))) {
            return "The Pages page layout exceeds the DOCX measurement range.";
        }
        foreach (IWorkTable table in projection.Tables) {
            if (table.ColumnCount > 63) {
                return $"Pages table '{table.Name}' exceeds Word's supported 63-column table layout.";
            }
            if (table.RowCount > 32_767 || (long)table.RowCount * table.ColumnCount > 100_000) {
                return $"Pages table '{table.Name}' is too large for bounded DOCX table reconstruction.";
            }
        }
        foreach (IWorkTextBox textBox in projection.TextBoxObjects) {
            if (textBox.Geometry is { } geometry
                && (!FitsEmuOffset(geometry.LeftPoints) || !FitsEmuOffset(geometry.TopPoints)
                    || !FitsEmuExtent(geometry.WidthPoints) || !FitsEmuExtent(geometry.HeightPoints))) {
                return "A Pages text box has geometry outside the DOCX measurement range.";
            }
        }
        foreach (IWorkImageAsset image in projection.Images) {
            if (image.Geometry is { } geometry
                && (geometry.WidthPoints <= 0 || geometry.HeightPoints <= 0)) {
                return "A Pages image has a zero-sized extent that cannot be represented by the DOCX image owner.";
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
                    if (run.Style.FontSizePoints is double fontSize
                        && (!IsFinite(fontSize) || fontSize < 0 || fontSize > int.MaxValue / 2d)) {
                        return "A Pages font size exceeds the DOCX measurement range.";
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

    private static bool FitsUnsignedTwips(double points) =>
        IsFinite(points) && points >= 0 && points <= uint.MaxValue / 20d;

    private static bool FitsSignedTwips(double? points) => !points.HasValue
        || IsFinite(points.Value) && Math.Abs(points.Value) <= int.MaxValue / 20d;

    private static bool FitsUnsignedNullableTwips(double? points) => !points.HasValue
        || IsFinite(points.Value) && points.Value >= 0 && points.Value <= uint.MaxValue / 20d;

    private static bool IsFinite(double value) => !double.IsNaN(value) && !double.IsInfinity(value);

    private static bool FitsEmuOffset(double points) =>
        !double.IsNaN(points) && !double.IsInfinity(points)
        && points >= int.MinValue / 12700d && points <= int.MaxValue / 12700d;

    private static bool FitsEmuExtent(double points) =>
        !double.IsNaN(points) && !double.IsInfinity(points)
        && points >= 0 && points <= long.MaxValue / 12700d;

    private static int ToEmusInt32(double points) => checked((int)Math.Round(points * 12700d,
        MidpointRounding.AwayFromZero));

    private static long ToEmusInt64(double points) => checked((long)Math.Round(points * 12700d,
        MidpointRounding.AwayFromZero));

    private static void AddRichText(IWorkTextContent content, Func<string, WordParagraph> addParagraph,
        Func<WordParagraph>? addPageBreak = null,
        Action<IWorkParagraphBreakKind>? addSectionBreak = null) {
        foreach (IWorkTextParagraph sourceParagraph in content.Paragraphs) {
            WordParagraph paragraph = addParagraph(string.Empty);
            ApplyParagraphStyle(paragraph, sourceParagraph.Style);
            if (sourceParagraph.ListLevel >= 0) {
                string marker = string.IsNullOrEmpty(sourceParagraph.ListLabel)
                    ? "\u2022"
                    : sourceParagraph.ListLabel!;
                paragraph.AddText(marker + " ");
            }
            foreach (IWorkTextRun sourceRun in sourceParagraph.Runs) {
                WordParagraph run;
                if (sourceRun.Hyperlink != null
                    && Uri.TryCreate(sourceRun.Hyperlink, UriKind.Absolute, out Uri? uri)) {
                    paragraph.AddHyperLink(sourceRun.Text, uri);
                    run = paragraph;
                } else {
                    run = paragraph.AddText(sourceRun.Text);
                }
                ApplyTextStyle(run, sourceRun.Style);
            }
            if (sourceParagraph.BreakKind == IWorkParagraphBreakKind.Page) addPageBreak?.Invoke();
            else if (sourceParagraph.BreakKind is IWorkParagraphBreakKind.Section
                     or IWorkParagraphBreakKind.Layout) addSectionBreak?.Invoke(sourceParagraph.BreakKind);
        }
    }

    private static void ApplyParagraphStyle(WordParagraph paragraph, IWorkParagraphStyle style) {
        if (style.Alignment.HasValue) {
            paragraph.ParagraphAlignment = style.Alignment.Value switch {
                IWorkTextAlignment.Center => WordParagraphAlignment.Center,
                IWorkTextAlignment.Right => WordParagraphAlignment.Right,
                IWorkTextAlignment.Justified => WordParagraphAlignment.Both,
                _ => WordParagraphAlignment.Left
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

    private static uint ToTwips(double points) {
        double value = Math.Round(points * 20d, MidpointRounding.AwayFromZero);
        if (value < 0 || value > uint.MaxValue) throw new InvalidDataException("A Pages page measurement exceeds the DOCX range.");
        return (uint)value;
    }
}

using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.OpenDocument;
using OfficeIMO.PowerPoint;

namespace OfficeIMO.PowerPoint.OpenDocument;

/// <summary>Explicit conversions between OfficeIMO PowerPoint and native OpenDocument presentation models.</summary>
public static class PowerPointOpenDocumentConversionExtensions {
    /// <summary>Converts a PowerPoint presentation to an in-memory ODP document.</summary>
    public static OdpPresentation ToOpenDocument(this PowerPointPresentation source,
        PowerPointOpenDocumentConversionOptions? options = null) => source.ToOpenDocumentResult(options).Value;

    /// <summary>Converts a PowerPoint presentation to an in-memory ODP document and reports every lossy mapping.</summary>
    public static OdfConversionResult<OdpPresentation> ToOpenDocumentResult(this PowerPointPresentation source,
        PowerPointOpenDocumentConversionOptions? options = null) {
        if (source == null) throw new ArgumentNullException(nameof(source));
        PowerPointOpenDocumentConversionOptions effective = NormalizeOptions(options);
        OdpPresentation target = OdpPresentation.Create();
        var report = new OdfConversionReport("PPTX", "ODP");
        target.Metadata.Title = source.BuiltinDocumentProperties.Title;
        target.PageWidth = OdfLength.Points(source.SlideSize.WidthPoints);
        target.PageHeight = OdfLength.Points(source.SlideSize.HeightPoints);

        int textBoxes = 0, paragraphs = 0, textRuns = 0, pictures = 0, tables = 0, autoShapes = 0;
        int notes = 0, transitions = 0, backgrounds = 0, unsupportedBackgrounds = 0, unsupportedShapes = 0, unsupportedPictures = 0;
        int listParagraphs = 0, runDetails = 0, transformedShapes = 0;
        for (int slideIndex = 0; slideIndex < source.Slides.Count; slideIndex++) {
            PowerPointSlide sourceSlide = source.Slides[slideIndex];
            OdpSlide targetSlide = target.AddSlide("Slide" + (slideIndex + 1).ToString(System.Globalization.CultureInfo.InvariantCulture));
            targetSlide.Hidden = sourceSlide.Hidden;
            MapBackground(sourceSlide, targetSlide, ref backgrounds, ref unsupportedBackgrounds);
            if (MapTransition(sourceSlide.Transition, targetSlide)) transitions++;

            foreach (PowerPointShape shape in sourceSlide.Shapes.OrderBy(item => item.DrawingOrder)) {
                if (shape is PowerPointTextBox textBox) {
                    OdpTextBox converted = targetSlide.AddTextBox(ToOdfRect(textBox), null, textBox.Name);
                    CopyShapeAppearance(textBox, converted, effective);
                    foreach (PowerPointParagraph paragraph in textBox.Paragraphs) {
                        OdpParagraph targetParagraph = converted.AddParagraph();
                        IReadOnlyList<PowerPointTextRun> runs = paragraph.Runs;
                        if (runs.Count == 0) targetParagraph.Text = paragraph.Text;
                        else {
                            foreach (PowerPointTextRun run in runs) {
                                OdpRun targetRun = targetParagraph.AddRun(run.Text);
                                if (effective.IncludeBasicFormatting) {
                                    targetRun.Bold = run.Bold ? true : (bool?)null;
                                    targetRun.Italic = run.Italic ? true : (bool?)null;
                                    if (run.FontSize.HasValue) targetRun.FontSize = OdfLength.Points(run.FontSize.Value);
                                    targetRun.FontFamily = run.FontName;
                                    if (!string.IsNullOrWhiteSpace(run.Color)) targetRun.Color = ParseColor(run.Color);
                                }
                                if (run.Underline || run.Strikethrough || run.Hyperlink != null || run.HighlightColor != null) runDetails++;
                                textRuns++;
                            }
                        }
                        if (paragraph.BulletCharacter != null || paragraph.IsNumbered) listParagraphs++;
                        paragraphs++;
                    }
                    textBoxes++;
                } else if (shape is PowerPointPicture picture) {
                    if (!effective.IncludeImages) { unsupportedPictures++; continue; }
                    try {
                        OdpImage converted = targetSlide.AddImage(picture.GetImageBytes(), FileNameForContentType(picture.ContentType), ToOdfRect(picture), picture.Name);
                        CopyShapeAppearance(picture, converted, effective);
                        if (picture.CropLeftRatio > 0D || picture.CropTopRatio > 0D || picture.CropRightRatio > 0D || picture.CropBottomRatio > 0D) {
                            converted.Crop = new OdfInsets(
                                OdfLength.Points(picture.CropTopRatio * picture.HeightPoints),
                                OdfLength.Points(picture.CropRightRatio * picture.WidthPoints),
                                OdfLength.Points(picture.CropBottomRatio * picture.HeightPoints),
                                OdfLength.Points(picture.CropLeftRatio * picture.WidthPoints));
                        }
                        pictures++;
                    } catch (Exception exception) when (exception is InvalidOperationException || exception is NotSupportedException) {
                        unsupportedPictures++;
                    }
                } else if (shape is PowerPointTable table) {
                    int rowCount = Math.Max(1, table.Rows);
                    if (rowCount > effective.MaxTableRows) {
                        throw new InvalidDataException($"PowerPoint table rows ({rowCount}) exceed the configured conversion limit ({effective.MaxTableRows}).");
                    }
                    int columnCount = Math.Max(1, table.Columns);
                    if (columnCount > effective.MaxTableColumns) {
                        throw new InvalidDataException($"PowerPoint table columns ({columnCount}) exceed the configured conversion limit ({effective.MaxTableColumns}).");
                    }
                    OdpTable converted = targetSlide.AddTable(ToOdfRect(table), rowCount, columnCount, table.Name);
                    CopyShapeAppearance(table, converted, effective);
                    var merges = new List<(int Row, int Column, int RowSpan, int ColumnSpan)>();
                    for (int row = 0; row < table.Rows; row++) {
                        for (int column = 0; column < table.Columns; column++) {
                            PowerPointTableCell cell = table.GetCell(row, column);
                            if (cell.IsMergedCell) continue;
                            converted.Cell(row, column).Text = cell.Text;
                            if (cell.IsMergeAnchor) merges.Add((row, column, cell.Merge.rows, cell.Merge.columns));
                        }
                    }
                    foreach (var merge in merges) converted.Merge(merge.Row, merge.Column, merge.RowSpan, merge.ColumnSpan);
                    tables++;
                } else if (shape is PowerPointAutoShape autoShape) {
                    OdpShape converted;
                    if (autoShape.ShapeType == ShapeTypeValues.Ellipse) converted = targetSlide.AddEllipse(ToOdfRect(autoShape), autoShape.Name);
                    else if (autoShape.ShapeType == ShapeTypeValues.Line) {
                        converted = targetSlide.AddLine(OdfLength.Points(autoShape.LeftPoints), OdfLength.Points(autoShape.TopPoints),
                            OdfLength.Points(autoShape.RightPoints), OdfLength.Points(autoShape.BottomPoints), autoShape.Name);
                    } else {
                        converted = targetSlide.AddRectangle(ToOdfRect(autoShape), autoShape.Name);
                        if (autoShape.ShapeType != ShapeTypeValues.Rectangle) transformedShapes++;
                    }
                    CopyShapeAppearance(autoShape, converted, effective);
                    autoShapes++;
                } else if (shape is PowerPointConnectionShape connection) {
                    OdpLine converted = targetSlide.AddLine(OdfLength.Points(connection.LeftPoints), OdfLength.Points(connection.TopPoints),
                        OdfLength.Points(connection.RightPoints), OdfLength.Points(connection.BottomPoints), connection.Name);
                    CopyShapeAppearance(connection, converted, effective);
                    autoShapes++;
                    transformedShapes++;
                } else {
                    unsupportedShapes++;
                }
                if (shape.Rotation.HasValue || shape.HorizontalFlip.HasValue || shape.VerticalFlip.HasValue) transformedShapes++;
            }

            if (effective.IncludeSpeakerNotes && sourceSlide.HasSpeakerNotes) {
                string noteText = sourceSlide.GetSpeakerNotesText();
                if (noteText.Length > 0) {
                    foreach (string paragraph in SplitParagraphs(noteText)) targetSlide.GetOrCreateSpeakerNotes().AddParagraph(paragraph);
                    notes++;
                }
            }
        }

        AddConverted(report, "slides", source.Slides.Count);
        AddConverted(report, "text-boxes", textBoxes);
        AddConverted(report, "paragraphs", paragraphs);
        AddConverted(report, "text-runs", textRuns);
        AddConverted(report, "images", pictures);
        AddConverted(report, "tables", tables);
        AddConverted(report, "basic-shapes", autoShapes);
        AddConverted(report, "speaker-notes", notes);
        AddConverted(report, "solid-backgrounds", backgrounds);
        AddUnsupported(report, "slide-backgrounds", unsupportedBackgrounds, "Image, gradient, theme, and unsupported backgrounds are not translated.");
        if (transitions > 0) report.Add("slide-transitions", OdfConversionMappingStatus.Approximated, transitions,
            "Common transition families are mapped without PowerPoint-specific speed and timing metadata.");
        if (listParagraphs > 0) report.Add("text-lists", OdfConversionMappingStatus.Approximated, listParagraphs,
            "List text is retained as paragraphs; PowerPoint bullet and numbering definitions are not translated.");
        AddUnsupported(report, "run-format-details", runDetails, "Underline, strike, highlight, and run hyperlinks are not translated.");
        AddUnsupported(report, "shape-transforms", transformedShapes, "Complex geometry, rotation, flips, and connector semantics are approximated or omitted.");
        AddUnsupported(report, "images", unsupportedPictures, "Images disabled by options or unavailable from an embedded image part were skipped.");
        AddUnsupported(report, "shapes", unsupportedShapes, "Charts, SmartArt, media, groups, and other advanced drawing shapes are not translated.");
        report.Add("masters-layouts", OdfConversionMappingStatus.Approximated, source.Slides.Count,
            "Slide content is placed on one default ODP master and blank layout.");
        AddAdvancedPowerPointFindings(source.InspectFeatures(), report);
        return new OdfConversionResult<OdpPresentation>(target, report);
    }

    /// <summary>Converts an ODP document to an in-memory PowerPoint presentation.</summary>
    public static PowerPointPresentation ToPowerPointPresentation(this OdpPresentation source,
        PowerPointOpenDocumentConversionOptions? options = null) => source.ToPowerPointPresentationResult(options).Value;

    /// <summary>Converts an ODP document to an in-memory PowerPoint presentation and reports every lossy mapping.</summary>
    public static OdfConversionResult<PowerPointPresentation> ToPowerPointPresentationResult(this OdpPresentation source,
        PowerPointOpenDocumentConversionOptions? options = null) {
        if (source == null) throw new ArgumentNullException(nameof(source));
        PowerPointOpenDocumentConversionOptions effective = NormalizeOptions(options);
        PowerPointPresentation target = PowerPointPresentation.Create();
        var report = new OdfConversionReport("ODP", "PPTX");
        target.BuiltinDocumentProperties.Title = source.Metadata.Title;
        target.SlideSize.SetSizePoints(source.PageWidth.ToPoints(), source.PageHeight.ToPoints());

        int textBoxes = 0, paragraphs = 0, textRuns = 0, pictures = 0, tables = 0, basicShapes = 0;
        int notes = 0, transitions = 0, unsupportedTransitions = 0, unsupportedShapes = 0, unsupportedPictures = 0, transformedShapes = 0;
        int listParagraphs = 0, approximatedRuns = 0;
        foreach (OdpSlide sourceSlide in source.Slides) {
            PowerPointSlide targetSlide = target.AddSlide();
            targetSlide.Hidden = sourceSlide.Hidden;
            OdfColor? backgroundColor = sourceSlide.BackgroundColor;
            if (!backgroundColor.HasValue && !string.IsNullOrWhiteSpace(sourceSlide.MasterPageName)) {
                backgroundColor = source.MasterPages.FirstOrDefault(master =>
                    string.Equals(master.Name, sourceSlide.MasterPageName, StringComparison.Ordinal))?.BackgroundColor;
            }
            if (backgroundColor.HasValue) targetSlide.BackgroundColor = backgroundColor.Value.ToString().TrimStart('#');
            if (MapTransition(sourceSlide, targetSlide)) transitions++;
            else if (!string.IsNullOrWhiteSpace(sourceSlide.TransitionStyle) || !string.IsNullOrWhiteSpace(sourceSlide.TransitionType)) unsupportedTransitions++;

            foreach (OdpShape shape in sourceSlide.Shapes) {
                if (shape is OdpTextBox textBox) {
                    IReadOnlyList<OdpParagraph> sourceParagraphs = textBox.Paragraphs;
                    listParagraphs += textBox.Lists.Sum(list => list.Items.Count);
                    PowerPointTextBox converted = targetSlide.AddTextBox(string.Empty, ToPowerPointBox(textBox.Bounds));
                    if (sourceParagraphs.Count > 0) converted.SetParagraphs(sourceParagraphs.Select(paragraph => paragraph.Text));
                    converted.Name = textBox.Name;
                    CopyShapeAppearance(textBox, converted, effective);
                    for (int index = 0; index < sourceParagraphs.Count && index < converted.Paragraphs.Count; index++) {
                        OdpParagraph sourceParagraph = sourceParagraphs[index];
                        PowerPointParagraph targetParagraph = converted.Paragraphs[index];
                        IReadOnlyList<OdpRun> runs = sourceParagraph.Runs;
                        if (runs.Count > 0 && string.Equals(string.Concat(runs.Select(run => run.Text)), sourceParagraph.Text, StringComparison.Ordinal)) {
                            IReadOnlyList<PowerPointTextRun> existing = targetParagraph.Runs;
                            PowerPointTextRun first = existing.Count > 0 ? existing[0] : targetParagraph.AddRun(string.Empty);
                            ApplyOdpRun(runs[0], first, effective);
                            for (int runIndex = 1; runIndex < runs.Count; runIndex++) {
                                PowerPointTextRun added = targetParagraph.AddRun(runs[runIndex].Text);
                                ApplyOdpRun(runs[runIndex], added, effective);
                            }
                            textRuns += runs.Count;
                        } else {
                            if (runs.Count > 0) approximatedRuns++;
                            PowerPointTextRun? run = targetParagraph.Runs.FirstOrDefault();
                            if (run != null) {
                                run.Bold = sourceParagraph.Bold == true;
                                if (sourceParagraph.FontSize.HasValue) run.FontSize = checked((int)Math.Round(sourceParagraph.FontSize.Value.ToPoints()));
                            }
                        }
                        paragraphs++;
                    }
                    textBoxes++;
                } else if (shape is OdpImage image) {
                    if (!effective.IncludeImages) {
                        unsupportedPictures++;
                        continue;
                    }
                    try {
                        byte[] imageBytes = image.GetImageBytes();
                        if (!TryGetImagePartType(image.Path, imageBytes, out ImagePartType imageType)) {
                            unsupportedPictures++;
                            continue;
                        }
                        using var stream = new MemoryStream(imageBytes, writable: false);
                        PowerPointPicture converted = targetSlide.AddPicture(stream, imageType, ToPowerPointBox(image.Bounds));
                        converted.Name = image.Name;
                        CopyShapeAppearance(image, converted, effective);
                        if (image.Crop.HasValue) ApplyOdpCrop(image, converted);
                        pictures++;
                    } catch (Exception exception) when (exception is NotSupportedException || exception is InvalidDataException ||
                        exception is ArgumentException) {
                        unsupportedPictures++;
                    }
                } else if (shape is OdpTable table) {
                    int rowCount = Math.Max(1, table.Rows.Count);
                    if (rowCount > effective.MaxTableRows) {
                        throw new InvalidDataException($"ODP table rows ({rowCount}) exceed the configured conversion limit ({effective.MaxTableRows}).");
                    }
                    int columnCount = Math.Max(1, table.Rows.Select(row => row.Cells.Count).DefaultIfEmpty(1).Max());
                    if (columnCount > effective.MaxTableColumns) {
                        throw new InvalidDataException($"ODP table columns ({columnCount}) exceed the configured conversion limit ({effective.MaxTableColumns}).");
                    }
                    PowerPointTable converted = targetSlide.AddTable(rowCount, columnCount, ToPowerPointBox(table.Bounds));
                    converted.Name = table.Name;
                    CopyShapeAppearance(table, converted, effective);
                    var merges = new List<(int Row, int Column, int RowSpan, int ColumnSpan)>();
                    for (int row = 0; row < table.Rows.Count; row++) {
                        IReadOnlyList<OdpTableCell> cells = table.Rows[row].Cells;
                        for (int column = 0; column < cells.Count; column++) {
                            OdpTableCell cell = cells[column];
                            if (cell.IsCovered) continue;
                            converted.GetCell(row, column).Text = cell.Text;
                            if (cell.RowSpan > 1 || cell.ColumnSpan > 1) merges.Add((row, column, cell.RowSpan, cell.ColumnSpan));
                        }
                    }
                    foreach (var merge in merges) converted.MergeCells(merge.Row, merge.Column,
                        merge.Row + merge.RowSpan - 1, merge.Column + merge.ColumnSpan - 1);
                    tables++;
                } else if (shape is OdpRectangle rectangle) {
                    PowerPointAutoShape converted = targetSlide.AddRectanglePoints(rectangle.Bounds.X.ToPoints(), rectangle.Bounds.Y.ToPoints(),
                        rectangle.Bounds.Width.ToPoints(), rectangle.Bounds.Height.ToPoints(), rectangle.Name);
                    CopyShapeAppearance(rectangle, converted, effective);
                    basicShapes++;
                } else if (shape is OdpEllipse ellipse) {
                    PowerPointAutoShape converted = targetSlide.AddEllipsePoints(ellipse.Bounds.X.ToPoints(), ellipse.Bounds.Y.ToPoints(),
                        ellipse.Bounds.Width.ToPoints(), ellipse.Bounds.Height.ToPoints(), ellipse.Name);
                    CopyShapeAppearance(ellipse, converted, effective);
                    basicShapes++;
                } else if (shape is OdpLine line) {
                    PowerPointAutoShape converted = targetSlide.AddLinePoints(line.X1.ToPoints(), line.Y1.ToPoints(), line.X2.ToPoints(), line.Y2.ToPoints(), line.Name);
                    CopyShapeAppearance(line, converted, effective);
                    basicShapes++;
                } else {
                    unsupportedShapes++;
                }
                if (!string.IsNullOrWhiteSpace(shape.Transform)) transformedShapes++;
            }

            if (effective.IncludeSpeakerNotes && sourceSlide.SpeakerNotes != null) {
                string noteText = string.Join(Environment.NewLine, sourceSlide.SpeakerNotes.Paragraphs.Select(paragraph => paragraph.Text));
                if (noteText.Length > 0) { targetSlide.Notes.Text = noteText; notes++; }
            }
        }

        AddConverted(report, "slides", source.Slides.Count);
        AddConverted(report, "text-boxes", textBoxes);
        AddConverted(report, "paragraphs", paragraphs);
        AddConverted(report, "text-runs", textRuns);
        AddConverted(report, "images", pictures);
        AddConverted(report, "tables", tables);
        AddConverted(report, "basic-shapes", basicShapes);
        AddConverted(report, "speaker-notes", notes);
        if (transitions > 0) report.Add("slide-transitions", OdfConversionMappingStatus.Approximated, transitions,
            "Common ODF transition styles are mapped to PowerPoint transition families.");
        if (listParagraphs > 0) report.Add("text-lists", OdfConversionMappingStatus.Approximated, listParagraphs,
            "ODP list text is retained as paragraphs; PowerPoint bullet and numbering definitions are not translated.");
        if (approximatedRuns > 0) report.Add("inline-formatting", OdfConversionMappingStatus.Approximated, approximatedRuns,
            "Mixed plain text and styled ODP spans are flattened when their exact inline order cannot be represented by the typed surface.");
        AddUnsupported(report, "slide-transitions", unsupportedTransitions, "The ODF transition family is not supported by the PowerPoint adapter.");
        AddUnsupported(report, "images", unsupportedPictures, "Images disabled by options or using an unsupported PowerPoint image format were skipped.");
        AddUnsupported(report, "shapes", unsupportedShapes, "Groups and unsupported ODF drawing elements are not translated.");
        AddUnsupported(report, "shape-transforms", transformedShapes, "Raw ODF transform expressions are not translated.");
        if (source.MasterPages.Count > 1 || source.Layouts.Count > 1) report.Add("masters-layouts", OdfConversionMappingStatus.Approximated,
            source.MasterPages.Count + source.Layouts.Count, "Content is placed on PowerPoint's default master and layout.");
        foreach (OdfFeatureFinding finding in source.InspectFeatures().Findings.Where(item => item.Name != "presentation-transitions")) {
            report.Add("source-" + finding.Name, OdfConversionMappingStatus.Unsupported, finding.Count,
                "The source ODP feature cannot be transferred to PPTX by this adapter.");
        }
        return new OdfConversionResult<PowerPointPresentation>(target, report);
    }

    private static PowerPointOpenDocumentConversionOptions NormalizeOptions(PowerPointOpenDocumentConversionOptions? options) {
        PowerPointOpenDocumentConversionOptions effective = options ?? new PowerPointOpenDocumentConversionOptions();
        if (effective.MaxTableRows <= 0) throw new ArgumentOutOfRangeException(nameof(options), effective.MaxTableRows,
            $"{nameof(PowerPointOpenDocumentConversionOptions.MaxTableRows)} must be positive.");
        if (effective.MaxTableColumns <= 0) throw new ArgumentOutOfRangeException(nameof(options), effective.MaxTableColumns,
            $"{nameof(PowerPointOpenDocumentConversionOptions.MaxTableColumns)} must be positive.");
        return effective;
    }

    private static void ApplyOdpRun(OdpRun source, PowerPointTextRun target, PowerPointOpenDocumentConversionOptions options) {
        target.Text = source.Text;
        if (!options.IncludeBasicFormatting) return;
        target.Bold = source.Bold == true;
        target.Italic = source.Italic == true;
        if (source.FontSize.HasValue) target.FontSize = checked((int)Math.Round(source.FontSize.Value.ToPoints()));
        target.FontName = source.FontFamily;
        if (source.Color.HasValue) target.Color = source.Color.Value.ToString().TrimStart('#');
    }

    private static void CopyShapeAppearance(PowerPointShape source, OdpShape target, PowerPointOpenDocumentConversionOptions options) {
        target.Hidden = source.Hidden;
        if (!options.IncludeBasicFormatting) return;
        if (!string.IsNullOrWhiteSpace(source.FillColor)) target.FillColor = ParseColor(source.FillColor);
        if (!string.IsNullOrWhiteSpace(source.OutlineColor)) target.StrokeColor = ParseColor(source.OutlineColor);
        if (source.OutlineWidthPoints.HasValue) target.StrokeWidth = OdfLength.Points(source.OutlineWidthPoints.Value);
    }

    private static void CopyShapeAppearance(OdpShape source, PowerPointShape target, PowerPointOpenDocumentConversionOptions options) {
        target.Hidden = source.Hidden;
        if (!options.IncludeBasicFormatting) return;
        if (source.FillColor.HasValue) target.FillColor = source.FillColor.Value.ToString().TrimStart('#');
        if (source.StrokeColor.HasValue) target.OutlineColor = source.StrokeColor.Value.ToString().TrimStart('#');
        if (source.StrokeWidth.HasValue) target.OutlineWidthPoints = source.StrokeWidth.Value.ToPoints();
    }

    private static void MapBackground(PowerPointSlide source, OdpSlide target, ref int converted, ref int unsupported) {
        PowerPointSlideBackground background = source.GetBackground();
        if (background.Kind == PowerPointSlideBackgroundKind.SolidColor && !string.IsNullOrWhiteSpace(background.Color)) {
            target.BackgroundColor = ParseColor(background.Color);
            converted++;
        } else if (background.Kind != PowerPointSlideBackgroundKind.None) unsupported++;
    }

    private static bool MapTransition(SlideTransition transition, OdpSlide target) {
        if (transition == SlideTransition.None) return false;
        target.TransitionType = "automatic";
        switch (transition) {
            case SlideTransition.Fade: target.TransitionStyle = "fade"; break;
            case SlideTransition.Wipe: target.TransitionStyle = "wipe"; break;
            case SlideTransition.Cut: target.TransitionStyle = "none"; break;
            default: target.TransitionStyle = transition.ToString().ToLowerInvariant(); break;
        }
        return true;
    }

    private static bool MapTransition(OdpSlide source, PowerPointSlide target) {
        string value = (source.TransitionStyle ?? source.TransitionType ?? string.Empty).ToLowerInvariant();
        if (value.Length == 0) return false;
        if (value.Contains("fade")) target.Transition = SlideTransition.Fade;
        else if (value.Contains("wipe")) target.Transition = SlideTransition.Wipe;
        else if (value.Contains("cut") || value == "none") target.Transition = SlideTransition.Cut;
        else return false;
        return true;
    }

    private static OdfRect ToOdfRect(PowerPointShape shape) => new OdfRect(
        OdfLength.Points(shape.LeftPoints), OdfLength.Points(shape.TopPoints),
        OdfLength.Points(Math.Max(0.01D, shape.WidthPoints)), OdfLength.Points(Math.Max(0.01D, shape.HeightPoints)));

    private static PowerPointLayoutBox ToPowerPointBox(OdfRect bounds) => PowerPointLayoutBox.FromPoints(
        bounds.X.ToPoints(), bounds.Y.ToPoints(), Math.Max(0.01D, bounds.Width.ToPoints()), Math.Max(0.01D, bounds.Height.ToPoints()));

    private static void ApplyOdpCrop(OdpImage source, PowerPointPicture target) {
        if (!source.Crop.HasValue) return;
        OdfInsets crop = source.Crop.Value;
        double width = Math.Max(0.01D, source.Bounds.Width.ToPoints());
        double height = Math.Max(0.01D, source.Bounds.Height.ToPoints());
        target.Crop(ClampPercent(crop.Left.ToPoints() / width * 100D), ClampPercent(crop.Top.ToPoints() / height * 100D),
            ClampPercent(crop.Right.ToPoints() / width * 100D), ClampPercent(crop.Bottom.ToPoints() / height * 100D));
    }

    private static double ClampPercent(double value) => Math.Max(0D, Math.Min(100D, value));

    private static OdfColor? ParseColor(string? value) {
        if (string.IsNullOrWhiteSpace(value)) return null;
        string hex = value!.Trim().TrimStart('#');
        if (hex.Length == 8) hex = hex.Substring(0, 6);
        return hex.Length == 6 ? OdfColor.Parse(hex) : (OdfColor?)null;
    }

    private static string FileNameForContentType(string? contentType) {
        switch ((contentType ?? string.Empty).ToLowerInvariant()) {
            case "image/jpeg": return "image.jpg";
            case "image/gif": return "image.gif";
            case "image/bmp": return "image.bmp";
            case "image/tiff": return "image.tiff";
            case "image/svg+xml": return "image.svg";
            case "image/x-emf": return "image.emf";
            case "image/x-wmf": return "image.wmf";
            default: return "image.png";
        }
    }

    private static bool TryGetImagePartType(string path, byte[] bytes, out ImagePartType type) {
        string normalizedPath = path;
        int suffix = normalizedPath.IndexOfAny(new[] { '?', '#' });
        if (suffix >= 0) normalizedPath = normalizedPath.Substring(0, suffix);
        try { normalizedPath = Uri.UnescapeDataString(normalizedPath); } catch (UriFormatException) { }
        switch (System.IO.Path.GetExtension(normalizedPath).ToLowerInvariant()) {
            case ".png": type = ImagePartType.Png; return true;
            case ".jpg":
            case ".jpeg": type = ImagePartType.Jpeg; return true;
            case ".gif": type = ImagePartType.Gif; return true;
            case ".bmp": type = ImagePartType.Bmp; return true;
            case ".tif":
            case ".tiff": type = ImagePartType.Tiff; return true;
        }
        if (bytes.Length >= 8 && bytes[0] == 0x89 && bytes[1] == 0x50 && bytes[2] == 0x4E && bytes[3] == 0x47) {
            type = ImagePartType.Png; return true;
        }
        if (bytes.Length >= 3 && bytes[0] == 0xFF && bytes[1] == 0xD8 && bytes[2] == 0xFF) {
            type = ImagePartType.Jpeg; return true;
        }
        if (bytes.Length >= 6 && bytes[0] == (byte)'G' && bytes[1] == (byte)'I' && bytes[2] == (byte)'F') {
            type = ImagePartType.Gif; return true;
        }
        if (bytes.Length >= 2 && bytes[0] == (byte)'B' && bytes[1] == (byte)'M') {
            type = ImagePartType.Bmp; return true;
        }
        if (bytes.Length >= 4 && ((bytes[0] == (byte)'I' && bytes[1] == (byte)'I' && bytes[2] == 42 && bytes[3] == 0) ||
                                (bytes[0] == (byte)'M' && bytes[1] == (byte)'M' && bytes[2] == 0 && bytes[3] == 42))) {
            type = ImagePartType.Tiff; return true;
        }
        type = ImagePartType.Png; return false;
    }

    private static IEnumerable<string> SplitParagraphs(string text) => text.Replace("\r\n", "\n")
        .Split(new[] { "\n\n" }, StringSplitOptions.None);

    private static void AddAdvancedPowerPointFindings(PowerPointFeatureReport source, OdfConversionReport target) {
        foreach (PowerPointFeatureFinding finding in source.PreservedFeatures.Concat(source.UnsupportedFeatures).Where(item => item.Count > 0)) {
            target.Add("source-" + Slug(finding.Name), OdfConversionMappingStatus.Unsupported, finding.Count, finding.Note);
        }
    }

    private static string Slug(string value) => new string(value.ToLowerInvariant().Select(character =>
        char.IsLetterOrDigit(character) ? character : '-').ToArray()).Trim('-');

    private static void AddConverted(OdfConversionReport report, string feature, int count) {
        if (count > 0) report.Add(feature, OdfConversionMappingStatus.Converted, count);
    }

    private static void AddUnsupported(OdfConversionReport report, string feature, int count, string? message) {
        if (count > 0) report.Add(feature, OdfConversionMappingStatus.Unsupported, count, message);
    }
}

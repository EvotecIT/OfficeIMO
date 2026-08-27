using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.OpenDocument;
using OfficeIMO.Drawing;
using OfficeIMO.PowerPoint;

namespace OfficeIMO.PowerPoint.OpenDocument;

/// <summary>Explicit conversions between OfficeIMO PowerPoint and native OpenDocument presentation models.</summary>
public static partial class PowerPointOpenDocumentConversionExtensions {
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

        int textBoxes = 0, pictures = 0, tables = 0, autoShapes = 0;
        int notes = 0, transitions = 0, backgrounds = 0, unsupportedBackgrounds = 0, unsupportedShapes = 0, unsupportedPictures = 0;
        int transformedShapes = 0, skippedBasicFormatting = 0, skippedNotes = 0;
        int unsupportedShapeHyperlinks = 0;
        var textState = new PowerPointToOdpTextConversionState();
        var imageValidationBudget = new OdfImageValidationBudget();
        for (int slideIndex = 0; slideIndex < source.Slides.Count; slideIndex++) {
            PowerPointSlide sourceSlide = source.Slides[slideIndex];
            OdpSlide targetSlide = target.AddSlide("Slide" + (slideIndex + 1).ToString(System.Globalization.CultureInfo.InvariantCulture));
            targetSlide.Hidden = sourceSlide.Hidden;
            MapBackground(sourceSlide, targetSlide, ref backgrounds, ref unsupportedBackgrounds);
            if (MapTransition(sourceSlide.Transition, targetSlide)) transitions++;

            foreach (PowerPointShape shape in sourceSlide.Shapes.OrderBy(item => item.DrawingOrder)) {
                if (shape.Hyperlink != null) unsupportedShapeHyperlinks++;
                if (shape is PowerPointTextBox textBox) {
                    OdpTextBox converted = targetSlide.AddTextBox(ToOdfRect(textBox), null, textBox.Name);
                    CopyShapeAppearance(textBox, converted, effective);
                    CopyPowerPointParagraphsToOdp(textBox.Paragraphs,
                        () => converted.AddParagraph(), effective, textState);
                    textBoxes++;
                } else if (shape is PowerPointPicture picture) {
                    if (!effective.IncludeImages) { unsupportedPictures++; continue; }
                    try {
                        byte[] imageBytes = picture.GetImageBytes();
                        string imageFileName = FileNameForContentType(picture.ContentType);
                        if (!OdfImagePayloadValidator.TryResolvePreservedFileName(
                            imageBytes,
                            imageFileName,
                            out string storedFileName,
                            imageValidationBudget)) {
                            unsupportedPictures++;
                            continue;
                        }
                        OdpImage converted = targetSlide.AddImage(imageBytes, storedFileName, ToOdfRect(picture), picture.Name);
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
                            CopyPowerPointTableCellToOdp(cell, converted.Cell(row, column),
                                effective, textState);
                            if (cell.IsMergeAnchor) merges.Add((row, column, cell.Merge.rows, cell.Merge.columns));
                        }
                    }
                    foreach (var merge in merges) converted.Merge(merge.Row, merge.Column, merge.RowSpan, merge.ColumnSpan);
                    tables++;
                } else if (shape is PowerPointAutoShape autoShape) {
                    OdpShape converted;
                    if (autoShape.ShapeType == OfficePresetShapeType.Ellipse) converted = targetSlide.AddEllipse(ToOdfRect(autoShape), autoShape.Name);
                    else if (autoShape.ShapeType == OfficePresetShapeType.Line) {
                        converted = targetSlide.AddLine(OdfLength.Points(autoShape.LeftPoints), OdfLength.Points(autoShape.TopPoints),
                            OdfLength.Points(autoShape.RightPoints), OdfLength.Points(autoShape.BottomPoints), autoShape.Name);
                    } else {
                        converted = targetSlide.AddRectangle(ToOdfRect(autoShape), autoShape.Name);
                        if (autoShape.ShapeType != OfficePresetShapeType.Rectangle) transformedShapes++;
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
                if (!effective.IncludeBasicFormatting && HasBasicFormatting(shape)) skippedBasicFormatting++;
                if (shape.Rotation.HasValue || shape.HorizontalFlip.HasValue || shape.VerticalFlip.HasValue) transformedShapes++;
            }

            if (effective.IncludeSpeakerNotes && sourceSlide.HasSpeakerNotes) {
                IReadOnlyList<PowerPointParagraph> noteParagraphs = sourceSlide.Notes.Paragraphs;
                if (noteParagraphs.Any(paragraph => paragraph.Text.Length > 0
                    || paragraph.InlineNodes.Any(node =>
                        node.Kind != PowerPointParagraphInlineKind.Run || node.Text.Length > 0))) {
                    OdpNotes convertedNotes = targetSlide.GetOrCreateSpeakerNotes();
                    CopyPowerPointParagraphsToOdp(noteParagraphs,
                        () => convertedNotes.AddParagraph(), effective, textState);
                    notes++;
                }
            } else if (!effective.IncludeSpeakerNotes && sourceSlide.HasSpeakerNotes) {
                skippedNotes++;
            }
        }

        AddConverted(report, "slides", source.Slides.Count);
        AddConverted(report, "text-boxes", textBoxes);
        AddConverted(report, "paragraphs", textState.Paragraphs);
        AddConverted(report, "text-runs", textState.TextRuns);
        AddConverted(report, "line-breaks", textState.LineBreaks);
        AddConverted(report, "images", pictures);
        AddConverted(report, "tables", tables);
        AddConverted(report, "basic-shapes", autoShapes);
        AddConverted(report, "speaker-notes", notes);
        AddConverted(report, "solid-backgrounds", backgrounds);
        AddUnsupported(report, "slide-backgrounds", unsupportedBackgrounds, "Image, gradient, theme, and unsupported backgrounds are not translated.");
        if (transitions > 0) report.Add("slide-transitions", OdfConversionMappingStatus.Approximated, transitions,
            "Common transition families are mapped without PowerPoint-specific speed and timing metadata.");
        if (textState.ListParagraphs > 0) report.Add("text-lists", OdfConversionMappingStatus.Approximated, textState.ListParagraphs,
            "List text is retained as paragraphs; PowerPoint bullet and numbering definitions are not translated.");
        if (textState.Fields > 0) report.Add("paragraph-fields", OdfConversionMappingStatus.Approximated,
            textState.Fields, "PowerPoint dynamic fields retain their displayed text but not their update semantics.");
        if (textState.ApproximatedAlignments > 0) report.Add("paragraph-alignments",
            OdfConversionMappingStatus.Approximated, textState.ApproximatedAlignments,
            "PowerPoint distributed and low-justification variants are represented as ODF justification.");
        if (textState.ApproximatedTextDecorations > 0) report.Add("text-decorations",
            OdfConversionMappingStatus.Approximated, textState.ApproximatedTextDecorations,
            "PowerPoint words-only and heavy underline variants are represented by the nearest ODF line pattern without their word or weight semantics.");
        int totalSkippedBasicFormatting = skippedBasicFormatting + textState.SkippedBasicFormatting;
        if (totalSkippedBasicFormatting > 0) report.Add("basic-formatting", OdfConversionMappingStatus.Skipped, totalSkippedBasicFormatting,
            "Common text, fill, and outline formatting was omitted because IncludeBasicFormatting is disabled.");
        if (skippedNotes > 0) report.Add("speaker-notes", OdfConversionMappingStatus.Skipped, skippedNotes,
            "Speaker notes were omitted because IncludeSpeakerNotes is disabled.");
        AddUnsupported(report, "shape-transforms", transformedShapes, "Complex geometry, rotation, flips, and connector semantics are approximated or omitted.");
        AddUnsupported(report, "hyperlink-tooltips", textState.UnsupportedHyperlinkTooltips,
            "PowerPoint hyperlink tooltips have no equivalent in the current ODP hyperlink surface and were omitted.");
        AddUnsupported(report, "shape-hyperlinks", unsupportedShapeHyperlinks,
            "PowerPoint shape-level click hyperlinks, including internal slide jumps, are not translated to ODP.");
        AddUnsupported(report, "run-interactions", textState.UnsupportedRunInteractions,
            "PowerPoint run actions, mouse-over interactions, and action sounds outside ordinary click hyperlinks are not represented in ODP.");
        AddUnsupported(report, "images", unsupportedPictures, "Images disabled by options or unavailable from an embedded image part were skipped.");
        AddUnsupported(report, "shapes", unsupportedShapes, "Charts, SmartArt, media, groups, and other advanced drawing shapes are not translated.");
        report.Add("masters-layouts", OdfConversionMappingStatus.Approximated, source.Slides.Count,
            "Slide content is placed on one default ODP master and blank layout.");
        AddAdvancedPowerPointFindings(source.InspectFeatures(), report);
        return new OdfConversionResult<OdpPresentation>(target, report).ApplyPolicy(effective.LossPolicy);
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
        int unsupportedMeasurements = 0;
        if (source.PageWidth.TryToPoints(out double pageWidth) && source.PageHeight.TryToPoints(out double pageHeight)) {
            target.SlideSize.SetSizePoints(pageWidth, pageHeight);
        } else {
            unsupportedMeasurements++;
        }

        int textBoxes = 0, paragraphs = 0, textRuns = 0, hyperlinks = 0, externalHyperlinks = 0, pictures = 0, tables = 0, basicShapes = 0;
        int notes = 0, transitions = 0, unsupportedTransitions = 0, unsupportedShapes = 0, unsupportedPictures = 0, transformedShapes = 0;
        int listParagraphs = 0, approximatedRuns = 0, unsupportedHyperlinks = 0, unsupportedHyperlinkBehaviors = 0;
        int skippedBasicFormatting = 0, skippedNotes = 0, noteContainers = 0;
        int approximatedTextDecorations = CountNonSolidTextDecorations(source);
        int unsupportedWritingModes = 0, approximatedParagraphAlignments = 0;
        int approximatedFontFamilyLists = 0, unsupportedFontFamilies = 0;
        var pendingInternalLinks = new List<(PowerPointTextRun Run, int SlideIndex)>();
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
                    if (!TryToPowerPointBox(textBox.Bounds, out PowerPointLayoutBox textBoxBounds)) {
                        unsupportedMeasurements++;
                        unsupportedShapes++;
                        continue;
                    }
                    IReadOnlyList<OdpParagraph> sourceParagraphs = textBox.Paragraphs;
                    listParagraphs += textBox.Lists.Sum(list => list.Items.Count);
                    PowerPointTextBox converted = targetSlide.AddTextBox(string.Empty, textBoxBounds);
                    converted.Name = textBox.Name;
                    unsupportedMeasurements += CopyShapeAppearance(textBox, converted, effective);
                    CopyOdpParagraphsToPowerPoint(sourceParagraphs,
                        paragraphTexts => converted.SetParagraphs(paragraphTexts), source.Slides,
                        pendingInternalLinks, effective, ref paragraphs, ref textRuns, ref hyperlinks,
                        ref externalHyperlinks, ref unsupportedHyperlinks, ref unsupportedHyperlinkBehaviors, ref approximatedRuns,
                        ref skippedBasicFormatting, ref unsupportedWritingModes,
                        ref approximatedParagraphAlignments, ref unsupportedMeasurements,
                        ref approximatedFontFamilyLists, ref unsupportedFontFamilies);
                    textBoxes++;
                } else if (shape is OdpImage image) {
                    if (!effective.IncludeImages) {
                        unsupportedPictures++;
                        continue;
                    }
                    try {
                        byte[] imageBytes = image.GetImageBytes();
                        if (!TryGetImagePartType(image.Path, imageBytes, out OfficeImageFormat imageType)) {
                            unsupportedPictures++;
                            continue;
                        }
                        if (!TryToPowerPointBox(image.Bounds, out PowerPointLayoutBox imageBounds)) {
                            unsupportedMeasurements++;
                            unsupportedPictures++;
                            continue;
                        }
                        using var stream = new MemoryStream(imageBytes, writable: false);
                        PowerPointPicture converted = targetSlide.AddPicture(stream, imageType, imageBounds);
                        converted.Name = image.Name;
                        unsupportedMeasurements += CopyShapeAppearance(image, converted, effective);
                        unsupportedMeasurements += ApplyOdpCrop(image, converted);
                        pictures++;
                    } catch (Exception exception) when (exception is NotSupportedException || exception is InvalidDataException ||
                        exception is ArgumentException) {
                        unsupportedPictures++;
                    }
                } else if (shape is OdpTable table) {
                    if (!TryToPowerPointBox(table.Bounds, out PowerPointLayoutBox tableBounds)) {
                        unsupportedMeasurements++;
                        unsupportedShapes++;
                        continue;
                    }
                    int rowCount = Math.Max(1, table.Rows.Count);
                    if (rowCount > effective.MaxTableRows) {
                        throw new InvalidDataException($"ODP table rows ({rowCount}) exceed the configured conversion limit ({effective.MaxTableRows}).");
                    }
                    int columnCount = Math.Max(1, table.Rows.Select(row => row.Cells.Count).DefaultIfEmpty(1).Max());
                    if (columnCount > effective.MaxTableColumns) {
                        throw new InvalidDataException($"ODP table columns ({columnCount}) exceed the configured conversion limit ({effective.MaxTableColumns}).");
                    }
                    PowerPointTable converted = targetSlide.AddTable(rowCount, columnCount, tableBounds);
                    converted.Name = table.Name;
                    unsupportedMeasurements += CopyShapeAppearance(table, converted, effective);
                    var merges = new List<(int Row, int Column, int RowSpan, int ColumnSpan)>();
                    for (int row = 0; row < table.Rows.Count; row++) {
                        IReadOnlyList<OdpTableCell> cells = table.Rows[row].Cells;
                        for (int column = 0; column < cells.Count; column++) {
                            OdpTableCell cell = cells[column];
                            if (cell.IsCovered) continue;
                            CopyOdpParagraphsToPowerPoint(cell.Paragraphs,
                                paragraphTexts => converted.GetCell(row, column).SetParagraphs(paragraphTexts), source.Slides,
                                pendingInternalLinks, effective, ref paragraphs, ref textRuns, ref hyperlinks,
                                ref externalHyperlinks, ref unsupportedHyperlinks, ref unsupportedHyperlinkBehaviors, ref approximatedRuns,
                                ref skippedBasicFormatting, ref unsupportedWritingModes,
                                ref approximatedParagraphAlignments, ref unsupportedMeasurements,
                                ref approximatedFontFamilyLists, ref unsupportedFontFamilies);
                            if (cell.RowSpan > 1 || cell.ColumnSpan > 1) merges.Add((row, column, cell.RowSpan, cell.ColumnSpan));
                        }
                    }
                    foreach (var merge in merges) converted.MergeCells(merge.Row, merge.Column,
                        merge.Row + merge.RowSpan - 1, merge.Column + merge.ColumnSpan - 1);
                    tables++;
                } else if (shape is OdpRectangle rectangle) {
                    if (!TryGetRectPoints(rectangle.Bounds, out double x, out double y, out double width, out double height)) {
                        unsupportedMeasurements++;
                        unsupportedShapes++;
                        continue;
                    }
                    PowerPointAutoShape converted = targetSlide.AddRectanglePoints(x, y, width, height, rectangle.Name);
                    unsupportedMeasurements += CopyShapeAppearance(rectangle, converted, effective);
                    basicShapes++;
                } else if (shape is OdpEllipse ellipse) {
                    if (!TryGetRectPoints(ellipse.Bounds, out double x, out double y, out double width, out double height)) {
                        unsupportedMeasurements++;
                        unsupportedShapes++;
                        continue;
                    }
                    PowerPointAutoShape converted = targetSlide.AddEllipsePoints(x, y, width, height, ellipse.Name);
                    unsupportedMeasurements += CopyShapeAppearance(ellipse, converted, effective);
                    basicShapes++;
                } else if (shape is OdpLine line) {
                    if (!line.X1.TryToPoints(out double x1) || !line.Y1.TryToPoints(out double y1)
                        || !line.X2.TryToPoints(out double x2) || !line.Y2.TryToPoints(out double y2)) {
                        unsupportedMeasurements++;
                        unsupportedShapes++;
                        continue;
                    }
                    PowerPointAutoShape converted = targetSlide.AddLinePoints(x1, y1, x2, y2, line.Name);
                    unsupportedMeasurements += CopyShapeAppearance(line, converted, effective);
                    basicShapes++;
                } else {
                    unsupportedShapes++;
                }
                if (!effective.IncludeBasicFormatting && HasBasicFormatting(shape)) skippedBasicFormatting++;
                if (!string.IsNullOrWhiteSpace(shape.Transform)) transformedShapes++;
            }

            if (sourceSlide.SpeakerNotes != null) noteContainers++;
            if (effective.IncludeSpeakerNotes && sourceSlide.SpeakerNotes != null) {
                IReadOnlyList<OdpParagraph> noteParagraphs = sourceSlide.SpeakerNotes.Paragraphs;
                if (noteParagraphs.Any(paragraph => paragraph.Text.Length > 0)) {
                    CopyOdpParagraphsToPowerPoint(noteParagraphs,
                        paragraphTexts => targetSlide.Notes.SetParagraphs(paragraphTexts), source.Slides,
                        pendingInternalLinks, effective, ref paragraphs, ref textRuns, ref hyperlinks,
                        ref externalHyperlinks, ref unsupportedHyperlinks, ref unsupportedHyperlinkBehaviors, ref approximatedRuns,
                        ref skippedBasicFormatting, ref unsupportedWritingModes,
                        ref approximatedParagraphAlignments, ref unsupportedMeasurements,
                        ref approximatedFontFamilyLists, ref unsupportedFontFamilies);
                    notes++;
                }
            } else if (!effective.IncludeSpeakerNotes && sourceSlide.SpeakerNotes != null &&
                       sourceSlide.SpeakerNotes.Paragraphs.Any(paragraph => paragraph.Text.Length > 0)) {
                skippedNotes++;
            }
        }

        foreach ((PowerPointTextRun run, int slideIndex) in pendingInternalLinks) {
            run.SetHyperlink(target.Slides[slideIndex]);
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
            "Inline ODP elements outside plain text, spans, and hyperlinks were flattened to text.");
        if (approximatedTextDecorations > 0) report.Add("text-decorations", OdfConversionMappingStatus.Approximated,
            approximatedTextDecorations, "Patterned ODF line-through and non-wave patterned double underline variants are simplified to PowerPoint's nearest native decoration.");
        if (approximatedFontFamilyLists > 0) report.Add("font-family-fallbacks", OdfConversionMappingStatus.Approximated,
            approximatedFontFamilyLists, "PowerPoint run properties retain the first ODF font family but cannot retain the authored fallback list.");
        AddUnsupported(report, "font-families", unsupportedFontFamilies,
            "Malformed ODF font-family syntax was omitted instead of being emitted as an invalid PowerPoint typeface name.");
        AddUnsupported(report, "writing-mode", unsupportedWritingModes,
            "Vertical and unsupported ODF writing modes cannot be represented by the PowerPoint paragraph model.");
        if (approximatedParagraphAlignments > 0) report.Add("paragraph-alignments",
            OdfConversionMappingStatus.Approximated, approximatedParagraphAlignments,
            "Logical ODF start/end alignment is projected to the matching physical PowerPoint edge.");
        AddUnsupported(report, "hyperlinks", unsupportedHyperlinks,
            "Hyperlink targets that could not be resolved as slides or valid URI references were omitted.");
        AddUnsupported(report, "hyperlink-target-behavior", unsupportedHyperlinkBehaviors,
            "ODF target-frame-name and XLink show behavior have no equivalent in PowerPoint run hyperlinks and were omitted.");
        if (skippedBasicFormatting > 0) report.Add("basic-formatting", OdfConversionMappingStatus.Skipped, skippedBasicFormatting,
            "Common text, fill, and outline formatting was omitted because IncludeBasicFormatting is disabled.");
        if (skippedNotes > 0) report.Add("speaker-notes", OdfConversionMappingStatus.Skipped, skippedNotes,
            "Speaker notes were omitted because IncludeSpeakerNotes is disabled.");
        AddUnsupported(report, "slide-transitions", unsupportedTransitions, "The ODF transition family is not supported by the PowerPoint adapter.");
        AddUnsupported(report, "images", unsupportedPictures, "Images disabled by options or using an unsupported PowerPoint image format were skipped.");
        AddUnsupported(report, "shapes", unsupportedShapes, "Groups and unsupported ODF drawing elements are not translated.");
        AddUnsupported(report, "shape-transforms", transformedShapes, "Raw ODF transform expressions are not translated.");
        AddUnsupported(report, "relative-measurements", unsupportedMeasurements,
            "Relative or unsupported ODF text measurements could not be projected to fixed PowerPoint point sizes and were omitted.");
        if (source.MasterPages.Count > 0 || source.Layouts.Count > 0) report.Add("masters-layouts", OdfConversionMappingStatus.Approximated,
            source.MasterPages.Count + source.Layouts.Count, "Content is placed on PowerPoint's default master and layout.");
        AddUnmappedOdfFindings(source.InspectFeatures(), report, externalHyperlinks, noteContainers,
            source.MasterPages.Count, transitions + unsupportedTransitions);
        return new OdfConversionResult<PowerPointPresentation>(target, report).ApplyPolicy(effective.LossPolicy);
    }

    private static bool TryResolveSlideFragment(
        string href,
        IReadOnlyList<OdpSlide> slides,
        out int zeroBasedIndex) {
        zeroBasedIndex = -1;
        if (!OdfUriReference.TryDecodeFragment(href, out string fragment)) return false;
        const string prefix = "slide-";
        if (fragment.StartsWith(prefix, StringComparison.OrdinalIgnoreCase)
            && int.TryParse(fragment.Substring(prefix.Length), System.Globalization.NumberStyles.None,
                System.Globalization.CultureInfo.InvariantCulture, out int oneBased)
            && oneBased >= 1 && oneBased <= slides.Count) {
            zeroBasedIndex = oneBased - 1;
            return true;
        }
        for (int index = 0; index < slides.Count; index++) {
            if (!string.Equals(slides[index].Name, fragment, StringComparison.Ordinal)) continue;
            zeroBasedIndex = index;
            return true;
        }
        return false;
    }

    private static bool IsExternalOdfHref(string href) =>
        !string.IsNullOrWhiteSpace(href) && !href.StartsWith("#", StringComparison.Ordinal)
        && (href.StartsWith("//", StringComparison.Ordinal) || Uri.TryCreate(href, UriKind.Absolute, out _));

    private static void AddUnmappedOdfFindings(
        OdfFeatureReport features,
        OdfConversionReport report,
        int hyperlinks,
        int notes,
        int masterPages,
        int transitions) {
        foreach (OdfFeatureDiagnostic diagnostic in features.Diagnostics) {
            report.Add("source-inspection", OdfConversionMappingStatus.Unsupported, 1,
                diagnostic.Code + " in " + diagnostic.PartPath + ": " + diagnostic.Message);
        }
        int remainingHyperlinks = hyperlinks, remainingNotes = notes;
        int remainingMasterPages = masterPages, remainingTransitions = transitions;
        foreach (OdfFeatureFinding finding in features.Findings) {
            int handled = 0;
            if (finding.Name == "external-links") handled = Consume(ref remainingHyperlinks, finding.Count);
            else if (finding.Name == "presentation-notes") handled = Consume(ref remainingNotes, finding.Count);
            else if (finding.Name == "master-pages") handled = Consume(ref remainingMasterPages, finding.Count);
            else if (finding.Name == "presentation-transitions") handled = Consume(ref remainingTransitions, finding.Count);
            int remaining = Math.Max(0, finding.Count - handled);
            if (remaining > 0) report.Add("source-" + finding.Name, OdfConversionMappingStatus.Unsupported, remaining,
                "The source ODP feature cannot be transferred to PPTX by this adapter.");
        }
    }

    private static int Consume(ref int available, int requested) {
        int consumed = Math.Min(available, requested);
        available -= consumed;
        return consumed;
    }

    private static PowerPointOpenDocumentConversionOptions NormalizeOptions(PowerPointOpenDocumentConversionOptions? options) {
        PowerPointOpenDocumentConversionOptions effective = options ?? new PowerPointOpenDocumentConversionOptions();
        if (effective.MaxTableRows <= 0) throw new ArgumentOutOfRangeException(nameof(options), effective.MaxTableRows,
            $"{nameof(PowerPointOpenDocumentConversionOptions.MaxTableRows)} must be positive.");
        if (effective.MaxTableColumns <= 0) throw new ArgumentOutOfRangeException(nameof(options), effective.MaxTableColumns,
            $"{nameof(PowerPointOpenDocumentConversionOptions.MaxTableColumns)} must be positive.");
        return effective;
    }

    private static void ApplyPowerPointRun(PowerPointTextRun source, OdpRun target,
        PowerPointOpenDocumentConversionOptions options) {
        if (!options.IncludeBasicFormatting) return;
        target.Bold = source.Bold ? true : (bool?)null;
        target.Italic = source.Italic ? true : (bool?)null;
        ApplyPowerPointTextSemantics(source, target);
        if (source.FontSizePoints.HasValue) target.FontSize = OdfLength.Points(source.FontSizePoints.Value);
        target.FontFamily = source.FontName;
        if (!string.IsNullOrWhiteSpace(source.Color)) target.Color = ParseColor(source.Color);
        if (!string.IsNullOrWhiteSpace(source.HighlightColor)) target.BackgroundColor = ParseColor(source.HighlightColor);
    }

    private static void ApplyPowerPointRun(PowerPointTextRun source, OdpHyperlink target,
        PowerPointOpenDocumentConversionOptions options) {
        if (!options.IncludeBasicFormatting) return;
        target.Bold = source.Bold ? true : (bool?)null;
        target.Italic = source.Italic ? true : (bool?)null;
        ApplyPowerPointTextSemantics(source, target);
        if (source.FontSizePoints.HasValue) target.FontSize = OdfLength.Points(source.FontSizePoints.Value);
        target.FontFamily = source.FontName;
        if (!string.IsNullOrWhiteSpace(source.Color)) target.Color = ParseColor(source.Color);
        if (!string.IsNullOrWhiteSpace(source.HighlightColor)) target.BackgroundColor = ParseColor(source.HighlightColor);
    }

    private static bool HasBasicFormatting(PowerPointTextRun run) =>
        run.Bold || run.Italic || run.Underline || run.Strikethrough || run.FontSizePoints.HasValue ||
        run.UnderlineStyle.HasValue || run.StrikeStyle.HasValue || run.BaselinePercent.HasValue || run.Capitalization.HasValue ||
        !string.IsNullOrWhiteSpace(run.FontName) || !string.IsNullOrWhiteSpace(run.Color) ||
        !string.IsNullOrWhiteSpace(run.HighlightColor);

    private static bool HasBasicFormatting(OdpRun run) =>
        run.Bold.HasValue || run.Italic.HasValue || run.Underline.HasValue || run.StrikeThrough.HasValue ||
        run.UnderlineStyle.HasValue || run.UnderlineType.HasValue || run.LineThroughStyle.HasValue ||
        run.LineThroughType.HasValue || run.TextPosition.HasValue || run.TextTransform.HasValue || run.SmallCaps.HasValue ||
        run.FontSize.HasValue || !string.IsNullOrWhiteSpace(run.FontFamily) || run.Color.HasValue || run.BackgroundColor.HasValue;

    private static bool HasBasicFormatting(OdpParagraph paragraph) =>
        paragraph.Bold.HasValue || paragraph.Italic.HasValue || paragraph.Underline.HasValue || paragraph.StrikeThrough.HasValue ||
        paragraph.UnderlineStyle.HasValue || paragraph.UnderlineType.HasValue || paragraph.LineThroughStyle.HasValue ||
        paragraph.LineThroughType.HasValue || paragraph.TextPosition.HasValue || paragraph.TextTransform.HasValue || paragraph.SmallCaps.HasValue ||
        paragraph.FontSize.HasValue || !string.IsNullOrWhiteSpace(paragraph.FontFamily) || paragraph.Color.HasValue || paragraph.BackgroundColor.HasValue;

    private static bool HasBasicFormatting(OdpHyperlink run) =>
        run.Bold.HasValue || run.Italic.HasValue || run.Underline.HasValue || run.StrikeThrough.HasValue ||
        run.UnderlineStyle.HasValue || run.UnderlineType.HasValue || run.LineThroughStyle.HasValue ||
        run.LineThroughType.HasValue || run.TextPosition.HasValue || run.TextTransform.HasValue || run.SmallCaps.HasValue ||
        run.FontSize.HasValue || !string.IsNullOrWhiteSpace(run.FontFamily) || run.Color.HasValue || run.BackgroundColor.HasValue;

    private static void ApplyPowerPointTextSemantics(PowerPointTextRun source, OdpRun target) {
        (target.UnderlineStyle, target.UnderlineType) = MapPowerPointUnderline(source.UnderlineStyle);
        target.LineThroughStyle = source.Strikethrough ? OdfTextDecorationStyle.Solid : OdfTextDecorationStyle.None;
        target.LineThroughType = source.StrikeStyle == PowerPointStrikeStyle.Double
            ? OdfTextDecorationType.Double
            : source.Strikethrough ? OdfTextDecorationType.Single : OdfTextDecorationType.None;
        target.TextPosition = MapPowerPointTextPosition(source.BaselinePercent);
        target.TextTransform = source.Capitalization == PowerPointCapitalization.AllCaps
            ? OdfTextTransform.Uppercase
            : OdfTextTransform.None;
        target.SmallCaps = source.Capitalization == PowerPointCapitalization.SmallCaps ? true : (bool?)null;
    }

    private static void ApplyPowerPointTextSemantics(PowerPointTextRun source, OdpHyperlink target) {
        (target.UnderlineStyle, target.UnderlineType) = MapPowerPointUnderline(source.UnderlineStyle);
        target.LineThroughStyle = source.Strikethrough ? OdfTextDecorationStyle.Solid : OdfTextDecorationStyle.None;
        target.LineThroughType = source.StrikeStyle == PowerPointStrikeStyle.Double
            ? OdfTextDecorationType.Double
            : source.Strikethrough ? OdfTextDecorationType.Single : OdfTextDecorationType.None;
        target.TextPosition = MapPowerPointTextPosition(source.BaselinePercent);
        target.TextTransform = source.Capitalization == PowerPointCapitalization.AllCaps
            ? OdfTextTransform.Uppercase
            : OdfTextTransform.None;
        target.SmallCaps = source.Capitalization == PowerPointCapitalization.SmallCaps ? true : (bool?)null;
    }

    private static (OdfTextDecorationStyle? Style, OdfTextDecorationType? Type) MapPowerPointUnderline(
        PowerPointUnderlineStyle? style) => style switch {
        null or PowerPointUnderlineStyle.None => (OdfTextDecorationStyle.None, OdfTextDecorationType.None),
        PowerPointUnderlineStyle.Double => (OdfTextDecorationStyle.Solid, OdfTextDecorationType.Double),
        PowerPointUnderlineStyle.WavyDouble => (OdfTextDecorationStyle.Wave, OdfTextDecorationType.Double),
        PowerPointUnderlineStyle.Dotted or PowerPointUnderlineStyle.HeavyDotted => (OdfTextDecorationStyle.Dotted, OdfTextDecorationType.Single),
        PowerPointUnderlineStyle.Dash or PowerPointUnderlineStyle.DashHeavy => (OdfTextDecorationStyle.Dash, OdfTextDecorationType.Single),
        PowerPointUnderlineStyle.DashLong or PowerPointUnderlineStyle.DashLongHeavy => (OdfTextDecorationStyle.LongDash, OdfTextDecorationType.Single),
        PowerPointUnderlineStyle.DotDash or PowerPointUnderlineStyle.DotDashHeavy => (OdfTextDecorationStyle.DotDash, OdfTextDecorationType.Single),
        PowerPointUnderlineStyle.DotDotDash or PowerPointUnderlineStyle.DotDotDashHeavy => (OdfTextDecorationStyle.DotDotDash, OdfTextDecorationType.Single),
        PowerPointUnderlineStyle.Wavy or PowerPointUnderlineStyle.WavyHeavy => (OdfTextDecorationStyle.Wave, OdfTextDecorationType.Single),
        _ => (OdfTextDecorationStyle.Solid, OdfTextDecorationType.Single)
    };

    private static OdfTextPosition? MapPowerPointTextPosition(double? value) => value switch {
        > 0D => OdfTextPosition.Superscript,
        < 0D => OdfTextPosition.Subscript,
        0D => OdfTextPosition.Normal,
        _ => null
    };

    private static bool HasBasicFormatting(PowerPointShape shape) =>
        !string.IsNullOrWhiteSpace(shape.FillColor) || !string.IsNullOrWhiteSpace(shape.OutlineColor) || shape.OutlineWidthPoints.HasValue;

    private static bool HasBasicFormatting(OdpShape shape) =>
        shape.FillColor.HasValue || shape.StrokeColor.HasValue || shape.StrokeWidth.HasValue;

    private static void CopyShapeAppearance(PowerPointShape source, OdpShape target, PowerPointOpenDocumentConversionOptions options) {
        target.Hidden = source.Hidden;
        if (!options.IncludeBasicFormatting) return;
        if (!string.IsNullOrWhiteSpace(source.FillColor)) target.FillColor = ParseColor(source.FillColor);
        if (!string.IsNullOrWhiteSpace(source.OutlineColor)) target.StrokeColor = ParseColor(source.OutlineColor);
        if (source.OutlineWidthPoints.HasValue) target.StrokeWidth = OdfLength.Points(source.OutlineWidthPoints.Value);
    }

    private static int CopyShapeAppearance(OdpShape source, PowerPointShape target, PowerPointOpenDocumentConversionOptions options) {
        target.Hidden = source.Hidden;
        if (!options.IncludeBasicFormatting) return 0;
        if (source.FillColor.HasValue) target.FillColor = source.FillColor.Value.ToString().TrimStart('#');
        if (source.StrokeColor.HasValue) target.OutlineColor = source.StrokeColor.Value.ToString().TrimStart('#');
        if (source.StrokeWidth.HasValue) {
            if (source.StrokeWidth.Value.TryToPoints(out double points)) target.OutlineWidthPoints = points;
            else return 1;
        }
        return 0;
    }

    private static void MapBackground(PowerPointSlide source, OdpSlide target, ref int converted, ref int unsupported) {
        PowerPointSlideBackground background = source.GetBackground();
        if (background.Kind == PowerPointSlideBackgroundKind.SolidColor && !string.IsNullOrWhiteSpace(background.Color)) {
            target.BackgroundColor = ParseColor(background.Color);
            converted++;
        } else if (background.Kind != PowerPointSlideBackgroundKind.None) unsupported++;
    }

    private static bool MapTransition(PowerPointSlideTransition transition, OdpSlide target) {
        if (transition == PowerPointSlideTransition.None) return false;
        target.TransitionType = "automatic";
        switch (transition) {
            case PowerPointSlideTransition.Fade: target.TransitionStyle = "fade"; break;
            case PowerPointSlideTransition.Wipe: target.TransitionStyle = "wipe"; break;
            case PowerPointSlideTransition.Cut: target.TransitionStyle = "none"; break;
            default: target.TransitionStyle = transition.ToString().ToLowerInvariant(); break;
        }
        return true;
    }

    private static bool MapTransition(OdpSlide source, PowerPointSlide target) {
        string value = (source.TransitionStyle ?? source.TransitionType ?? string.Empty).ToLowerInvariant();
        if (value.Length == 0) return false;
        if (value.Contains("fade")) target.Transition = PowerPointSlideTransition.Fade;
        else if (value.Contains("wipe")) target.Transition = PowerPointSlideTransition.Wipe;
        else if (value.Contains("cut") || value == "none") target.Transition = PowerPointSlideTransition.Cut;
        else return false;
        return true;
    }

    private static OdfRect ToOdfRect(PowerPointShape shape) => new OdfRect(
        OdfLength.Points(shape.LeftPoints), OdfLength.Points(shape.TopPoints),
        OdfLength.Points(Math.Max(0.01D, shape.WidthPoints)), OdfLength.Points(Math.Max(0.01D, shape.HeightPoints)));

    private static bool TryToPowerPointBox(OdfRect bounds, out PowerPointLayoutBox box) {
        if (!TryGetRectPoints(bounds, out double x, out double y, out double width, out double height)) {
            box = default;
            return false;
        }
        box = PowerPointLayoutBox.FromPoints(x, y, width, height);
        return true;
    }

    private static bool TryGetRectPoints(OdfRect bounds, out double x, out double y, out double width, out double height) {
        bool xValid = bounds.X.TryToPoints(out x);
        bool yValid = bounds.Y.TryToPoints(out y);
        bool widthValid = bounds.Width.TryToPoints(out width);
        bool heightValid = bounds.Height.TryToPoints(out height);
        bool valid = xValid && yValid && widthValid && heightValid;
        if (!valid) return false;
        width = Math.Max(0.01D, width);
        height = Math.Max(0.01D, height);
        return true;
    }

    private static int ApplyOdpCrop(OdpImage source, PowerPointPicture target) {
        if (!source.Crop.HasValue) return 0;
        OdfInsets crop = source.Crop.Value;
        if (!source.Bounds.Width.TryToPoints(out double width) || !source.Bounds.Height.TryToPoints(out double height)
            || !crop.Left.TryToPoints(out double left) || !crop.Top.TryToPoints(out double top)
            || !crop.Right.TryToPoints(out double right) || !crop.Bottom.TryToPoints(out double bottom)) return 1;
        width = Math.Max(0.01D, width);
        height = Math.Max(0.01D, height);
        target.Crop(ClampPercent(left / width * 100D), ClampPercent(top / height * 100D),
            ClampPercent(right / width * 100D), ClampPercent(bottom / height * 100D));
        return 0;
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

    private static bool TryGetImagePartType(string path, byte[] bytes, out OfficeImageFormat type) {
        string normalizedPath = path;
        int suffix = normalizedPath.IndexOfAny(new[] { '?', '#' });
        if (suffix >= 0) normalizedPath = normalizedPath.Substring(0, suffix);
        try { normalizedPath = Uri.UnescapeDataString(normalizedPath); } catch (UriFormatException) { }
        switch (System.IO.Path.GetExtension(normalizedPath).ToLowerInvariant()) {
            case ".png": type = OfficeImageFormat.Png; return true;
            case ".jpg":
            case ".jpeg": type = OfficeImageFormat.Jpeg; return true;
            case ".gif": type = OfficeImageFormat.Gif; return true;
            case ".bmp": type = OfficeImageFormat.Bmp; return true;
            case ".tif":
            case ".tiff": type = OfficeImageFormat.Tiff; return true;
        }
        if (bytes.Length >= 8 && bytes[0] == 0x89 && bytes[1] == 0x50 && bytes[2] == 0x4E && bytes[3] == 0x47) {
            type = OfficeImageFormat.Png; return true;
        }
        if (bytes.Length >= 3 && bytes[0] == 0xFF && bytes[1] == 0xD8 && bytes[2] == 0xFF) {
            type = OfficeImageFormat.Jpeg; return true;
        }
        if (bytes.Length >= 6 && bytes[0] == (byte)'G' && bytes[1] == (byte)'I' && bytes[2] == (byte)'F') {
            type = OfficeImageFormat.Gif; return true;
        }
        if (bytes.Length >= 2 && bytes[0] == (byte)'B' && bytes[1] == (byte)'M') {
            type = OfficeImageFormat.Bmp; return true;
        }
        if (bytes.Length >= 4 && ((bytes[0] == (byte)'I' && bytes[1] == (byte)'I' && bytes[2] == 42 && bytes[3] == 0) ||
                                (bytes[0] == (byte)'M' && bytes[1] == (byte)'M' && bytes[2] == 0 && bytes[3] == 42))) {
            type = OfficeImageFormat.Tiff; return true;
        }
        type = OfficeImageFormat.Png; return false;
    }

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

using System.Globalization;
using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private sealed partial class LayoutContext {
        private void RenderCanvasBlock(PdfCanvasBlock canvas) {
            EnsurePage();
            for (int itemIndex = 0; itemIndex < canvas.Items.Count; itemIndex++) {
                PdfCanvasItem item = canvas.Items[itemIndex];
                switch (item) {
                    case PdfCanvasAnchoredImageItem anchoredPicture:
                        RenderAnchoredPicture(anchoredPicture);
                        break;
                    case PdfCanvasBehindTextItem behindText:
                        RenderBehindTextCanvas(behindText);
                        break;
                    case PdfCanvasParagraphAnchorItem paragraphAnchor:
                        RenderCanvasEffect(new PdfCanvasEffectItem(paragraphAnchor.Items,
                            OfficeTransform.Translate(0D, currentOpts.PageHeight.ToPoints() - (paragraphCanvasTop ?? y)), 1D, OfficeBlendMode.Normal));
                        break;
                    case PdfCanvasMarginAnchorItem marginAnchor:
                        RenderCanvasEffect(new PdfCanvasEffectItem(marginAnchor.Items,
                            OfficeTransform.Translate(currentPage!.Options.MarginLeft.ToPoints() - marginAnchor.AuthoredLeftMargin, 0D),
                            1D, OfficeBlendMode.Normal));
                        break;
                    case PdfCanvasArtifactItem artifact:
                        RenderCanvasArtifact(artifact);
                        break;
                    case PdfCanvasActualTextItem actualText:
                        RenderCanvasActualText(actualText);
                        break;
                    case PdfCanvasStructureItem structure:
                        RenderCanvasStructure(structure);
                        break;
                    case PdfCanvasFigureItem figure:
                        RenderCanvasFigure(figure);
                        break;
                    case PdfCanvasOutlineItem outline:
                        RenderCanvasOutline(outline);
                        break;
                    case PdfCanvasNamedDestinationItem destination:
                        RenderCanvasNamedDestination(destination);
                        break;
                    case PdfCanvasNamedDestinationLinkItem destinationLink:
                        ResolveFloatingBookmarks(currentOpts.PageHeight.ToPoints() - destinationLink.Y);
                        RenderCanvasNamedDestinationLink(destinationLink);
                        break;
                    case PdfCanvasTextItem text:
                        ResolveFloatingBookmarks(currentOpts.PageHeight.ToPoints() - text.Y);
                        RenderCanvasText(text);
                        break;
                    case PdfCanvasSearchableTextItem searchableText:
                        ResolveFloatingBookmarks(currentOpts.PageHeight.ToPoints() - searchableText.Y);
                        if (!TryRenderSearchableRun(canvas.Items, ref itemIndex)) RenderCanvasSearchableText(searchableText);
                        break;
                    case PdfCanvasTextBoxItem textBox:
                        ResolveFloatingBookmarks(currentOpts.PageHeight.ToPoints() - textBox.Y);
                        RenderCanvasTextBox(textBox);
                        break;
                    case PdfCanvasShapeItem shape:
                        ResolveFloatingBookmarks(currentOpts.PageHeight.ToPoints() - shape.Y);
                        RenderCanvasShape(shape);
                        break;
                    case PdfCanvasDrawingItem drawing:
                        ResolveFloatingBookmarks(currentOpts.PageHeight.ToPoints() - drawing.Y);
                        RenderCanvasDrawing(drawing);
                        break;
                    case PdfCanvasImageItem image:
                        ResolveFloatingBookmarks(currentOpts.PageHeight.ToPoints() - image.Y);
                        RenderCanvasImage(image);
                        break;
                    case PdfCanvasTextAnnotationItem textAnnotation:
                        ResolveFloatingBookmarks(currentOpts.PageHeight.ToPoints() - textAnnotation.Y);
                        RenderCanvasTextAnnotation(textAnnotation);
                        break;
                    case PdfCanvasFreeTextAnnotationItem freeTextAnnotation:
                        ResolveFloatingBookmarks(currentOpts.PageHeight.ToPoints() - freeTextAnnotation.Y);
                        RenderCanvasFreeTextAnnotation(freeTextAnnotation);
                        break;
                    case PdfCanvasHighlightAnnotationItem highlightAnnotation:
                        ResolveFloatingBookmarks(currentOpts.PageHeight.ToPoints() - highlightAnnotation.Y);
                        RenderCanvasHighlightAnnotation(highlightAnnotation);
                        break;
                    case PdfCanvasFormFieldItem formField:
                        ResolveFloatingBookmarks(currentOpts.PageHeight.ToPoints() - formField.Y);
                        RenderCanvasFormField(formField);
                        break;
                    case PdfCanvasTableItem table:
                        ResolveFloatingBookmarks(currentOpts.PageHeight.ToPoints() - table.Y);
                        RenderCanvasTable(table);
                        break;
                    case PdfCanvasClipItem clip:
                        RenderCanvasClip(clip);
                        break;
                    case PdfCanvasEffectItem effect:
                        RenderCanvasEffect(effect);
                        break;
                }
            }
        }

        private void RenderCanvasArtifact(PdfCanvasArtifactItem item) {
            EnsurePage();
            LayoutResult.Page artifactPage = currentPage!;
            int imageCount = artifactPage.Images.Count;
            int linkAnnotationCount = artifactPage.Annotations.Count;
            int textAnnotationCount = artifactPage.TextAnnotations.Count;
            int freeTextAnnotationCount = artifactPage.FreeTextAnnotations.Count;
            int highlightAnnotationCount = artifactPage.HighlightAnnotations.Count;
            int formFieldCount = artifactPage.FormFields.Count;
            int bookmarkCount = artifactPage.Bookmarks.Count;
            sb.Append("/Artifact << /Type /Layout >> BDC\n");
            bool previousAccessibility = _suppressCanvasAccessibilityWrappers;
            bool previousStructure = _suppressCanvasStructureRegistration;
            _suppressCanvasAccessibilityWrappers = true;
            _suppressCanvasStructureRegistration = true;
            try {
                RenderCanvasBlock(new PdfCanvasBlock(item.Items));
            } finally {
                _suppressCanvasAccessibilityWrappers = previousAccessibility;
                _suppressCanvasStructureRegistration = previousStructure;
                // Raster paint also participates in compliance evidence. An explicit
                // artifact owns its images even though their automatic wrappers are suppressed.
                for (int index = imageCount; index < artifactPage.Images.Count; index++)
                    artifactPage.Images[index].IsInlineDecoration = true;
                RemoveCanvasArtifactInteractions(
                    artifactPage,
                    linkAnnotationCount,
                    textAnnotationCount,
                    freeTextAnnotationCount,
                    highlightAnnotationCount,
                    formFieldCount,
                    bookmarkCount);
            }
            sb.Append("EMC\n");
        }

        private void RenderCanvasNamedDestinationLink(PdfCanvasNamedDestinationLinkItem item) {
            ValidateCanvasBox(item.X, item.Y, item.Width, item.Height, "Canvas named-destination link");
            double topY = currentOpts.PageHeight.ToPoints() - item.Y;
            currentPage!.Annotations.Add(new LinkAnnotation {
                X1 = item.X,
                Y1 = topY - item.Height,
                X2 = item.X + item.Width,
                Y2 = topY,
                DestinationName = item.DestinationName,
                Contents = item.Contents
            });
            pageDirty = true;
        }

        private static void RemoveCanvasArtifactInteractions(
            LayoutResult.Page page,
            int linkAnnotationCount,
            int textAnnotationCount,
            int freeTextAnnotationCount,
            int highlightAnnotationCount,
            int formFieldCount,
            int bookmarkCount) {
            page.Annotations.RemoveRange(linkAnnotationCount, page.Annotations.Count - linkAnnotationCount);
            page.TextAnnotations.RemoveRange(textAnnotationCount, page.TextAnnotations.Count - textAnnotationCount);
            page.FreeTextAnnotations.RemoveRange(freeTextAnnotationCount, page.FreeTextAnnotations.Count - freeTextAnnotationCount);
            page.HighlightAnnotations.RemoveRange(highlightAnnotationCount, page.HighlightAnnotations.Count - highlightAnnotationCount);
            page.FormFields.RemoveRange(formFieldCount, page.FormFields.Count - formFieldCount);
            page.Bookmarks.RemoveRange(bookmarkCount, page.Bookmarks.Count - bookmarkCount);
        }

        private void RenderCanvasFormField(PdfCanvasFormFieldItem item) {
            ValidateCanvasBox(item.X, item.Y, item.Width, item.Height, "Canvas form field");
            double topY = currentOpts.PageHeight.ToPoints() - item.Y;
            double bottomY = topY - item.Height;

            if (item.Kind == PdfCanvasFormFieldKind.RadioButton) {
                var field = new FormFieldAnnotation {
                    X1 = item.X,
                    Y1 = bottomY,
                    X2 = item.X + item.Width,
                    Y2 = topY,
                    Kind = FormFieldAnnotationKind.RadioButtonGroup,
                    Name = item.Name,
                    Value = item.IsSelected ? item.Option : "Off",
                    Options = new[] { item.Option },
                    ExportValues = new[] { item.ExportValue },
                    ButtonSize = Math.Min(item.Width, item.Height),
                    Style = item.Style
                };
                field.RadioWidgets.Add(new RadioButtonWidgetAnnotation {
                    X1 = item.X,
                    Y1 = bottomY,
                    X2 = item.X + item.Width,
                    Y2 = topY,
                    Option = item.Option,
                    Style = item.Style,
                    StructureParentElement = _canvasStructureParentElement
                });
                field.StructureParentElement = _canvasStructureParentElement;
                CaptureFormNavigationTarget(field);
                currentPage!.FormFields.Add(field);
                pageDirty = true;
                return;
            }

            var annotation = new FormFieldAnnotation {
                X1 = item.X,
                Y1 = bottomY,
                X2 = item.X + item.Width,
                Y2 = topY,
                Kind = item.Kind == PdfCanvasFormFieldKind.Text
                    ? FormFieldAnnotationKind.Text
                    : item.Kind == PdfCanvasFormFieldKind.CheckBox
                        ? FormFieldAnnotationKind.CheckBox
                        : FormFieldAnnotationKind.Choice,
                Name = item.Name,
                Value = item.Value,
                AppearanceValue = item.AppearanceValue,
                AppearanceStyle = item.AppearanceStyle,
                Values = item.Values,
                FontSize = item.FontSize,
                IsChecked = item.IsSelected,
                CheckedValueName = item.Option,
                ExportValue = item.ExportValue,
                Options = item.Options,
                ChoiceOptions = item.ChoiceOptions,
                SelectedIndices = item.SelectedIndices,
                IsComboBox = item.IsComboBox,
                AllowsMultipleSelection = item.AllowsMultipleSelection,
                Style = item.Style,
                StructureParentElement = _canvasStructureParentElement
            };
            CaptureFormNavigationTarget(annotation);
            currentPage!.FormFields.Add(annotation);
            pageDirty = true;
        }

        private void RenderCanvasSearchableText(PdfCanvasSearchableTextItem item) {
            if (_suppressCanvasActualTextChildren) return;
            EnsurePage();
            PdfStandardFont font = ChooseNormal(currentOpts.DefaultFont);
            double baselineY = currentOpts.PageHeight.ToPoints() - item.Y - (item.UsesBounds ? item.Height : 0D);
            double baselineX = item.X, textWidth = item.Width, textHeight = item.Height;
            double a = 1D, b = 0D, c = 0D, d = 1D;
            if (item.Geometry is PdfSelectionQuad quad) {
                double advanceX = quad.BottomRight.X - quad.BottomLeft.X, advanceY = quad.BottomRight.Y - quad.BottomLeft.Y;
                double upX = quad.TopLeft.X - quad.BottomLeft.X, upY = quad.TopLeft.Y - quad.BottomLeft.Y;
                textWidth = Math.Sqrt(advanceX * advanceX + advanceY * advanceY);
                textHeight = Math.Sqrt(upX * upX + upY * upY);
                Guard.Positive(textWidth, nameof(item)); Guard.Positive(textHeight, nameof(item));
                a = advanceX / textWidth; b = -advanceY / textWidth;
                c = upX / textHeight; d = -upY / textHeight;
                baselineX = quad.BottomLeft.X; baselineY = currentOpts.PageHeight.ToPoints() - quad.BottomLeft.Y;
            }
            // One anchor per Unicode scalar keeps reader spacing heuristics proportional
            // to characters instead of presenting the entire word as one stretched space.
            int anchorCount = CountLogicalAnchorScalars(item.Text);
            PdfTextShowCommand anchor = EncodeBoundedLogicalTextAnchor(font, currentOpts, anchorCount);
            double horizontalScaling = ResolveLogicalAnchorScaling(anchor, textWidth, textHeight);
            int? markedContentId = RegisterTextStructureElement("Span", _canvasStructureParentElement, logicalOrder: item.LogicalOrder);

            var content = new ContentStreamBuilder(sb)
                .SaveState()
                .BeginText()
                .WordSpacing(0D).TextRise(0D)
                .HorizontalTextScaling(horizontalScaling)
                .TextRenderingMode(3)
                .TextMatrix(a, b, c, d, baselineX, baselineY);
            sb.Append("/Span << /ActualText ")
                .Append(PdfSyntaxEscaper.TextString(item.Text));
            if (markedContentId.HasValue) {
                sb.Append(" /MCID ")
                    .Append(markedContentId.Value.ToString(CultureInfo.InvariantCulture));
            }
            sb.Append(" >> BDC\n");
            PdfTextShowCommand space = EncodeBoundedLogicalTextAnchor(font, currentOpts, 1);
            bool cff = currentOpts.TryGetEmbeddedStandardOpenTypeCffFontProgram(font, out var _);
            foreach (var run in currentPage!.SearchableFonts.Encode(item.Text, space, cff)) {
                content.Font(run.Bank.Name, textHeight, preserveLogicalPrecision: item.UsesBounds).ShowHexText(run.Hex);
            }
            sb.Append("EMC\n");
            content.EndText().RestoreState();

            MarkSimpleFont(font);
            pageDirty = true;
        }

        private void RenderCanvasActualText(PdfCanvasActualTextItem item) {
            if (_suppressCanvasActualTextChildren) {
                RenderCanvasBlock(new PdfCanvasBlock(item.Items));
                return;
            }
            EnsurePage();
            double actualTextX = item.HasPosition ? item.X : 0D;
            double actualTextY = item.HasPosition ? currentOpts.PageHeight.ToPoints() - item.Y - (item.UsesBounds ? item.Height : 0D) : 0D;
            RenderLogicalText(item.Text, actualTextX, actualTextY,
                () => RenderCanvasBlock(new PdfCanvasBlock(item.Items)), item.Width, item.Height, item.LogicalOrder);
        }

        private void RenderCanvasStructure(PdfCanvasStructureItem item) {
            if (_suppressCanvasStructureRegistration) {
                RenderCanvasBlock(new PdfCanvasBlock(item.Items));
                return;
            }
            PageStructElement? structureElement = ResolveCanvasStructure(item.Role, item.Options);
            PageStructElement? previous = _canvasStructureParentElement;
            _canvasStructureParentElement = structureElement ?? previous;
            try {
                RenderCanvasBlock(new PdfCanvasBlock(item.Items));
            } finally {
                _canvasStructureParentElement = previous;
            }
        }

        private PageStructElement? ResolveCanvasStructure(PdfCanvasStructureRole role, PdfCanvasStructureOptions options) {
            string structureType = MapCanvasStructureType(role);
            string headerScope = MapCanvasTableHeaderScope(options.HeaderScope);
            string alternativeText = options.AlternativeText ?? string.Empty;
            string? language = options.Language ?? _canvasStructureParentElement?.Language ?? currentOpts.Language;
            PageStructElement? structureElement;
            var structureKey = (options.StructureElementKey ?? string.Empty, structureType, _canvasStructureParentElement, headerScope, options.ColumnSpan, options.RowSpan, alternativeText, language, options.ListNumbering);
            if (options.StructureElementKey != null && canvasStructureElements.TryGetValue(structureKey, out PageStructElement? existingStructureElement)) {
                structureElement = existingStructureElement;
                if (!currentPage!.StructElements.Contains(structureElement)) structureElement.SpansPages = true;
            } else {
                structureElement = RegisterStructureContainer(
                    structureType,
                    _canvasStructureParentElement,
                    headerScope,
                    options.ColumnSpan,
                    options.RowSpan,
                    options.AlternativeText);
                if (structureElement != null) {
                    structureElement.Language = language;
                    structureElement.AssociatedFiles = options.AssociatedFileSnapshots;
                    structureElement.LogicalOrder = options.LogicalOrder;
                    structureElement.ListNumbering = options.ListNumbering;
                    structureElement.StructureElementKey = options.StructureElementKey;
                }
                if (options.StructureElementKey != null && structureElement != null) {
                    canvasStructureElements[structureKey] = structureElement;
                }
            }
            return structureElement;
        }

        private static string MapCanvasStructureType(PdfCanvasStructureRole role) {
            if (role == PdfCanvasStructureRole.Section) return "Sect";
            if (role == PdfCanvasStructureRole.Division) return "Div";
            if (role == PdfCanvasStructureRole.Paragraph) return "P";
            if (role == PdfCanvasStructureRole.Heading1) return "H1";
            if (role == PdfCanvasStructureRole.Heading2) return "H2";
            if (role == PdfCanvasStructureRole.Heading3) return "H3";
            if (role == PdfCanvasStructureRole.Heading4) return "H4";
            if (role == PdfCanvasStructureRole.Heading5) return "H5";
            if (role == PdfCanvasStructureRole.Heading6) return "H6";
            if (role == PdfCanvasStructureRole.List) return "L";
            if (role == PdfCanvasStructureRole.ListItem) return "LI";
            if (role == PdfCanvasStructureRole.ListLabel) return "Lbl";
            if (role == PdfCanvasStructureRole.ListBody) return "LBody";
            if (role == PdfCanvasStructureRole.Table) return "Table";
            if (role == PdfCanvasStructureRole.TableRow) return "TR";
            if (role == PdfCanvasStructureRole.TableHeaderCell) return "TH";
            if (role == PdfCanvasStructureRole.TableCell) return "TD";
            if (role == PdfCanvasStructureRole.Caption) return "Caption";
            if (role == PdfCanvasStructureRole.Figure) return "Figure";
            if (role == PdfCanvasStructureRole.TableBody) return "TBody";
            if (role == PdfCanvasStructureRole.Span) return "Span";
            if (role == PdfCanvasStructureRole.Formula) return "Formula";
            return "Note";
        }

        private static string MapCanvasTableHeaderScope(PdfCanvasTableHeaderScope? scope) {
            if (scope == PdfCanvasTableHeaderScope.Row) return "Row";
            if (scope == PdfCanvasTableHeaderScope.Column) return "Column";
            if (scope == PdfCanvasTableHeaderScope.Both) return "Both";
            return string.Empty;
        }

        private void RenderCanvasFigure(PdfCanvasFigureItem item) {
            EnsurePage();
            int? markedContentId = RegisterFigureStructureElement(item.AlternativeText, _canvasStructureParentElement, item.LogicalOrder);
            sb.Append("/Figure << /Alt ")
                .Append(PdfSyntaxEscaper.TextString(item.AlternativeText));
            if (markedContentId.HasValue) {
                sb.Append(" /MCID ")
                    .Append(markedContentId.Value.ToString(CultureInfo.InvariantCulture));
            }

            sb.Append(" >> BDC\n");
            bool previous = _suppressCanvasAccessibilityWrappers;
            _suppressCanvasAccessibilityWrappers = true;
            try {
                RenderCanvasBlock(new PdfCanvasBlock(item.Items));
            } finally {
                _suppressCanvasAccessibilityWrappers = previous;
            }
            sb.Append("EMC\n");
        }

        private void RenderCanvasText(PdfCanvasTextItem item) {
            ValidateCanvasBox(item.X, item.Y, item.Width, item.Height, "Canvas text");
            double size = item.FontSize ?? currentOpts.DefaultFontSize.ToPoints();
            double leading = item.LineHeight ?? size * 1.2D;
            var block = new RichParagraphBlock(item.Runs, item.Align, item.DefaultColor);
            var wrap = item.PreservePositionedText
                ? CreatePositionedTextLine(item.Runs, size, leading, currentOpts, item.PositionedFontMetricScale)
                : WrapRichRunsCore(item.Runs, item.Width, size, ChooseNormal(currentOpts.DefaultFont), leading, null, DefaultParagraphTabStopWidth, currentOpts);
            if (wrap.Lines.Count == 0) {
                return;
            }

            if (!item.PreservePositionedText)
                ReportFullyClippedFrameLine("canvas-text-content-clipped", "PdfCanvasText",
                    wrap.Lines, wrap.LineHeights,
                    FirstTextBaselineFromTop(ChooseNormal(currentOpts.DefaultFont), size, currentOpts.PageHeight.ToPoints() - item.Y),
                    size, leading, ChooseNormal(currentOpts.DefaultFont), currentOpts.PageHeight.ToPoints() - item.Y - item.Height,
                    item.Height, item.X, currentOpts.PageHeight.ToPoints() - item.Y, item.Width, item.Height);

            double scaleX = 1D;
            if (item.PreservePositionedText && item.PositionedAdvanceWidth is double advance && advance > 0D && !double.IsInfinity(advance)) {
                double writtenWidth = MeasureRichLineWidth(wrap.Lines[0], currentOpts);
                if (writtenWidth > 0D && !double.IsInfinity(writtenWidth)) scaleX = advance / writtenWidth;
            }
            // Keep the adapter's line geometry when final fallback-font selection changes its metrics.
            if (Math.Abs(scaleX - 1D) > 0.000001D) {
                RenderOpaqueEffectGroupInline(new OfficeTransform(scaleX, 0D, 0D, 1D, item.X * (1D - scaleX), 0D),
                    () => WriteCanvasText(item, size, leading, block, wrap.Lines, wrap.LineHeights, item.Width / scaleX));
            } else {
                WriteCanvasText(item, size, leading, block, wrap.Lines, wrap.LineHeights, item.Width);
            }
        }

        private void WriteCanvasText(PdfCanvasTextItem item, double size, double leading, RichParagraphBlock block,
            List<List<RichSeg>> lines, List<double> lineHeights, double width) {
            double topY = currentOpts.PageHeight.ToPoints() - item.Y;
            double bottomY = topY - item.Height;
            string? structureType = _suppressCanvasAccessibilityWrappers ? null : MapCanvasTextStructureType(item.StructureRole);
            int? markedContentId = structureType == null ? null : RegisterTextStructureElement(structureType, _canvasStructureParentElement, logicalOrder: item.LogicalOrder);
            ReportRichPaintDiagnostics(PaintDiagnosticCodes.CanvasText, "PdfCanvasText", block, lines, lineHeights,
                item.PreservePositionedText ? topY - size : FirstTextBaselineFromTop(ChooseNormal(currentOpts.DefaultFont), size, topY),
                size, leading, item.X, width, new TableCellContentFrame(item.X, topY + item.PositionedClipTopOverflow, width, item.Height + item.PositionedClipTopOverflow),
                clipsToFrame: !item.PreservePositionedText);
            WriteClippedRichParagraph(
                sb,
                block,
                lines,
                lineHeights,
                currentOpts,
                item.PreservePositionedText ? topY - size : FirstTextBaselineFromTop(ChooseNormal(currentOpts.DefaultFont), size, topY),
                size,
                leading,
                currentPage!.Annotations,
                item.X,
                bottomY,
                width,
                item.Height + item.PositionedClipTopOverflow,
                item.X,
                width,
                structureType: structureType,
                markedContentId: markedContentId,
                structurePage: _suppressCanvasStructureRegistration ? null : currentPage,
                suppressActualText: _suppressCanvasActualTextChildren,
                clipToFrame: !item.PreservePositionedText, fontPage: currentPage);
            MarkRichFonts(item.Runs);
            DrawDebugCanvasItemBox(item.X, bottomY, width, item.Height);
            pageDirty = true;
        }

        private static string MapCanvasTextStructureType(PdfCanvasTextStructureRole role) {
            if (role == PdfCanvasTextStructureRole.Heading1) return "H1";
            if (role == PdfCanvasTextStructureRole.Heading2) return "H2";
            if (role == PdfCanvasTextStructureRole.Heading3) return "H3";
            if (role == PdfCanvasTextStructureRole.Heading4) return "H4";
            if (role == PdfCanvasTextStructureRole.Heading5) return "H5";
            if (role == PdfCanvasTextStructureRole.Heading6) return "H6";
            if (role == PdfCanvasTextStructureRole.Span) return "Span";
            return "P";
        }

        private void RenderCanvasTextBox(PdfCanvasTextBoxItem item) {
            ValidateCanvasBox(item.X, item.Y, item.Width, item.Height, "Canvas text box");
            PdfCanvasTextBoxStyle style = item.Style;
            double topY = currentOpts.PageHeight.ToPoints() - item.Y;
            double bottomY = topY - item.Height;
            bool rotated = item.RotationAngle != 0D;
            if (rotated) {
                BeginRotatedCanvasFrame(item.X, bottomY, item.Width, item.Height, item.RotationAngle);
            }

            if (CreateCanvasTextBoxShape(item.Width, item.Height, style) is { } shape) {
                ShapeBlock block = PdfDocument.CreateShapeBlock(shape, PdfAlign.Left, spacingBefore: 0D, spacingAfter: 0D);
                PdfDrawingStyle drawingStyle = ResolveDrawingStyle(block, currentOpts);
                PdfDocument.ValidateDrawingStyle(drawingStyle, "Canvas text box");
                DrawShapeAt(block, drawingStyle, item.X, item.Width, topY);
            }

            double paddingLeft = style.EffectivePaddingLeft;
            double paddingTop = style.EffectivePaddingTop;
            double paddingBottom = style.EffectivePaddingBottom;
            double textX = item.X + paddingLeft;
            double textHeight = item.Height - paddingTop - paddingBottom;
            double textTopY = topY - paddingTop;
            double textBottomY = bottomY + paddingBottom;
            var prepared = PrepareCanvasTextBoxText(item.Runs, item.Width, style, currentOpts);
            double textWidth = prepared.TextWidth;
            double size = prepared.FontSize;
            double leading = prepared.Leading;
            PdfStandardFont baseFont = prepared.BaseFont;
            var blockText = new RichParagraphBlock(item.Runs, style.Align, style.TextColor);
            if (prepared.Lines.Count > 0) {
                double textContentHeight = prepared.ContentHeight;
                double verticalOffset = GetCanvasTextBoxVerticalOffset(style.VerticalAlign, textHeight, textContentHeight);
                ReportFullyClippedFrameLine("canvas-text-box-content-clipped", "PdfCanvasTextBox",
                    prepared.Lines, prepared.LineHeights,
                    FirstTextBaselineFromTop(baseFont, size, textTopY - verticalOffset),
                    size, leading, baseFont, textBottomY, textHeight, item.X, topY, item.Width, item.Height);
                if (textContentHeight > textHeight + 0.01D) {
                    item.DiagnosticHandler?.Invoke(new PdfLayoutDiagnostic(
                        PdfLayoutDiagnosticKind.ClippedContent,
                        "PdfCanvasTextBox",
                        "The PDF text box render pass clipped text because wrapped content exceeded the available text area.",
                        item.X,
                        item.Y,
                        item.Width,
                        item.Height));
                }

                ReportRichPaintDiagnostics(PaintDiagnosticCodes.CanvasTextBox, "PdfCanvasTextBox", blockText, prepared.Lines, prepared.LineHeights,
                    FirstTextBaselineFromTop(baseFont, size, textTopY - verticalOffset), size, leading, textX, textWidth,
                    new TableCellContentFrame(textX, textTopY, textWidth, textHeight), baselineFont: baseFont,
                    localTransform: CreateCanvasRotationTransform(item.X, bottomY, item.Width, item.Height, item.RotationAngle));
                var annotations = rotated ? new System.Collections.Generic.List<LinkAnnotation>() : currentPage!.Annotations;
                int? markedContentId = _suppressCanvasAccessibilityWrappers ? null : RegisterTextStructureElement("P", _canvasStructureParentElement);
                WriteClippedRichParagraph(
                    sb,
                    blockText,
                    prepared.Lines,
                    prepared.LineHeights,
                    currentOpts,
                    FirstTextBaselineFromTop(baseFont, size, textTopY - verticalOffset),
                    size,
                    leading,
                    annotations,
                    textX,
                    textBottomY,
                    textWidth,
                    textHeight,
                    textX,
                    textWidth,
                    structureType: _suppressCanvasAccessibilityWrappers ? null : "P",
                    markedContentId: markedContentId,
                    structurePage: _suppressCanvasStructureRegistration ? null : currentPage,
                    suppressActualText: _suppressCanvasActualTextChildren,
                    baselineFont: baseFont, fontPage: currentPage);
                MarkRichFonts(item.Runs);
                if (rotated && annotations.Count > 0) {
                    RotateCanvasLinkAnnotations(annotations, item.X, bottomY, item.Width, item.Height, item.RotationAngle);
                    currentPage!.Annotations.AddRange(annotations);
                }
            }

            if (rotated) {
                new ContentStreamBuilder(sb)
                    .RestoreState();
            }

            DrawDebugCanvasItemBox(item.X, bottomY, item.Width, item.Height);
            pageDirty = true;
        }

        private void RenderCanvasShape(PdfCanvasShapeItem item) {
            ShapeBlock block = item.Block;
            ValidateCanvasBox(item.X, item.Y, block.Shape.Width, block.Shape.Height, "Canvas shape");
            PdfDrawingStyle style = ResolveDrawingStyle(block, currentOpts);
            PdfDocument.ValidateDrawingStyle(style, "Canvas shape");
            double topY = currentOpts.PageHeight.ToPoints() - item.Y;
            int annotationStart = currentPage!.Annotations.Count;
            int? structElementIndex = DrawShapeAt(block, style, item.X, block.Shape.Width, topY);
            if (!string.IsNullOrEmpty(block.LinkUri)) {
                currentPage.Annotations.Add(new LinkAnnotation {
                    X1 = item.X,
                    Y1 = topY - block.Shape.Height,
                    X2 = item.X + block.Shape.Width,
                    Y2 = topY,
                    Uri = block.LinkUri!,
                    Contents = block.LinkContents,
                    StructElementIndex = structElementIndex
                });
            }

            RotateCanvasLinkAnnotations(currentPage.Annotations, annotationStart, item.X, topY - block.Shape.Height, block.Shape.Width, block.Shape.Height, item.RotationAngle);
            DrawDebugCanvasItemBox(item.X, topY - block.Shape.Height, block.Shape.Width, block.Shape.Height);
            pageDirty = true;
        }

        private void RenderCanvasDrawing(PdfCanvasDrawingItem item) {
            DrawingBlock block = item.Block;
            ValidateCanvasBox(item.X, item.Y, item.Width, item.Height, "Canvas drawing");
            PdfDrawingStyle style = ResolveDrawingStyle(block, currentOpts);
            PdfDocument.ValidateDrawingStyle(style, "Canvas drawing");
            double topY = currentOpts.PageHeight.ToPoints() - item.Y;
            double bottomY = topY - item.Height;
            int annotationStart = currentPage!.Annotations.Count;
            int imageStart = currentPage.Images.Count;
            bool rotated = item.RotationAngle != 0D;
            if (rotated) {
                BeginRotatedCanvasFrame(item.X, bottomY, item.Width, item.Height, item.RotationAngle);
            }

            bool markedContent = false;
            int? structElementIndex = item.SourceStructure == null ? AppendDrawingMarkedContentBegin(
                style,
                out markedContent,
                recordDrawingEvidence: !ChildImagesOwnDrawingAccessibility(block.Drawing, style)) : null;
            var previousSourceStructure = _drawingSourceStructure;
            _drawingSourceStructure = item.SourceStructure;
            bool previousSuppressAccessibilityWrappers = _suppressCanvasAccessibilityWrappers;
            if (markedContent || style.Decorative || item.SourceStructure != null) {
                _suppressCanvasAccessibilityWrappers = true;
            }
            OfficeTransform previousEffectToPage = _canvasEffectToPage;
            if (rotated) _canvasEffectToPage = CreateCanvasRotationTransform(item.X, bottomY, item.Width, item.Height, item.RotationAngle)
                .Then(previousEffectToPage);
            try {
                double scaleX = item.Width / block.Drawing.Width;
                double scaleY = item.Height / block.Drawing.Height;
                bool scaled = Math.Abs(scaleX - 1D) > 0.0001D || Math.Abs(scaleY - 1D) > 0.0001D;
                if (scaled) {
                    DrawScaledCanvasDrawing(block.Drawing, scaleX, scaleY, item.X, topY);
                } else {
                    DrawDrawingElements(block.Drawing, item.X, topY);
                }
            } finally {
                _canvasEffectToPage = previousEffectToPage;
                _drawingSourceStructure = previousSourceStructure;
                _suppressCanvasAccessibilityWrappers = previousSuppressAccessibilityWrappers;
                AppendDrawingMarkedContentEnd(markedContent);
            }
            if (!string.IsNullOrEmpty(block.LinkUri)) {
                currentPage.Annotations.Add(new LinkAnnotation {
                    X1 = item.X,
                    Y1 = bottomY,
                    X2 = item.X + item.Width,
                    Y2 = topY,
                    Uri = block.LinkUri!,
                    Contents = block.LinkContents,
                    StructElementIndex = structElementIndex
                });
            }

            if (rotated) {
                new ContentStreamBuilder(sb)
                    .RestoreState();
                TransformCanvasPageImageBounds(
                    currentPage.Images,
                    imageStart,
                    CreateCanvasRotationTransform(item.X, bottomY, item.Width, item.Height, item.RotationAngle));
            }

            RotateCanvasLinkAnnotations(currentPage.Annotations, annotationStart, item.X, bottomY, item.Width, item.Height, item.RotationAngle);
            DrawDebugCanvasItemBox(item.X, bottomY, item.Width, item.Height);
            pageDirty = true;
        }

        private void RenderCanvasImage(PdfCanvasImageItem item) {
            ImageBlock block = item.Block;
            ValidateCanvasBox(item.X, item.Y, block.Width, block.Height, "Canvas image");
            PdfImageStyle imageStyle = ResolveImageStyle(block, currentOpts);
            PdfDocument.ValidateImageStyleForBox(imageStyle, block.Width, block.Height, nameof(imageStyle.ClipPath));
            PdfDocument.ValidateImageFitDimensions(block.Info, imageStyle.Fit, nameof(imageStyle.Fit));
            double bottomY = currentOpts.PageHeight.ToPoints() - item.Y - block.Height;
            if (block.Info.Format == OfficeIMO.Drawing.OfficeImageFormat.Svg) {
                RenderCanvasSvgImage(item, imageStyle, bottomY);
                return;
            }
            PageImage pageImage = CreatePageImage(block, imageStyle, item.X, bottomY, block.Width, block.Height);
            pageImage.SuppressAccessibilityWrapper = _suppressCanvasAccessibilityWrappers;
            pageImage.StructureParentElement = _canvasStructureParentElement;
            pageImage.LogicalOrder = item.LogicalOrder;
            pageImage.RotationAngle += item.RotationAngle;
            pageImage.HorizontalFlip ^= item.HorizontalFlip;
            pageImage.VerticalFlip ^= item.VerticalFlip;
            ReserveImageNavigationTarget(pageImage);
            currentPage!.Images.Add(pageImage);
            pageImage.IsForeground = item.Foreground;
            if (item.Foreground) pageImage.ForegroundTransform = _canvasEffectToPage;
            if (paragraphCanvasTop.HasValue) ReportAnchoredImagePageClipping(item, bottomY);
            pageImage.ForegroundZOrder = item.ForegroundZOrder;
            // Foreground images use a distinct final image layer; ordinary canvas order stays inline.
            if (!item.Foreground) {
                pageImage.InlineDrawToken = AllocateInlineImageDrawToken(currentPage);
                sb.Append(pageImage.InlineDrawToken);
            }

            AddImageLinkAnnotation(block, imageStyle, pageImage, item.X, bottomY, block.Width, block.Height);
            DrawDebugCanvasItemBox(item.X, bottomY, block.Width, block.Height);
            pageDirty = true;
        }

        private void ValidateCanvasBox(double x, double yFromTop, double boxWidth, double boxHeight, string name) {
            if (double.IsNaN(x) || double.IsNaN(yFromTop) || double.IsInfinity(x) || double.IsInfinity(yFromTop)) {
                throw new ArgumentOutOfRangeException(name, name + " coordinates must be finite.");
            }

            if (double.IsNaN(boxWidth) || double.IsNaN(boxHeight) || double.IsInfinity(boxWidth) || double.IsInfinity(boxHeight) || boxWidth < 0D || boxHeight < 0D || (boxWidth == 0D && boxHeight == 0D)) {
                throw new ArgumentOutOfRangeException(name, name + " dimensions must be finite non-negative values with at least one positive dimension.");
            }

            if (_canvasClipDepth > 0) {
                return;
            }

            if (x + boxWidth > currentOpts.PageWidth.ToPoints() + 0.001D || yFromTop + boxHeight > currentOpts.PageHeight.ToPoints() + 0.001D) {
                throw new ArgumentException(name + " exceeds the current page bounds.");
            }
        }

        private void BeginRotatedCanvasFrame(double x, double bottomY, double width, double height, double rotationAngle) {
            OfficeTransform transform = CreateCanvasRotationTransform(x, bottomY, width, height, rotationAngle);
            new ContentStreamBuilder(sb)
                .SaveState()
                .TransformMatrix(transform);
        }

        private static OfficeTransform CreateCanvasRotationTransform(double x, double bottomY, double width, double height, double rotationAngle) {
            double angle = rotationAngle * Math.PI / 180D;
            double cos = Math.Cos(angle);
            double sin = Math.Sin(angle);
            double centerX = x + width / 2D;
            double centerY = bottomY + height / 2D;
            double e = centerX - cos * centerX + sin * centerY;
            double f = centerY - sin * centerX - cos * centerY;
            return new OfficeTransform(cos, sin, -sin, cos, e, f);
        }

        private static void RotateCanvasLinkAnnotations(System.Collections.Generic.List<LinkAnnotation> annotations, double x, double bottomY, double width, double height, double rotationAngle) {
            RotateCanvasLinkAnnotations(annotations, 0, x, bottomY, width, height, rotationAngle);
        }

        private static void RotateCanvasLinkAnnotations(System.Collections.Generic.List<LinkAnnotation> annotations, int startIndex, double x, double bottomY, double width, double height, double rotationAngle) {
            if (rotationAngle == 0D) {
                return;
            }

            double angle = rotationAngle * Math.PI / 180D;
            double cos = Math.Cos(angle);
            double sin = Math.Sin(angle);
            double centerX = x + width / 2D;
            double centerY = bottomY + height / 2D;
            for (int i = startIndex; i < annotations.Count; i++) {
                RotateCanvasLinkAnnotation(annotations[i], centerX, centerY, cos, sin);
            }
        }

        private static void RotateCanvasPageImages(System.Collections.Generic.List<PageImage> images, int startIndex, double x, double bottomY, double width, double height, double rotationAngle) {
            if (rotationAngle == 0D) {
                return;
            }

            OfficeTransform transform = CreateCanvasRotationTransform(x, bottomY, width, height, rotationAngle);
            TransformCanvasPageImageBounds(images, startIndex, transform);
            double angle = rotationAngle * Math.PI / 180D;
            double cos = Math.Cos(angle);
            double sin = Math.Sin(angle);
            double centerX = x + width / 2D;
            double centerY = bottomY + height / 2D;
            for (int i = startIndex; i < images.Count; i++) {
                PageImage image = images[i];
                if (!string.IsNullOrEmpty(image.InlineDrawToken)) {
                    continue;
                }

                double imageCenterX = image.X + image.W / 2D;
                double imageCenterY = image.Y + image.H / 2D;
                RotateCanvasPoint(imageCenterX, imageCenterY, centerX, centerY, cos, sin, out double rotatedCenterX, out double rotatedCenterY);
                image.X = rotatedCenterX - image.W / 2D;
                image.Y = rotatedCenterY - image.H / 2D;
                image.RotationAngle += rotationAngle;
            }
        }

        private static void RotateCanvasFormFields(System.Collections.Generic.List<FormFieldAnnotation> formFields, int startIndex, double x, double bottomY, double width, double height, double rotationAngle) {
            if (rotationAngle == 0D) {
                return;
            }

            double angle = rotationAngle * Math.PI / 180D;
            double cos = Math.Cos(angle);
            double sin = Math.Sin(angle);
            double centerX = x + width / 2D;
            double centerY = bottomY + height / 2D;
            for (int i = startIndex; i < formFields.Count; i++) {
                RotateCanvasFormField(formFields[i], centerX, centerY, cos, sin);
            }
        }

        private static void RotateCanvasLinkAnnotation(LinkAnnotation annotation, double centerX, double centerY, double cos, double sin) {
            RotateCanvasPoint(annotation.X1, annotation.Y1, centerX, centerY, cos, sin, out double x1, out double y1);
            RotateCanvasPoint(annotation.X1, annotation.Y2, centerX, centerY, cos, sin, out double x2, out double y2);
            RotateCanvasPoint(annotation.X2, annotation.Y1, centerX, centerY, cos, sin, out double x3, out double y3);
            RotateCanvasPoint(annotation.X2, annotation.Y2, centerX, centerY, cos, sin, out double x4, out double y4);

            annotation.X1 = Math.Min(Math.Min(x1, x2), Math.Min(x3, x4));
            annotation.Y1 = Math.Min(Math.Min(y1, y2), Math.Min(y3, y4));
            annotation.X2 = Math.Max(Math.Max(x1, x2), Math.Max(x3, x4));
            annotation.Y2 = Math.Max(Math.Max(y1, y2), Math.Max(y3, y4));
        }

        private static void RotateCanvasFormField(FormFieldAnnotation formField, double centerX, double centerY, double cos, double sin) {
            RotateCanvasPoint(formField.X1, formField.Y1, centerX, centerY, cos, sin, out double x1, out double y1);
            RotateCanvasPoint(formField.X1, formField.Y2, centerX, centerY, cos, sin, out double x2, out double y2);
            RotateCanvasPoint(formField.X2, formField.Y1, centerX, centerY, cos, sin, out double x3, out double y3);
            RotateCanvasPoint(formField.X2, formField.Y2, centerX, centerY, cos, sin, out double x4, out double y4);

            formField.X1 = Math.Min(Math.Min(x1, x2), Math.Min(x3, x4));
            formField.Y1 = Math.Min(Math.Min(y1, y2), Math.Min(y3, y4));
            formField.X2 = Math.Max(Math.Max(x1, x2), Math.Max(x3, x4));
            formField.Y2 = Math.Max(Math.Max(y1, y2), Math.Max(y3, y4));
            for (int index = 0; index < formField.RadioWidgets.Count; index++) {
                RadioButtonWidgetAnnotation widget = formField.RadioWidgets[index];
                RotateCanvasRectangle(widget, centerX, centerY, cos, sin);
            }
        }

        private static void RotateCanvasRectangle(RadioButtonWidgetAnnotation widget, double centerX, double centerY, double cos, double sin) {
            RotateCanvasPoint(widget.X1, widget.Y1, centerX, centerY, cos, sin, out double x1, out double y1);
            RotateCanvasPoint(widget.X1, widget.Y2, centerX, centerY, cos, sin, out double x2, out double y2);
            RotateCanvasPoint(widget.X2, widget.Y1, centerX, centerY, cos, sin, out double x3, out double y3);
            RotateCanvasPoint(widget.X2, widget.Y2, centerX, centerY, cos, sin, out double x4, out double y4);
            widget.X1 = Math.Min(Math.Min(x1, x2), Math.Min(x3, x4));
            widget.Y1 = Math.Min(Math.Min(y1, y2), Math.Min(y3, y4));
            widget.X2 = Math.Max(Math.Max(x1, x2), Math.Max(x3, x4));
            widget.Y2 = Math.Max(Math.Max(y1, y2), Math.Max(y3, y4));
        }

        private static void RotateCanvasPoint(double x, double y, double centerX, double centerY, double cos, double sin, out double rotatedX, out double rotatedY) {
            double dx = x - centerX;
            double dy = y - centerY;
            rotatedX = centerX + cos * dx - sin * dy;
            rotatedY = centerY + sin * dx + cos * dy;
        }
    }
}

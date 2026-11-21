using W = DocumentFormat.OpenXml.Wordprocessing;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Word.Pdf;

public static partial class WordPdfConverterExtensions {
    /// <summary>
    /// Renders consecutive paragraphs that share one <c>w:framePr</c> as a positioned frame.
    /// Word lays a text frame out like a borderless one-cell floating table: the same anchors,
    /// alignment, offsets, text clearance and wrapping. The frame takes no vertical flow
    /// unless it wraps as <c>notBeside</c>.
    /// </summary>
    private static bool TryRenderNativeParagraphFrame(INativePdfFlow pdf, IReadOnlyList<WordElement> elements, ref int index,
        Func<WordParagraph, (int Level, string Marker)?> getMarker, NativeNoteNumbering notes, WordPdfExportOptions? options,
        double? contentWidth, NativeDocumentDefaults defaults, NativeFontMap fontMap,
        IReadOnlyDictionary<W.Paragraph, string>? headingDestinations = null) {
        if (elements[index] is not WordParagraph first || !CanRenderNativeParagraphFrame(first) ||
            ResolveNativeParagraphFrame(first) is not { } frame) return false;
        bool notBeside = frame.Wrap?.Value == W.TextWrappingValues.NotBeside;
        if (!notBeside && !pdf.SupportsPositionedTables) return false;

        var paragraphs = new List<WordParagraph> { first };
        string key = frame.OuterXml;
        while (index + paragraphs.Count < elements.Count && elements[index + paragraphs.Count] is WordParagraph next &&
               CanRenderNativeParagraphFrame(next) && ResolveNativeParagraphFrame(next) is { } nextFrame && nextFrame.OuterXml == key) paragraphs.Add(next);

        WordTable table = CreateNativeFrameTable(first._document, frame, paragraphs, notBeside);
        foreach (WordParagraph paragraph in paragraphs) {
            if (paragraph._paragraph != null) options?.DiagnosedFrameParagraphs.Add(paragraph._paragraph);
        }

        // The frame content is a copy: list markers, numbering and story relationships still
        // resolve through the source paragraphs.
        var sources = new Dictionary<W.Paragraph, W.Paragraph>();
        List<W.Paragraph> directCopies = table._table.Descendants<W.TableCell>().First().Elements<W.Paragraph>().ToList();
        for (int i = 0; i < paragraphs.Count && i < directCopies.Count; i++) {
            W.Paragraph source = paragraphs[i]._paragraph!;
            sources[directCopies[i]] = source;
            List<W.Paragraph> nestedCopies = directCopies[i].Descendants<W.Paragraph>().ToList();
            List<W.Paragraph> nestedSources = source.Descendants<W.Paragraph>().ToList();
            for (int nested = 0; nested < nestedCopies.Count && nested < nestedSources.Count; nested++)
                sources[nestedCopies[nested]] = nestedSources[nested];
        }
        List<W.Paragraph> copies = sources.Keys.ToList();
        Func<WordParagraph, (int Level, string Marker)?> frameMarker = paragraph =>
            paragraph._paragraph != null && sources.TryGetValue(paragraph._paragraph, out W.Paragraph? source)
                ? getMarker(new WordParagraph(first._document, source)) : getMarker(paragraph);

        if (defaults.RunningContentContext == null) {
            // Bookmarks in the frame stay link targets at the frame's position.
            foreach (WordParagraph paragraph in paragraphs) {
                RenderNativeParagraphBookmarks(pdf, paragraph);
                if (paragraph._paragraph != null && !GetNativeParagraphBookmarkNames(paragraph).Any() &&
                    headingDestinations != null && headingDestinations.TryGetValue(paragraph._paragraph, out string? destination))
                    pdf.Bookmark(destination);
            }
        }

        var collapseFlow = pdf as NativeSpacingCollapseFlow;
        var collapseState = collapseFlow?.CaptureState();
        PdfCore.TableBlock? block;
        // The copies are diagnosed through their sources; their entries last only while the frame is built.
        foreach (W.Paragraph copy in copies) options?.DiagnosedFrameParagraphs.Add(copy);
        try {
            using (WordDocumentTraversal.UseListParagraphAliases(first._document, sources))
            using (WordPartOwnership.UseDetachedOwner(table._table, WordPartOwnership.Resolve(first._document, first._paragraph!)))
                block = BuildNativeTableBlock(table, frameMarker, notes, options, contentWidth, defaults, fontMap, pdf.SupportsPositionedTables);
        } finally {
            foreach (W.Paragraph copy in copies) options?.DiagnosedFrameParagraphs.Remove(copy);
        }
        if (block != null) {
            if (block.Style is { Position: { } position } blockStyle) blockStyle.Position = CreateNativeFramePosition(position, frame);
            pdf.Table(block.Cells.Select(row => row.ToArray()), block.Align, block.Style);
        }
        if (collapseFlow != null && collapseState is { } state) {
            if (!notBeside) {
                // A floating frame leaves the flow untouched: the next paragraph collapses its
                // space before against the paragraph above the frame. At a section start Word
                // drops that paragraph's space before, as the frame starts the page.
                collapseFlow.RestorePendingSpacingAfter(state.Pending);
            } else if (state.Preserve) {
                // A notBeside frame at a section start passes the carried spacing on to the
                // paragraph after it, which does not collapse with the frame's space after.
                collapseFlow.RestoreState(state);
            } else {
                // Elsewhere the next paragraph collapses with the frame's last space after.
                collapseFlow.RestorePendingSpacingAfter(CreateNativeParagraphStyle(paragraphs[paragraphs.Count - 1], defaults, fontMap).SpacingAfter?.ToPoints());
            }
        }
        index += paragraphs.Count - 1;
        return true;
    }

    /// <summary>
    /// Frames are placed for paragraphs directly in the body or in a header or footer whose
    /// content the frame projection draws in full: text, fields, inline pictures and form
    /// controls. Frames in table cells, notes, text boxes and multi-column sections, frames that
    /// start a page or hold a page or column break, body headings, and frames with paragraph
    /// shading, borders, charts, SmartArt, shapes, text boxes, floating pictures, embedded
    /// objects, equations, note references or page-number fields in the body keep the
    /// ordinary-paragraph rendering, which already draws those. Drop caps are not frames here.
    /// </summary>
    private static bool CanRenderNativeParagraphFrame(WordParagraph paragraph) {
        if (paragraph._paragraph?.Parent is not (W.Body or W.Header or W.Footer) || ResolveNativeParagraphFrame(paragraph) == null) return false;
        if (paragraph._paragraph.Descendants<W.Break>().Any(item => item.Type?.Value == W.BreakValues.Page || item.Type?.Value == W.BreakValues.Column))
            return false;
        if (HasNativePageBreakBefore(paragraph) || GetNativeEffectiveParagraphShadingFill(paragraph) != null ||
            HasNativeFrameParagraphBorders(paragraph) || GetNativeRuns(paragraph).Any(run => run.Chart != null) ||
            HasNativeFrameUnsupportedContent(paragraph._paragraph, inBody: paragraph._paragraph.Parent is W.Body)) return false;
        // A framed heading keeps the heading route, which also owns its outline entry.
        if (paragraph._paragraph.Parent is W.Body && GetHeadingLevel(paragraph) > 0) return false;
        return paragraph._paragraph.Parent is not W.Body || (paragraph.GetBodySection()?.ColumnCount ?? 1) <= 1;
    }

    private const string NativeFramePictureGraphicUri = "http://schemas.openxmlformats.org/drawingml/2006/picture";
    private static readonly string[] NativeFramePageFieldNames = { "PAGE", "NUMPAGES", "SECTIONPAGES", "SECTION", "PAGEREF" };

    private static bool HasNativeFrameUnsupportedContent(W.Paragraph paragraph, bool inBody) {
        if (inBody && HasNativeFramePageComplexField(paragraph)) return true;
        foreach (DocumentFormat.OpenXml.OpenXmlElement element in paragraph.Descendants()) {
            switch (element) {
                case DocumentFormat.OpenXml.Drawing.Wordprocessing.Anchor:
                case W.Picture:
                case W.EmbeddedObject:
                case W.TextBoxContent:
                case DocumentFormat.OpenXml.AlternateContent:
                case DocumentFormat.OpenXml.Math.OfficeMath:
                case DocumentFormat.OpenXml.Math.Paragraph:
                case W.FootnoteReference:
                case W.EndnoteReference:
                    return true;
                case DocumentFormat.OpenXml.Drawing.GraphicData data when data.Uri?.Value != NativeFramePictureGraphicUri:
                    return true;
                case W.SimpleField field when inBody && IsNativeFramePageField(field.Instruction?.Value):
                    return true;
            }
        }
        return false;
    }

    /// <summary>
    /// Reads each complex field's instruction as a whole, since Word often splits it across
    /// runs (" PA" + "GE "), including nested fields, and reports whether one paginates.
    /// </summary>
    private static bool HasNativeFramePageComplexField(W.Paragraph paragraph) {
        var instructions = new Stack<System.Text.StringBuilder?>();
        foreach (DocumentFormat.OpenXml.OpenXmlElement element in paragraph.Descendants()) {
            if (element is W.FieldChar fieldChar) {
                W.FieldCharValues? type = fieldChar.FieldCharType?.Value;
                if (type == W.FieldCharValues.Begin) {
                    instructions.Push(new System.Text.StringBuilder());
                } else if (instructions.Count > 0 && (type == W.FieldCharValues.Separate || type == W.FieldCharValues.End)) {
                    System.Text.StringBuilder? instruction = instructions.Pop();
                    if (instruction != null && IsNativeFramePageField(instruction.ToString())) return true;
                    // After the separator the field shows its result until its end mark.
                    if (type == W.FieldCharValues.Separate) instructions.Push(null);
                }
            } else if (element is W.FieldCode code && instructions.Count > 0 && instructions.Peek() is { } current) {
                current.Append(code.Text);
            }
        }
        return false;
    }

    private static bool IsNativeFramePageField(string? instruction) {
        string name = (instruction ?? string.Empty).Trim().Split(new[] { ' ', '\t' }, StringSplitOptions.RemoveEmptyEntries).FirstOrDefault() ?? string.Empty;
        return NativeFramePageFieldNames.Any(field => string.Equals(field, name, StringComparison.OrdinalIgnoreCase));
    }

    private static bool HasNativeFrameParagraphBorders(WordParagraph paragraph) {
        NativeParagraphBorders borders = MergeNativeParagraphBorders(GetNativeParagraphStyleDefaults(paragraph).Borders,
            paragraph._paragraph?.ParagraphProperties?.GetFirstChild<W.ParagraphBorders>());
        static bool Visible(NativeParagraphBorderSide side) => !side.IsEmpty && side.Style != W.BorderValues.None && side.Style != W.BorderValues.Nil;
        return Visible(borders.Top) || Visible(borders.Bottom) || Visible(borders.Left) || Visible(borders.Right) || Visible(borders.Between);
    }

    /// <summary>Returns the effective frame of a paragraph: direct formatting first, then its style chain.</summary>
    private static W.FrameProperties? ResolveNativeParagraphFrame(WordParagraph paragraph) {
        if (paragraph._paragraph == null) return null;
        W.FrameProperties? frame = paragraph._paragraph.ParagraphProperties?.GetFirstChild<W.FrameProperties>();
        if (frame == null) {
            foreach (W.Style style in GetNativeParagraphStyleChain(paragraph._document, paragraph.StyleId).Reverse()) {
                frame = style.StyleParagraphProperties?.GetFirstChild<W.FrameProperties>();
                if (frame != null) break;
            }
        }

        // Drop caps are frames too, but they follow their own text-flow rules.
        if (frame == null || frame.DropCap?.Value is W.DropCapLocationValues dropCap && dropCap != W.DropCapLocationValues.None) return null;
        return frame;
    }

    private static WordTable CreateNativeFrameTable(WordDocument document, W.FrameProperties frame,
        IReadOnlyList<WordParagraph> paragraphs, bool notBeside) {
        int? width = ParseNativeFrameTwips(frame.Width?.Value);
        // A reserved, undefined style keeps the document's default table style from formatting
        // frame text that was never in a table.
        var tableProperties = new W.TableProperties(new W.TableStyle { Val = NativeFrameTableStyleId });
        if (!notBeside) {
            var position = new W.TablePositionProperties {
                LeftFromText = (short?)ParseNativeFrameTwips(frame.HorizontalSpace?.Value),
                RightFromText = (short?)ParseNativeFrameTwips(frame.HorizontalSpace?.Value),
                TopFromText = (short?)ParseNativeFrameTwips(frame.VerticalSpace?.Value),
                BottomFromText = (short?)ParseNativeFrameTwips(frame.VerticalSpace?.Value),
                VerticalAnchor = frame.VerticalPosition?.Value == W.VerticalAnchorValues.Text ? W.VerticalAnchorValues.Text
                    : frame.VerticalPosition?.Value == W.VerticalAnchorValues.Margin ? W.VerticalAnchorValues.Margin : W.VerticalAnchorValues.Page,
                HorizontalAnchor = frame.HorizontalPosition?.Value == W.HorizontalAnchorValues.Text ? W.HorizontalAnchorValues.Text
                    : frame.HorizontalPosition?.Value == W.HorizontalAnchorValues.Margin ? W.HorizontalAnchorValues.Margin : W.HorizontalAnchorValues.Page
            };
            if (frame.XAlign?.Value is W.HorizontalAlignmentValues xAlign) position.TablePositionXAlignment = xAlign;
            else position.TablePositionX = ParseNativeFrameTwips(frame.X?.Value) ?? 0;
            if (frame.YAlign?.Value is W.VerticalAlignmentValues yAlign && yAlign != W.VerticalAlignmentValues.Inline) position.TablePositionYAlignment = yAlign;
            else position.TablePositionY = ParseNativeFrameTwips(frame.Y?.Value) ?? 0;
            tableProperties.Append(position);
            tableProperties.Append(new W.TableOverlap { Val = W.TableOverlapValues.Overlap });
        } else {
            ApplyNativeNotBesideFramePlacement(tableProperties, frame);
        }

        tableProperties.Append(width.HasValue
            ? new W.TableWidth { Width = width.Value.ToString(System.Globalization.CultureInfo.InvariantCulture), Type = W.TableWidthUnitValues.Dxa }
            : new W.TableWidth { Width = "0", Type = W.TableWidthUnitValues.Auto });
        tableProperties.Append(new W.TableBorders(
            new W.TopBorder { Val = W.BorderValues.None }, new W.LeftBorder { Val = W.BorderValues.None },
            new W.BottomBorder { Val = W.BorderValues.None }, new W.RightBorder { Val = W.BorderValues.None },
            new W.InsideHorizontalBorder { Val = W.BorderValues.None }, new W.InsideVerticalBorder { Val = W.BorderValues.None }));
        tableProperties.Append(new W.TableLayout { Type = width.HasValue ? W.TableLayoutValues.Fixed : W.TableLayoutValues.Autofit });
        tableProperties.Append(new W.TableCellMarginDefault(
            new W.TopMargin { Width = "0", Type = W.TableWidthUnitValues.Dxa },
            new W.TableCellLeftMargin { Width = 0, Type = W.TableWidthValues.Dxa },
            new W.BottomMargin { Width = "0", Type = W.TableWidthUnitValues.Dxa },
            new W.TableCellRightMargin { Width = 0, Type = W.TableWidthValues.Dxa }));

        var cell = new W.TableCell(new W.TableCellProperties(width.HasValue
            ? new W.TableCellWidth { Width = width.Value.ToString(System.Globalization.CultureInfo.InvariantCulture), Type = W.TableWidthUnitValues.Dxa }
            : new W.TableCellWidth { Width = "0", Type = W.TableWidthUnitValues.Auto }));
        for (int i = 0; i < paragraphs.Count; i++) {
            var clone = (W.Paragraph)paragraphs[i]._paragraph!.CloneNode(true);
            W.ParagraphProperties properties = clone.ParagraphProperties ??= new W.ParagraphProperties();
            properties.GetFirstChild<W.FrameProperties>()?.Remove();
            // A floating frame ends at its last line: the final space after is outside it.
            if (i == paragraphs.Count - 1 && !notBeside) {
                W.SpacingBetweenLines spacing = properties.GetFirstChild<W.SpacingBetweenLines>() ?? new W.SpacingBetweenLines();
                if (spacing.Parent == null) properties.AddChild(spacing, true);
                spacing.After = "0";
                spacing.AfterLines = null;
                spacing.AfterAutoSpacing = null;
            }
            cell.Append(clone);
        }

        var row = new W.TableRow();
        if (frame.Height?.Value is uint height && height > 0) {
            row.Append(new W.TableRowProperties(new W.TableRowHeight {
                Val = Math.Min(height, 31680U),
                HeightType = frame.HeightType?.Value == W.HeightRuleValues.Exact ? W.HeightRuleValues.Exact : W.HeightRuleValues.AtLeast
            }));
        }
        row.Append(cell);
        var grid = new W.TableGrid(new W.GridColumn { Width = (width ?? 0).ToString(System.Globalization.CultureInfo.InvariantCulture) });
        return new WordTable(document, new W.Table(tableProperties, grid, row));
    }

    private static PdfCore.PdfTablePosition CreateNativeFramePosition(PdfCore.PdfTablePosition position, W.FrameProperties frame) =>
        // Word treats a line that starts exactly at the frame's bottom edge as beside the
        // frame, so the clearance extends by one twip. Frames without wrapping let text
        // run underneath them.
        new(position.HorizontalAnchor, position.VerticalAnchor, position.HorizontalAlignment, position.VerticalAlignment,
            position.HorizontalOffset, position.VerticalOffset, position.DistanceLeft, position.DistanceRight, position.DistanceTop,
            global::OfficeIMO.OfficeLength.Points(position.DistanceBottom.ToPoints() + 0.05D),
            position.MirrorHorizontalOnEvenPages, position.AllowOverlap) {
            WrapsText = frame.Wrap?.Value != W.TextWrappingValues.None,
            AlignsInPageMargin = frame.XAlign?.Value is W.HorizontalAlignmentValues xAlign && xAlign != W.HorizontalAlignmentValues.Center
        };

    /// <summary>A notBeside frame keeps its own lines in the flow; only its horizontal placement applies.</summary>
    private static void ApplyNativeNotBesideFramePlacement(W.TableProperties properties, W.FrameProperties frame) {
        W.HorizontalAlignmentValues? xAlign = frame.XAlign?.Value;
        if (xAlign == W.HorizontalAlignmentValues.Right || xAlign == W.HorizontalAlignmentValues.Outside) {
            properties.Append(new W.TableJustification { Val = W.TableRowAlignmentValues.Right });
        } else if (xAlign == W.HorizontalAlignmentValues.Center) {
            properties.Append(new W.TableJustification { Val = W.TableRowAlignmentValues.Center });
        } else if (!xAlign.HasValue && ParseNativeFrameTwips(frame.X?.Value) is int x && x != 0 &&
                   frame.HorizontalPosition?.Value != W.HorizontalAnchorValues.Page) {
            properties.Append(new W.TableIndentation { Width = x, Type = W.TableWidthUnitValues.Dxa });
        }
    }

    private const string NativeFrameTableStyleId = "OfficeIMONativeFrameProjection";

    private static int? ParseNativeFrameTwips(string? value) =>
        int.TryParse(value, System.Globalization.NumberStyles.Integer, System.Globalization.CultureInfo.InvariantCulture, out int twips)
            ? Math.Max(-31680, Math.Min(31680, twips)) : null;
}

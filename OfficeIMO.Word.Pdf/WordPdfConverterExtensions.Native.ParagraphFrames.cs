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
        double? contentWidth, NativeDocumentDefaults defaults, NativeFontMap fontMap) {
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
        foreach (W.Paragraph clone in table._table.Descendants<W.Paragraph>()) options?.DiagnosedFrameParagraphs.Add(clone);

        var collapseFlow = pdf as NativeSpacingCollapseFlow;
        var collapseState = collapseFlow?.CaptureState();
        PdfCore.TableBlock? block = BuildNativeTableBlock(table, getMarker, notes, options, contentWidth, defaults, fontMap, pdf.SupportsPositionedTables);
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
    /// Frames are placed for paragraphs directly in the body or in a header or footer. Frames
    /// in table cells, notes, text boxes and multi-column sections, frames holding a page or
    /// column break, and drop caps keep the ordinary-paragraph approximation.
    /// </summary>
    private static bool CanRenderNativeParagraphFrame(WordParagraph paragraph) {
        if (paragraph._paragraph?.Parent is not (W.Body or W.Header or W.Footer) || ResolveNativeParagraphFrame(paragraph) == null) return false;
        if (paragraph._paragraph.Descendants<W.Break>().Any(item => item.Type?.Value == W.BreakValues.Page || item.Type?.Value == W.BreakValues.Column))
            return false;
        return paragraph._paragraph.Parent is not W.Body || (paragraph.GetBodySection()?.ColumnCount ?? 1) <= 1;
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
        var tableProperties = new W.TableProperties();
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

    private static int? ParseNativeFrameTwips(string? value) =>
        int.TryParse(value, System.Globalization.NumberStyles.Integer, System.Globalization.CultureInfo.InvariantCulture, out int twips)
            ? Math.Max(-31680, Math.Min(31680, twips)) : null;
}

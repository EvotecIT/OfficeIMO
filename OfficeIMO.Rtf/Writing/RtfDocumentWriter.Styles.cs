namespace OfficeIMO.Rtf.Writing;

internal static partial class RtfDocumentWriter {
    private static void WriteStyleSheet(StringBuilder builder, RtfDocument document, int unicodeSkipCount) {
        if (document.Styles.Count == 0) return;

        builder.Append(@"{\stylesheet");
        foreach (RtfStyle style in document.Styles.OrderBy(style => style.Kind).ThenBy(style => style.Id)) {
            WriteStyle(builder, style, unicodeSkipCount);
        }

        builder.Append('}');
    }

    private static void WriteStyle(StringBuilder builder, RtfStyle style, int unicodeSkipCount) {
        builder.Append('{');
        if (style.Kind == RtfStyleKind.Character) {
            builder.Append(@"\*\cs");
        } else if (style.Kind == RtfStyleKind.Table) {
            builder.Append(@"\*\ts");
        } else {
            builder.Append(@"\s");
        }

        builder.Append(style.Id.ToString(CultureInfo.InvariantCulture));
        WriteStyleKeyCode(builder, style.KeyCode, unicodeSkipCount);
        AppendOptionalTwips(builder, @"\sbasedon", style.BasedOnStyleId);
        AppendOptionalTwips(builder, @"\snext", style.NextStyleId);
        AppendOptionalTwips(builder, @"\slink", style.LinkedStyleId);
        AppendStyleFlag(builder, @"\additive", style.Additive);
        AppendStyleFlag(builder, @"\sautoupd", style.AutoUpdate);
        AppendStyleFlag(builder, @"\shidden", style.Hidden);
        AppendStyleFlag(builder, @"\slocked", style.Locked);
        AppendStyleFlag(builder, @"\spersonal", style.Personal);
        AppendStyleFlag(builder, @"\scompose", style.Compose);
        AppendStyleFlag(builder, @"\sreply", style.Reply);
        AppendStyleFlag(builder, @"\ssemihidden", style.SemiHidden);
        AppendStyleFlag(builder, @"\sunhideused", style.UnhideWhenUsed);
        AppendStyleFlag(builder, @"\sqformat", style.QuickFormat);
        AppendOptionalTwips(builder, @"\spriority", style.Priority);
        AppendOptionalTwips(builder, @"\styrsid", style.RevisionSaveId);
        WriteStyleFormatting(builder, style);
        WriteStyleParagraphFormatting(builder, style, unicodeSkipCount);
        WriteStyleTableFormatting(builder, style);

        builder.Append(' ');
        builder.Append(EscapeText(style.Name, unicodeSkipCount));
        builder.Append(";}");
    }

    private static void WriteStyleFormatting(StringBuilder builder, RtfStyle style) {
        AppendOptionalStyleToggle(builder, @"\b", style.Bold);
        AppendOptionalStyleToggle(builder, @"\i", style.Italic);

        if (style.UnderlineStyle.HasValue) {
            builder.Append(GetUnderlineControl(style.UnderlineStyle.Value));
        }

        if (style.FontSize.HasValue) {
            int halfPoints = (int)Math.Round(style.FontSize.Value * 2d, MidpointRounding.AwayFromZero);
            builder.Append(@"\fs");
            builder.Append(halfPoints.ToString(CultureInfo.InvariantCulture));
        }

        AppendOptionalTwips(builder, @"\f", style.FontId);
        AppendOptionalTwips(builder, @"\cf", style.ForegroundColorIndex);
        AppendOptionalTwips(builder, @"\highlight", style.HighlightColorIndex);
    }

    private static void WriteStyleParagraphFormatting(StringBuilder builder, RtfStyle style, int unicodeSkipCount) {
        AppendOptionalStyleToggle(builder, @"\pagebb", style.PageBreakBefore);
        AppendOptionalStyleToggle(builder, @"\keepn", style.KeepWithNext);
        AppendOptionalStyleToggle(builder, @"\keep", style.KeepLinesTogether);
        AppendOptionalStyleToggle(builder, @"\noline", style.SuppressLineNumbers);
        AppendOptionalStyleToggle(builder, @"\hyphpar", style.AutoHyphenation);
        AppendOptionalStyleToggle(builder, @"\contextualspace", style.ContextualSpacing);
        AppendOptionalStyleToggle(builder, @"\adjustright", style.AdjustRightIndent);

        if (style.SnapToLineGrid.HasValue) {
            builder.Append(style.SnapToLineGrid.Value ? @"\nosnaplinegrid0" : @"\nosnaplinegrid");
        }

        if (style.WidowControl.HasValue) {
            builder.Append(style.WidowControl.Value ? @"\widctlpar" : @"\nowidctlpar");
        }

        AppendOptionalTwips(builder, @"\outlinelevel", style.OutlineLevel);

        if (style.ParagraphDirection.HasValue) {
            builder.Append(style.ParagraphDirection.Value == RtfTextDirection.RightToLeft ? @"\rtlpar" : @"\ltrpar");
        }

        WriteParagraphFrame(builder, style.Frame);
        WriteLegacyNumbering(builder, style.LegacyNumbering, unicodeSkipCount);
        WriteTabStops(builder, style.TabStops);
        AppendOptionalTwips(builder, @"\li", style.LeftIndentTwips);
        AppendOptionalTwips(builder, @"\ri", style.RightIndentTwips);
        AppendOptionalTwips(builder, @"\fi", style.FirstLineIndentTwips);
        AppendOptionalTwips(builder, @"\sb", style.SpaceBeforeTwips);
        AppendOptionalTwips(builder, @"\sa", style.SpaceAfterTwips);
        AppendOptionalBinary(builder, @"\sbauto", style.SpaceBeforeAuto);
        AppendOptionalBinary(builder, @"\saauto", style.SpaceAfterAuto);
        AppendOptionalTwips(builder, @"\sl", style.LineSpacingTwips);
        AppendOptionalBinary(builder, @"\slmult", style.LineSpacingMultiple);
        AppendOptionalTwips(builder, @"\cbpat", style.BackgroundColorIndex);
        AppendOptionalTwips(builder, @"\cfpat", style.ShadingForegroundColorIndex);
        AppendOptionalTwips(builder, @"\shading", style.ShadingPatternPercent);
        WriteParagraphShadingPattern(builder, style.ShadingPattern);
        WriteParagraphBorder(builder, @"\brdrt", style.TopBorder);
        WriteParagraphBorder(builder, @"\brdrl", style.LeftBorder);
        WriteParagraphBorder(builder, @"\brdrb", style.BottomBorder);
        WriteParagraphBorder(builder, @"\brdrr", style.RightBorder);

        if (style.ParagraphAlignment.HasValue) {
            builder.Append(style.ParagraphAlignment.Value switch {
                RtfTextAlignment.Center => @"\qc",
                RtfTextAlignment.Right => @"\qr",
                RtfTextAlignment.Justify => @"\qj",
                _ => @"\ql"
            });
        }
    }

    private static void WriteStyleTableFormatting(StringBuilder builder, RtfStyle style) {
        if (style.Kind != RtfStyleKind.Table || !HasStyleTableFormatting(style.TableRowFormat)) {
            return;
        }

        RtfTableRow row = style.TableRowFormat;
        builder.Append(@"\tsrowd");
        if (row.RepeatHeader) {
            builder.Append(@"\trhdr");
        }

        if (row.KeepTogether) {
            builder.Append(@"\trkeep");
        }

        if (row.KeepWithNext) {
            builder.Append(@"\trkeepfollow");
        }

        AppendOptionalBinary(builder, @"\trautofit", row.AutoFit);
        if (row.Direction.HasValue) {
            builder.Append(row.Direction.Value == RtfTableRowDirection.RightToLeft ? @"\rtlrow" : @"\ltrrow");
        }

        AppendOptionalTwips(builder, @"\trrh", row.HeightTwips);
        AppendOptionalTwips(builder, @"\trgaph", row.CellGapTwips);
        AppendOptionalTwips(builder, @"\trleft", row.LeftIndentTwips);
        WriteTablePreferredWidth(builder, row);
        AppendOptionalTwips(builder, @"\trcbpat", row.BackgroundColorIndex);
        AppendOptionalTwips(builder, @"\trcfpat", row.ShadingForegroundColorIndex);
        AppendOptionalTwips(builder, @"\trpat", row.ShadingPatternValue);
        AppendOptionalTwips(builder, @"\trshdng", row.ShadingPatternPercent);
        WriteTableRowShadingPattern(builder, row.ShadingPattern);
        WriteTableRowBox(builder, @"\trpaddt", @"\trpaddft", row.PaddingTopTwips);
        WriteTableRowBox(builder, @"\trpaddl", @"\trpaddfl", row.PaddingLeftTwips);
        WriteTableRowBox(builder, @"\trpaddb", @"\trpaddfb", row.PaddingBottomTwips);
        WriteTableRowBox(builder, @"\trpaddr", @"\trpaddfr", row.PaddingRightTwips);
        WriteTableRowBox(builder, @"\trspdt", @"\trspdft", row.SpacingTopTwips);
        WriteTableRowBox(builder, @"\trspdl", @"\trspdfl", row.SpacingLeftTwips);
        WriteTableRowBox(builder, @"\trspdb", @"\trspdfb", row.SpacingBottomTwips);
        WriteTableRowBox(builder, @"\trspdr", @"\trspdfr", row.SpacingRightTwips);
        WriteTableRowPositioning(builder, row);
        if (row.Alignment.HasValue) {
            builder.Append(row.Alignment.Value switch {
                RtfTableAlignment.Center => @"\trqc",
                RtfTableAlignment.Right => @"\trqr",
                _ => @"\trql"
            });
        }

        WriteTableRowBorder(builder, @"\trbrdrt", row.TopBorder);
        WriteTableRowBorder(builder, @"\trbrdrl", row.LeftBorder);
        WriteTableRowBorder(builder, @"\trbrdrb", row.BottomBorder);
        WriteTableRowBorder(builder, @"\trbrdrr", row.RightBorder);
        WriteTableRowBorder(builder, @"\trbrdrh", row.HorizontalBorder);
        WriteTableRowBorder(builder, @"\trbrdrv", row.VerticalBorder);

        int boundary = 0;
        foreach (RtfTableCell cell in row.Cells) {
            WriteCellDefinition(builder, cell);
            boundary = cell.RightBoundaryTwips ?? boundary + 2400;
            builder.Append(@"\cellx");
            builder.Append(boundary.ToString(CultureInfo.InvariantCulture));
        }
    }

    private static bool HasStyleTableFormatting(RtfTableRow row) {
        return row.RepeatHeader ||
               row.KeepTogether ||
               row.KeepWithNext ||
               row.AutoFit.HasValue ||
               row.Direction.HasValue ||
               row.HeightTwips.HasValue ||
               row.CellGapTwips.HasValue ||
               row.LeftIndentTwips.HasValue ||
               row.Alignment.HasValue ||
               row.PreferredWidth.HasValue ||
               row.PreferredWidthUnit.HasValue ||
               row.BackgroundColorIndex.HasValue ||
               row.ShadingForegroundColorIndex.HasValue ||
               row.ShadingPatternValue.HasValue ||
               row.ShadingPatternPercent.HasValue ||
               row.ShadingPattern != RtfShadingPattern.None ||
               row.PaddingTopTwips.HasValue ||
               row.PaddingLeftTwips.HasValue ||
               row.PaddingBottomTwips.HasValue ||
               row.PaddingRightTwips.HasValue ||
               row.SpacingTopTwips.HasValue ||
               row.SpacingLeftTwips.HasValue ||
               row.SpacingBottomTwips.HasValue ||
               row.SpacingRightTwips.HasValue ||
               row.NoOverlap ||
               row.HorizontalAnchor.HasValue ||
               row.VerticalAnchor.HasValue ||
               row.HorizontalPosition.HasValue ||
               row.HorizontalPositionTwips.HasValue ||
               row.VerticalPosition.HasValue ||
               row.VerticalPositionTwips.HasValue ||
               row.TextWrapLeftTwips.HasValue ||
               row.TextWrapRightTwips.HasValue ||
               row.TextWrapTopTwips.HasValue ||
               row.TextWrapBottomTwips.HasValue ||
               row.TopBorder.HasAnyValue ||
               row.LeftBorder.HasAnyValue ||
               row.BottomBorder.HasAnyValue ||
               row.RightBorder.HasAnyValue ||
               row.HorizontalBorder.HasAnyValue ||
               row.VerticalBorder.HasAnyValue ||
               row.Cells.Count > 0;
    }

    private static void AppendOptionalStyleToggle(StringBuilder builder, string control, bool? value) {
        if (!value.HasValue) return;

        builder.Append(control);
        if (!value.Value) {
            builder.Append('0');
        }
    }

    private static void AppendStyleFlag(StringBuilder builder, string control, bool value) {
        if (value) {
            builder.Append(control);
        }
    }

    private static void WriteStyleKeyCode(StringBuilder builder, RtfStyleKeyCode? keyCode, int unicodeSkipCount) {
        if (keyCode == null) return;

        builder.Append(@"{\*\keycode");
        if (keyCode.Shift) {
            builder.Append(@"\shift");
        }

        if (keyCode.Control) {
            builder.Append(@"\ctrl");
        }

        if (keyCode.Alt) {
            builder.Append(@"\alt");
        }

        AppendOptionalTwips(builder, @"\fn", keyCode.FunctionKey);
        if (!string.IsNullOrWhiteSpace(keyCode.Key)) {
            builder.Append(' ');
            builder.Append(EscapeText(keyCode.Key!.Trim(), unicodeSkipCount));
        }

        builder.Append('}');
    }
}

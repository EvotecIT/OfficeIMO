namespace OfficeIMO.Rtf.Writing;

internal static partial class RtfDocumentWriter {
    private static void WriteTable(StringBuilder builder, RtfTable table, int? defaultLanguageId, int unicodeSkipCount) {
        foreach (RtfTableRow row in table.Rows) {
            builder.Append(@"\trowd\trgaph");
            builder.Append(row.CellGapTwips.GetValueOrDefault(108).ToString(CultureInfo.InvariantCulture));
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
            for (int index = 0; index < row.Cells.Count; index++) {
                RtfTableCell cell = row.Cells[index];
                WriteCellDefinition(builder, cell);
                boundary = cell.RightBoundaryTwips ?? boundary + 2400;
                builder.Append(@"\cellx");
                builder.Append(boundary.ToString(CultureInfo.InvariantCulture));
            }

            builder.AppendLine();
            foreach (RtfTableCell cell in row.Cells) {
                WriteCell(builder, cell, defaultLanguageId, unicodeSkipCount);
            }

            builder.Append(@"\row");
            builder.AppendLine();
        }
    }

    private static void WriteTableRowPositioning(StringBuilder builder, RtfTableRow row) {
        if (row.NoOverlap) {
            builder.Append(@"\tabsnoovrlp");
        }

        if (row.HorizontalAnchor.HasValue) {
            builder.Append(row.HorizontalAnchor.Value switch {
                RtfTableHorizontalAnchor.Margin => @"\tphmrg",
                RtfTableHorizontalAnchor.Page => @"\tphpg",
                _ => @"\tphcol"
            });
        }

        if (row.VerticalAnchor.HasValue) {
            builder.Append(row.VerticalAnchor.Value switch {
                RtfTableVerticalAnchor.Paragraph => @"\tpvpara",
                RtfTableVerticalAnchor.Page => @"\tpvpg",
                _ => @"\tpvmrg"
            });
        }

        WriteTableRowHorizontalPosition(builder, row.HorizontalPosition, row.HorizontalPositionTwips);
        WriteTableRowVerticalPosition(builder, row.VerticalPosition, row.VerticalPositionTwips);
        AppendOptionalTwips(builder, @"\tdfrmtxtLeft", row.TextWrapLeftTwips);
        AppendOptionalTwips(builder, @"\tdfrmtxtRight", row.TextWrapRightTwips);
        AppendOptionalTwips(builder, @"\tdfrmtxtTop", row.TextWrapTopTwips);
        AppendOptionalTwips(builder, @"\tdfrmtxtBottom", row.TextWrapBottomTwips);
    }

    private static void WriteTableRowHorizontalPosition(StringBuilder builder, RtfTableHorizontalPosition? position, int? twips) {
        if (!position.HasValue) return;

        switch (position.Value) {
            case RtfTableHorizontalPosition.Absolute:
                AppendOptionalTwips(builder, @"\tposx", twips);
                return;
            case RtfTableHorizontalPosition.NegativeAbsolute:
                AppendOptionalTwips(builder, @"\tposnegx", twips);
                return;
            case RtfTableHorizontalPosition.Center:
                builder.Append(@"\tposxc");
                return;
            case RtfTableHorizontalPosition.Right:
                builder.Append(@"\tposxr");
                return;
            case RtfTableHorizontalPosition.Inside:
                builder.Append(@"\tposxi");
                return;
            case RtfTableHorizontalPosition.Outside:
                builder.Append(@"\tposxo");
                return;
            default:
                builder.Append(@"\tposxl");
                return;
        }
    }

    private static void WriteTableRowVerticalPosition(StringBuilder builder, RtfTableVerticalPosition? position, int? twips) {
        if (!position.HasValue) return;

        switch (position.Value) {
            case RtfTableVerticalPosition.Absolute:
                AppendOptionalTwips(builder, @"\tposy", twips);
                return;
            case RtfTableVerticalPosition.NegativeAbsolute:
                AppendOptionalTwips(builder, @"\tposnegy", twips);
                return;
            case RtfTableVerticalPosition.Center:
                builder.Append(@"\tposyc");
                return;
            case RtfTableVerticalPosition.Bottom:
                builder.Append(@"\tposyb");
                return;
            case RtfTableVerticalPosition.Inline:
                builder.Append(@"\tposyil");
                return;
            case RtfTableVerticalPosition.Inside:
                builder.Append(@"\tposyin");
                return;
            case RtfTableVerticalPosition.Outside:
                builder.Append(@"\tposyoutv");
                return;
            default:
                builder.Append(@"\tposyt");
                return;
        }
    }

    private static void WriteTablePreferredWidth(StringBuilder builder, RtfTableRow row) {
        if (row.PreferredWidthUnit.HasValue) {
            builder.Append(@"\trftsWidth");
            builder.Append(ToRtfTableWidthUnitValue(row.PreferredWidthUnit.Value));
        }

        AppendOptionalTwips(builder, @"\trwWidth", row.PreferredWidth);
    }

    private static void WriteTableRowShadingPattern(StringBuilder builder, RtfShadingPattern pattern) {
        string? control = pattern switch {
            RtfShadingPattern.Horizontal => @"\trbghoriz",
            RtfShadingPattern.Vertical => @"\trbgvert",
            RtfShadingPattern.ForwardDiagonal => @"\trbgfdiag",
            RtfShadingPattern.BackwardDiagonal => @"\trbgbdiag",
            RtfShadingPattern.Cross => @"\trbgcross",
            RtfShadingPattern.DiagonalCross => @"\trbgdcross",
            RtfShadingPattern.DarkHorizontal => @"\trbgdkhor",
            RtfShadingPattern.DarkVertical => @"\trbgdkvert",
            RtfShadingPattern.DarkForwardDiagonal => @"\trbgdkfdiag",
            RtfShadingPattern.DarkBackwardDiagonal => @"\trbgdkbdiag",
            RtfShadingPattern.DarkCross => @"\trbgdkcross",
            RtfShadingPattern.DarkDiagonalCross => @"\trbgdkdcross",
            _ => null
        };
        if (control != null) {
            builder.Append(control);
        }
    }

    private static void WriteTableRowBox(StringBuilder builder, string valueControl, string unitControl, int? twips) {
        if (!twips.HasValue) return;

        builder.Append(valueControl);
        builder.Append(twips.Value.ToString(CultureInfo.InvariantCulture));
        builder.Append(unitControl);
        builder.Append('3');
    }

    private static string ToRtfTableWidthUnitValue(RtfTableWidthUnit unit) {
        return unit switch {
            RtfTableWidthUnit.Auto => "1",
            RtfTableWidthUnit.Percent => "2",
            _ => "3"
        };
    }

    private static void WriteTableRowBorder(StringBuilder builder, string sideControl, RtfTableRowBorder border) {
        if (!border.HasAnyValue) return;

        builder.Append(sideControl);
        builder.Append(border.Style switch {
            RtfTableCellBorderStyle.Double => @"\brdrdb",
            RtfTableCellBorderStyle.Dotted => @"\brdrdot",
            RtfTableCellBorderStyle.Dashed => @"\brdrdash",
            RtfTableCellBorderStyle.None => @"\brdrnil",
            _ => @"\brdrs"
        });
        AppendOptionalTwips(builder, @"\brdrw", border.Width);
        AppendOptionalTwips(builder, @"\brdrcf", border.ColorIndex);
    }

    private static void WriteCellDefinition(StringBuilder builder, RtfTableCell cell) {
        if (cell.HorizontalMerge == RtfTableCellMerge.First) {
            builder.Append(@"\clmgf");
        } else if (cell.HorizontalMerge == RtfTableCellMerge.Continue) {
            builder.Append(@"\clmrg");
        }

        if (cell.VerticalMerge == RtfTableCellMerge.First) {
            builder.Append(@"\clvmgf");
        } else if (cell.VerticalMerge == RtfTableCellMerge.Continue) {
            builder.Append(@"\clvmrg");
        }

        if (cell.PreferredWidthUnit.HasValue) {
            builder.Append(@"\clftsWidth");
            builder.Append(ToRtfTableWidthUnitValue(cell.PreferredWidthUnit.Value));
        }

        AppendOptionalTwips(builder, @"\clwWidth", cell.PreferredWidth);
        if (cell.HideCellMark) {
            builder.Append(@"\clhidemark");
        }

        if (cell.NoWrap) {
            builder.Append(@"\clNoWrap");
        }

        if (cell.FitText) {
            builder.Append(@"\clFitText");
        }

        AppendOptionalTwips(builder, @"\clcbpat", cell.BackgroundColorIndex);
        AppendOptionalTwips(builder, @"\clcfpat", cell.ShadingForegroundColorIndex);
        AppendOptionalTwips(builder, @"\clshdng", cell.ShadingPatternPercent);
        WriteCellShadingPattern(builder, cell.ShadingPattern);
        if (cell.VerticalAlignment.HasValue) {
            builder.Append(cell.VerticalAlignment.Value switch {
                RtfTableCellVerticalAlignment.Center => @"\clvertalc",
                RtfTableCellVerticalAlignment.Bottom => @"\clvertalb",
                _ => @"\clvertalt"
            });
        }

        WriteCellTextFlow(builder, cell.TextFlow);
        WriteCellBorder(builder, @"\clbrdrt", cell.TopBorder);
        WriteCellBorder(builder, @"\clbrdrl", cell.LeftBorder);
        WriteCellBorder(builder, @"\clbrdrb", cell.BottomBorder);
        WriteCellBorder(builder, @"\clbrdrr", cell.RightBorder);
        WriteCellBorder(builder, @"\cldglu", cell.TopLeftToBottomRightBorder);
        WriteCellBorder(builder, @"\cldgll", cell.TopRightToBottomLeftBorder);
        WriteCellPadding(builder, @"\clpadt", @"\clpadft", cell.PaddingTopTwips);
        WriteCellPadding(builder, @"\clpadl", @"\clpadfl", cell.PaddingLeftTwips);
        WriteCellPadding(builder, @"\clpadb", @"\clpadfb", cell.PaddingBottomTwips);
        WriteCellPadding(builder, @"\clpadr", @"\clpadfr", cell.PaddingRightTwips);
    }

    private static void WriteCellShadingPattern(StringBuilder builder, RtfShadingPattern pattern) {
        string? control = pattern switch {
            RtfShadingPattern.Horizontal => @"\clbghoriz",
            RtfShadingPattern.Vertical => @"\clbgvert",
            RtfShadingPattern.ForwardDiagonal => @"\clbgfdiag",
            RtfShadingPattern.BackwardDiagonal => @"\clbgbdiag",
            RtfShadingPattern.Cross => @"\clbgcross",
            RtfShadingPattern.DiagonalCross => @"\clbgdcross",
            RtfShadingPattern.DarkHorizontal => @"\clbgdkhor",
            RtfShadingPattern.DarkVertical => @"\clbgdkvert",
            RtfShadingPattern.DarkForwardDiagonal => @"\clbgdkfdiag",
            RtfShadingPattern.DarkBackwardDiagonal => @"\clbgdkbdiag",
            RtfShadingPattern.DarkCross => @"\clbgdkcross",
            RtfShadingPattern.DarkDiagonalCross => @"\clbgdkdcross",
            _ => null
        };
        if (control != null) {
            builder.Append(control);
        }
    }

    private static void WriteCellTextFlow(StringBuilder builder, RtfTableCellTextFlow? textFlow) {
        if (!textFlow.HasValue) {
            return;
        }

        builder.Append(textFlow.Value switch {
            RtfTableCellTextFlow.TopToBottomRightToLeft => @"\cltxtbrl",
            RtfTableCellTextFlow.BottomToTopLeftToRight => @"\cltxbtlr",
            RtfTableCellTextFlow.LeftToRightTopToBottomVertical => @"\cltxlrtbv",
            RtfTableCellTextFlow.TopToBottomRightToLeftVertical => @"\cltxtbrlv",
            _ => @"\cltxlrtb"
        });
    }

    private static void WriteCellBorder(StringBuilder builder, string sideControl, RtfTableCellBorder border) {
        if (!border.HasAnyValue) return;

        builder.Append(sideControl);
        builder.Append(border.Style switch {
            RtfTableCellBorderStyle.Double => @"\brdrdb",
            RtfTableCellBorderStyle.Dotted => @"\brdrdot",
            RtfTableCellBorderStyle.Dashed => @"\brdrdash",
            RtfTableCellBorderStyle.None => @"\brdrnil",
            _ => @"\brdrs"
        });
        AppendOptionalTwips(builder, @"\brdrw", border.Width);
        AppendOptionalTwips(builder, @"\brdrcf", border.ColorIndex);
    }

    private static void WriteCellPadding(StringBuilder builder, string valueControl, string unitControl, int? value) {
        if (!value.HasValue) return;

        builder.Append(valueControl);
        builder.Append(value.Value.ToString(CultureInfo.InvariantCulture));
        builder.Append(unitControl);
        builder.Append('3');
    }

    private static void WriteCell(StringBuilder builder, RtfTableCell cell, int? defaultLanguageId, int unicodeSkipCount) {
        if (cell.Blocks.Any(block => block is RtfTable)) {
            WriteCellWithNestedTables(builder, cell, defaultLanguageId, unicodeSkipCount);
            return;
        }

        if (cell.Paragraphs.Count == 0) {
            builder.Append(@"\pard\intbl \cell");
            return;
        }

        for (int i = 0; i < cell.Paragraphs.Count; i++) {
            RtfParagraph paragraph = cell.Paragraphs[i];
            WriteListText(builder, paragraph.ListText, defaultLanguageId, unicodeSkipCount);
            WriteParagraphStart(builder, paragraph, inTable: true, unicodeSkipCount);
            var state = new RunWriteState(defaultLanguageId);
            foreach (IRtfInline inline in paragraph.Inlines) {
                WriteInline(builder, inline, state, defaultLanguageId, unicodeSkipCount);
            }

            ResetRunState(builder, state);
            builder.Append(i == cell.Paragraphs.Count - 1 ? @"\cell" : @"\par");
        }
    }

    private static void WriteCellWithNestedTables(StringBuilder builder, RtfTableCell cell, int? defaultLanguageId, int unicodeSkipCount) {
        foreach (IRtfBlock block in cell.Blocks) {
            if (block is RtfParagraph paragraph) {
                WriteListText(builder, paragraph.ListText, defaultLanguageId, unicodeSkipCount);
                WriteParagraphStart(builder, paragraph, inTable: true, unicodeSkipCount);
                var state = new RunWriteState(defaultLanguageId);
                foreach (IRtfInline inline in paragraph.Inlines) WriteInline(builder, inline, state, defaultLanguageId, unicodeSkipCount);
                ResetRunState(builder, state);
                builder.Append(@"\par");
            } else if (block is RtfTable nested) {
                WriteNestedTable(builder, nested, defaultLanguageId, unicodeSkipCount, 2);
            }
        }

        builder.Append(@"\pard\intbl \cell");
    }

    private static void WriteNestedTable(StringBuilder builder, RtfTable table, int? defaultLanguageId, int unicodeSkipCount, int level) {
        foreach (RtfTableRow row in table.Rows) {
            foreach (RtfTableCell cell in row.Cells) {
                foreach (IRtfBlock block in cell.Blocks) {
                    if (block is RtfParagraph paragraph) {
                        WriteListText(builder, paragraph.ListText, defaultLanguageId, unicodeSkipCount);
                        WriteParagraphStart(builder, paragraph, inTable: true, unicodeSkipCount);
                        builder.Append(@"\itap");
                        builder.Append(level.ToString(CultureInfo.InvariantCulture));
                        builder.Append(' ');
                        var state = new RunWriteState(defaultLanguageId);
                        foreach (IRtfInline inline in paragraph.Inlines) WriteInline(builder, inline, state, defaultLanguageId, unicodeSkipCount);
                        ResetRunState(builder, state);
                        builder.Append(@"\par");
                    } else if (block is RtfTable nested) {
                        WriteNestedTable(builder, nested, defaultLanguageId, unicodeSkipCount, Math.Min(15, level + 1));
                    }
                }

                builder.Append(@"\nestcell{\nonesttables\par}");
            }

            builder.Append(@"\pard\intbl\itap");
            builder.Append(level.ToString(CultureInfo.InvariantCulture));
            builder.Append(@"{\*\nesttableprops\trowd\trgaph");
            builder.Append(row.CellGapTwips.GetValueOrDefault(108).ToString(CultureInfo.InvariantCulture));
            int boundary = 0;
            foreach (RtfTableCell cell in row.Cells) {
                WriteCellDefinition(builder, cell);
                boundary = cell.RightBoundaryTwips ?? boundary + 2400;
                builder.Append(@"\cellx");
                builder.Append(boundary.ToString(CultureInfo.InvariantCulture));
            }

            builder.Append(@"\nestrow}{\nonesttables\par}");
        }
    }
}

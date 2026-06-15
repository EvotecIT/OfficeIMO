namespace OfficeIMO.Rtf.Writing;

internal static partial class RtfDocumentWriter {
    private static void WriteSectionStart(StringBuilder builder, RtfSection section) {
        builder.Append(@"\sectd");
        builder.Append(section.BreakKind switch {
            RtfSectionBreakKind.Continuous => @"\sbknone",
            RtfSectionBreakKind.Column => @"\sbkcol",
            RtfSectionBreakKind.EvenPage => @"\sbkeven",
            RtfSectionBreakKind.OddPage => @"\sbkodd",
            _ => @"\sbkpage"
        });
        WritePageSetup(builder, section.PageSetup, isSection: true);
        WriteSectionVerticalAlignment(builder, section.VerticalAlignment);
        WriteSectionDirection(builder, section.Direction);
        WriteNoteSettings(builder, section.NoteSettings);
        WriteLineNumbering(builder, section.LineNumbering);
        AppendOptionalTwips(builder, @"\cols", section.ColumnCount);
        AppendOptionalTwips(builder, @"\colsx", section.ColumnSpaceTwips);
        WriteSectionColumns(builder, section);
        if (section.ColumnSeparator) {
            builder.Append(@"\linebetcol");
        }

        builder.AppendLine();
    }

    private static void WriteSectionColumns(StringBuilder builder, RtfSection section) {
        for (int index = 0; index < section.Columns.Count; index++) {
            RtfSectionColumn column = section.Columns[index];
            if (!column.HasAnyValue) {
                continue;
            }

            builder.Append(@"\colno");
            builder.Append((index + 1).ToString(CultureInfo.InvariantCulture));
            AppendOptionalTwips(builder, @"\colw", column.WidthTwips);
            AppendOptionalTwips(builder, @"\colsr", column.SpaceAfterTwips);
        }
    }

    private static void WriteSectionVerticalAlignment(StringBuilder builder, RtfSectionVerticalAlignment? alignment) {
        if (!alignment.HasValue) return;

        builder.Append(alignment.Value switch {
            RtfSectionVerticalAlignment.Center => @"\vertalc",
            RtfSectionVerticalAlignment.Bottom => @"\vertalb",
            RtfSectionVerticalAlignment.Justified => @"\vertalj",
            _ => @"\vertalt"
        });
    }

    private static void WriteSectionDirection(StringBuilder builder, RtfTextDirection? direction) {
        if (!direction.HasValue) return;

        builder.Append(direction.Value == RtfTextDirection.RightToLeft ? @"\rtlsect" : @"\ltrsect");
    }

    private static void WriteLineNumbering(StringBuilder builder, RtfLineNumbering lineNumbering) {
        if (!lineNumbering.HasAnyValue) return;

        AppendOptionalTwips(builder, @"\linemod", lineNumbering.CountBy);
        AppendOptionalTwips(builder, @"\linex", lineNumbering.DistanceFromTextTwips);
        AppendOptionalTwips(builder, @"\linestarts", lineNumbering.StartNumber);
        if (!lineNumbering.Restart.HasValue) return;

        builder.Append(lineNumbering.Restart.Value switch {
            RtfLineNumberRestart.EachPage => @"\lineppage",
            RtfLineNumberRestart.Continuous => @"\linecont",
            _ => @"\linerestart"
        });
    }

    private static void WritePageSetup(StringBuilder builder, RtfPageSetup pageSetup, bool isSection) {
        if (!pageSetup.HasAnyValue) return;

        AppendOptionalTwips(builder, isSection ? @"\pgwsxn" : @"\paperw", pageSetup.PaperWidthTwips);
        AppendOptionalTwips(builder, isSection ? @"\pghsxn" : @"\paperh", pageSetup.PaperHeightTwips);
        AppendOptionalTwips(builder, @"\psz", pageSetup.PrinterPaperSize);
        AppendOptionalTwips(builder, @"\binfsxn", pageSetup.FirstPagePaperSource);
        AppendOptionalTwips(builder, @"\binsxn", pageSetup.OtherPagesPaperSource);
        AppendOptionalTwips(builder, isSection ? @"\marglsxn" : @"\margl", pageSetup.MarginLeftTwips);
        AppendOptionalTwips(builder, isSection ? @"\margrsxn" : @"\margr", pageSetup.MarginRightTwips);
        AppendOptionalTwips(builder, isSection ? @"\margtsxn" : @"\margt", pageSetup.MarginTopTwips);
        AppendOptionalTwips(builder, isSection ? @"\margbsxn" : @"\margb", pageSetup.MarginBottomTwips);
        AppendOptionalTwips(builder, isSection ? @"\guttersxn" : @"\gutter", pageSetup.GutterWidthTwips);
        AppendOptionalTwips(builder, @"\headery", pageSetup.HeaderDistanceTwips);
        AppendOptionalTwips(builder, @"\footery", pageSetup.FooterDistanceTwips);
        if (pageSetup.RtlGutter) {
            builder.Append(@"\rtlgutter");
        }

        AppendOptionalTwips(builder, @"\pgnstarts", pageSetup.PageNumberStart);
        if (pageSetup.PageNumberRestart.HasValue) {
            builder.Append(pageSetup.PageNumberRestart.Value ? @"\pgnrestart" : @"\pgncont");
        }

        AppendOptionalTwips(builder, @"\pgnx", pageSetup.PageNumberPositionXTwips);
        AppendOptionalTwips(builder, @"\pgny", pageSetup.PageNumberPositionYTwips);
        WritePageNumberFormat(builder, pageSetup.PageNumberFormat);
        WritePageBorders(builder, pageSetup.PageBorders);

        if (pageSetup.Landscape) {
            builder.Append(isSection ? @"\lndscpsxn" : @"\landscape");
        }

        if (pageSetup.DifferentFirstPageHeaderFooter) {
            builder.Append(@"\titlepg");
        }
    }

    private static void WritePageNumberFormat(StringBuilder builder, RtfPageNumberFormat? format) {
        if (!format.HasValue) return;

        builder.Append(format.Value switch {
            RtfPageNumberFormat.UpperRoman => @"\pgnucrm",
            RtfPageNumberFormat.LowerRoman => @"\pgnlcrm",
            RtfPageNumberFormat.UpperLetter => @"\pgnucltr",
            RtfPageNumberFormat.LowerLetter => @"\pgnlcltr",
            RtfPageNumberFormat.DoubleByteDecimal => @"\pgndecd",
            _ => @"\pgndec"
        });
    }

    private static void WritePageBorders(StringBuilder builder, RtfPageBorders pageBorders) {
        if (!pageBorders.HasAnyValue) return;

        if (pageBorders.IncludeHeader) {
            builder.Append(@"\pgbrdrhead");
        }

        if (pageBorders.IncludeFooter) {
            builder.Append(@"\pgbrdrfoot");
        }

        WritePageBorderDisplayOptions(builder, pageBorders);
        if (pageBorders.SnapToPageBorder) {
            builder.Append(@"\pgbrdrsnap");
        }

        WritePageBorder(builder, @"\pgbrdrt", pageBorders.Top);
        WritePageBorder(builder, @"\pgbrdrb", pageBorders.Bottom);
        WritePageBorder(builder, @"\pgbrdrl", pageBorders.Left);
        WritePageBorder(builder, @"\pgbrdrr", pageBorders.Right);
    }

    private static void WritePageBorderDisplayOptions(StringBuilder builder, RtfPageBorders pageBorders) {
        if (!pageBorders.Scope.HasValue && !pageBorders.DisplayBehindText.HasValue && !pageBorders.OffsetFrom.HasValue) {
            return;
        }

        int value = pageBorders.Scope switch {
            RtfPageBorderScope.FirstPageInSection => 1,
            RtfPageBorderScope.AllExceptFirstPageInSection => 2,
            RtfPageBorderScope.WholeDocument => 3,
            _ => 0
        };
        if (pageBorders.DisplayBehindText == true) {
            value |= 8;
        }

        if (pageBorders.OffsetFrom == RtfPageBorderOffset.PageEdge) {
            value |= 32;
        }

        AppendOptionalTwips(builder, @"\pgbrdropt", value);
    }

    private static void WritePageBorder(StringBuilder builder, string sideControl, RtfPageBorder border) {
        if (!border.HasAnyValue) return;

        builder.Append(sideControl);
        builder.Append(border.Style switch {
            RtfPageBorderStyle.Double => @"\brdrdb",
            RtfPageBorderStyle.Dotted => @"\brdrdot",
            RtfPageBorderStyle.Dashed => @"\brdrdash",
            RtfPageBorderStyle.Shadow => @"\brdrsh",
            RtfPageBorderStyle.None => @"\brdrnil",
            _ => @"\brdrs"
        });
        AppendOptionalTwips(builder, @"\brdrw", border.Width);
        AppendOptionalTwips(builder, @"\brsp", border.Space);
        AppendOptionalTwips(builder, @"\brdrcf", border.ColorIndex);
        if (border.Frame) {
            builder.Append(@"\brdrframe");
        }
    }
}

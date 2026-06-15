using System.Globalization;

namespace OfficeIMO.Html.Rtf;

internal static partial class RtfHtmlWriter {
    private static void AppendSections(StringBuilder builder, RtfDocument document, RtfHtmlSaveOptions options, string newline) {
        for (int index = 0; index < document.Sections.Count; index++) {
            RtfSection section = document.Sections[index];
            if (index > 0) {
                builder.Append(newline);
            }

            builder.Append("<section data-officeimo-rtf-section=\"true\"");
            AppendMetadataAttribute(builder, "data-officeimo-rtf-section-layout", EncodeSectionLayout(section));
            AppendSectionStyle(builder, section);
            builder.Append('>');
            if (section.Blocks.Count > 0) {
                builder.Append(newline);
                AppendBlocks(builder, section.Blocks, options, document, newline);
                builder.Append(newline);
            }

            builder.Append("</section>");
        }
    }

    private static void AppendSectionStyle(StringBuilder builder, RtfSection section) {
        var style = new StringBuilder();
        if (section.BreakKind != RtfSectionBreakKind.Continuous) {
            style.Append("break-before:page;page-break-before:always;");
        }

        if (section.ColumnCount.HasValue) {
            style.Append("column-count:");
            style.Append(section.ColumnCount.Value.ToString(CultureInfo.InvariantCulture));
            style.Append(';');
        }

        if (section.ColumnSpaceTwips.HasValue) {
            style.Append("column-gap:");
            style.Append(FormatPoints(section.ColumnSpaceTwips.Value / 20d));
            style.Append("pt;");
        }

        if (section.Direction.HasValue) {
            AppendLanguageDirectionStyle(style, null, section.Direction);
        }

        if (style.Length == 0) {
            return;
        }

        builder.Append(" style=\"");
        builder.Append(EncodeAttribute(style.ToString()));
        builder.Append('"');
    }

    private static string? EncodeSectionLayout(RtfSection section) {
        if (!section.HasAnyLayoutValue) {
            return null;
        }

        var values = new Dictionary<string, string>(StringComparer.Ordinal);
        AddEnum(values, "break", (RtfSectionBreakKind?)section.BreakKind);
        AddNullableInt(values, "column.count", section.ColumnCount);
        AddNullableInt(values, "column.space", section.ColumnSpaceTwips);
        AddBool(values, "column.separator", section.ColumnSeparator);
        AddEnum(values, "verticalAlignment", section.VerticalAlignment);
        AddEnum(values, "direction", section.Direction);
        AddPageSetup(values, "page", section.PageSetup);
        AddNoteSettings(values, "note", section.NoteSettings);
        AddLineNumbering(values, "line", section.LineNumbering);
        for (int index = 0; index < section.Columns.Count; index++) {
            RtfSectionColumn column = section.Columns[index];
            AddNullableInt(values, "column." + index.ToString(CultureInfo.InvariantCulture) + ".width", column.WidthTwips);
            AddNullableInt(values, "column." + index.ToString(CultureInfo.InvariantCulture) + ".spaceAfter", column.SpaceAfterTwips);
        }

        return RtfHtmlMetadataCodec.Encode(values);
    }

    private static void AppendDocumentLayoutMetadata(StringBuilder builder, RtfDocument document, string newline) {
        AppendColorTableMetadata(builder, document, newline);

        var values = new Dictionary<string, string>(StringComparer.Ordinal);
        AddPageSetup(values, "page", document.PageSetup);
        AddNoteSettings(values, "note", document.NoteSettings);
        if (values.Count == 0) {
            return;
        }

        builder.Append(newline);
        builder.Append("<meta name=\"officeimo-rtf-document-layout\" content=\"");
        builder.Append(EncodeAttribute(RtfHtmlMetadataCodec.Encode(values)));
        builder.Append("\">");
    }

    private static void AppendColorTableMetadata(StringBuilder builder, RtfDocument document, string newline) {
        if (document.Colors.Count == 0) {
            return;
        }

        var values = new Dictionary<string, string>(StringComparer.Ordinal);
        for (int index = 0; index < document.Colors.Count; index++) {
            RtfColor color = document.Colors[index];
            string prefix = "color." + index.ToString(CultureInfo.InvariantCulture);
            AddInt(values, prefix + ".red", color.Red);
            AddInt(values, prefix + ".green", color.Green);
            AddInt(values, prefix + ".blue", color.Blue);
            AddEnum(values, prefix + ".theme", color.ThemeColor);
            AddNullableInt(values, prefix + ".tint", color.Tint);
            AddNullableInt(values, prefix + ".shade", color.Shade);
        }

        builder.Append(newline);
        builder.Append("<meta name=\"officeimo-rtf-colors\" content=\"");
        builder.Append(EncodeAttribute(RtfHtmlMetadataCodec.Encode(values)));
        builder.Append("\">");
    }

    private static void AddPageSetup(Dictionary<string, string> values, string prefix, RtfPageSetup pageSetup) {
        AddNullableInt(values, prefix + ".paperWidth", pageSetup.PaperWidthTwips);
        AddNullableInt(values, prefix + ".paperHeight", pageSetup.PaperHeightTwips);
        AddNullableInt(values, prefix + ".printerPaperSize", pageSetup.PrinterPaperSize);
        AddNullableInt(values, prefix + ".firstPagePaperSource", pageSetup.FirstPagePaperSource);
        AddNullableInt(values, prefix + ".otherPagesPaperSource", pageSetup.OtherPagesPaperSource);
        AddNullableInt(values, prefix + ".marginLeft", pageSetup.MarginLeftTwips);
        AddNullableInt(values, prefix + ".marginRight", pageSetup.MarginRightTwips);
        AddNullableInt(values, prefix + ".marginTop", pageSetup.MarginTopTwips);
        AddNullableInt(values, prefix + ".marginBottom", pageSetup.MarginBottomTwips);
        AddNullableInt(values, prefix + ".gutter", pageSetup.GutterWidthTwips);
        AddNullableInt(values, prefix + ".headerDistance", pageSetup.HeaderDistanceTwips);
        AddNullableInt(values, prefix + ".footerDistance", pageSetup.FooterDistanceTwips);
        AddNullableInt(values, prefix + ".pageNumberStart", pageSetup.PageNumberStart);
        AddNullableBool(values, prefix + ".pageNumberRestart", pageSetup.PageNumberRestart);
        AddNullableInt(values, prefix + ".pageNumberX", pageSetup.PageNumberPositionXTwips);
        AddNullableInt(values, prefix + ".pageNumberY", pageSetup.PageNumberPositionYTwips);
        AddEnum(values, prefix + ".pageNumberFormat", pageSetup.PageNumberFormat);
        AddBool(values, prefix + ".landscape", pageSetup.Landscape);
        AddBool(values, prefix + ".differentFirstPage", pageSetup.DifferentFirstPageHeaderFooter);
        AddBool(values, prefix + ".rtlGutter", pageSetup.RtlGutter);
        AddPageBorders(values, prefix + ".borders", pageSetup.PageBorders);
    }

    private static void AddPageBorders(Dictionary<string, string> values, string prefix, RtfPageBorders borders) {
        AddBool(values, prefix + ".includeHeader", borders.IncludeHeader);
        AddBool(values, prefix + ".includeFooter", borders.IncludeFooter);
        AddBool(values, prefix + ".snap", borders.SnapToPageBorder);
        AddEnum(values, prefix + ".scope", borders.Scope);
        AddNullableBool(values, prefix + ".behindText", borders.DisplayBehindText);
        AddEnum(values, prefix + ".offset", borders.OffsetFrom);
        AddPageBorder(values, prefix + ".top", borders.Top);
        AddPageBorder(values, prefix + ".bottom", borders.Bottom);
        AddPageBorder(values, prefix + ".left", borders.Left);
        AddPageBorder(values, prefix + ".right", borders.Right);
    }

    private static void AddPageBorder(Dictionary<string, string> values, string prefix, RtfPageBorder border) {
        AddEnum(values, prefix + ".style", border.Style == RtfPageBorderStyle.None ? (RtfPageBorderStyle?)null : border.Style);
        AddNullableInt(values, prefix + ".width", border.Width);
        AddNullableInt(values, prefix + ".space", border.Space);
        AddNullableInt(values, prefix + ".color", border.ColorIndex);
        AddBool(values, prefix + ".shadow", border.Shadow);
        AddBool(values, prefix + ".frame", border.Frame);
    }

    private static void AddNoteSettings(Dictionary<string, string> values, string prefix, RtfNoteSettings settings) {
        AddNullableInt(values, prefix + ".footnoteStart", settings.FootnoteStartNumber);
        AddEnum(values, prefix + ".footnoteRestart", settings.FootnoteRestart);
        AddEnum(values, prefix + ".footnoteFormat", settings.FootnoteNumberFormat);
        AddEnum(values, prefix + ".footnotePlacement", settings.FootnotePlacement);
        AddNullableInt(values, prefix + ".endnoteStart", settings.EndnoteStartNumber);
        AddEnum(values, prefix + ".endnoteRestart", settings.EndnoteRestart);
        AddEnum(values, prefix + ".endnoteFormat", settings.EndnoteNumberFormat);
        AddEnum(values, prefix + ".endnotePlacement", settings.EndnotePlacement);
    }

    private static void AddLineNumbering(Dictionary<string, string> values, string prefix, RtfLineNumbering lineNumbering) {
        AddNullableInt(values, prefix + ".countBy", lineNumbering.CountBy);
        AddNullableInt(values, prefix + ".distance", lineNumbering.DistanceFromTextTwips);
        AddNullableInt(values, prefix + ".start", lineNumbering.StartNumber);
        AddEnum(values, prefix + ".restart", lineNumbering.Restart);
    }

    private static void AddInt(Dictionary<string, string> values, string key, int value) {
        values[key] = value.ToString(CultureInfo.InvariantCulture);
    }

    private static void AddNullableInt(Dictionary<string, string> values, string key, int? value) {
        if (value.HasValue) {
            AddInt(values, key, value.Value);
        }
    }

    private static void AddBool(Dictionary<string, string> values, string key, bool value) {
        if (value) {
            values[key] = "true";
        }
    }

    private static void AddNullableBool(Dictionary<string, string> values, string key, bool? value) {
        if (value.HasValue) {
            values[key] = value.Value ? "true" : "false";
        }
    }

    private static void AddEnum<T>(Dictionary<string, string> values, string key, T? value) where T : struct {
        if (value.HasValue) {
            values[key] = value.Value.ToString()!;
        }
    }
}

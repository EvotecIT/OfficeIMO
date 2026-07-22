using System.Globalization;

namespace OfficeIMO.Html;

internal static partial class RtfHtmlWriter {
    internal static string Write(RtfDocument document, RtfToHtmlOptions options) {
        RtfTableTraversalGuard.ValidateDocument(document);
        string newline = options.GetNewLine();
        var builder = new StringBuilder();

        if (!options.FragmentOnly) {
            AppendDocumentStart(builder, document, options, newline);
        }

        if (document.Sections.Count > 0) {
            AppendSections(builder, document, options, newline);
        } else {
            AppendBlocks(builder, document.Blocks, options, document, newline);
        }

        if (!options.FragmentOnly) {
            builder.Append(newline);
            builder.Append("</body>");
            builder.Append(newline);
            builder.Append("</html>");
        }

        return builder.ToString();
    }

    private static void AppendBlocks(StringBuilder builder, IReadOnlyList<IRtfBlock> blocks, RtfToHtmlOptions options, RtfDocument document, string newline) {
        RtfListKind openList = RtfListKind.None;
        for (int i = 0; i < blocks.Count; i++) {
            IRtfBlock block = blocks[i];
            if (block is RtfParagraph paragraph && paragraph.ListKind != RtfListKind.None) {
                if (openList != paragraph.ListKind) {
                    CloseList(builder, openList);
                    OpenList(builder, paragraph.ListKind);
                    openList = paragraph.ListKind;
                }

                AppendParagraph(builder, paragraph, options, document);
            } else {
                CloseList(builder, openList);
                openList = RtfListKind.None;
                AppendBlock(builder, block, options, document);
            }

            if (i + 1 < blocks.Count) {
                builder.Append(newline);
            }
        }

        CloseList(builder, openList);
    }

    private static void AppendDocumentStart(StringBuilder builder, RtfDocument document, RtfToHtmlOptions options, string newline) {
        builder.Append("<!doctype html>");
        builder.Append(newline);
        builder.Append("<html");
        AppendLanguageDirectionAttributes(builder, document.Settings.DefaultLanguageId, document.Settings.Direction);
        AppendLanguageDirectionStyleAttribute(builder, document.Settings.DefaultLanguageId, document.Settings.Direction);
        builder.Append('>');
        builder.Append(newline);
        builder.Append("<head>");
        builder.Append(newline);
        builder.Append("<meta charset=\"utf-8\">");
        if (options.IncludeMetadata) {
            string? title = options.Title ?? document.Info.Title;
            if (!string.IsNullOrWhiteSpace(title)) {
                builder.Append(newline);
                builder.Append("<title>");
                builder.Append(Encode(title!));
                builder.Append("</title>");
            }

            AppendDocumentMetadata(builder, document, newline);
        }

        if (options.IncludeRoundTripMetadata) {
            AppendHeaderFooterMetadata(builder, document, options, newline);
            AppendDocumentLayoutMetadata(builder, document, newline);
            AppendDocumentSettingsMetadata(builder, document, newline);
            AppendFontTableMetadata(builder, document, newline);
            AppendStylesheetMetadata(builder, document, newline);
            AppendListTableMetadata(builder, document, newline);
            AppendUserPropertiesMetadata(builder, document, newline);
            AppendDocumentVariablesMetadata(builder, document, newline);
            AppendRevisionTablesMetadata(builder, document, newline);
            AppendFileReferencesMetadata(builder, document, newline);
            AppendXmlNamespacesMetadata(builder, document, newline);
        }
        builder.Append(newline);
        builder.Append("</head>");
        builder.Append(newline);
        builder.Append("<body>");
        builder.Append(newline);
    }

    private static void AppendBlock(StringBuilder builder, IRtfBlock block, RtfToHtmlOptions options, RtfDocument document) {
        switch (block) {
            case RtfParagraph paragraph:
                AppendParagraph(builder, paragraph, options, document);
                break;
            case RtfTable table:
                AppendTable(builder, table, options, document);
                break;
            case RtfImage image:
                AppendImage(builder, image, options);
                break;
            case RtfObject rtfObject:
                AppendObject(builder, rtfObject, options, document, blockTag: true);
                break;
            case RtfShape shape:
                AppendShape(builder, shape, options, document, blockTag: true);
                break;
            default:
                break;
        }
    }

    private static void AppendParagraph(StringBuilder builder, RtfParagraph paragraph, RtfToHtmlOptions options, RtfDocument document) {
        string tagName = GetParagraphTagName(paragraph);
        builder.Append('<');
        builder.Append(tagName);
        AppendLanguageDirectionAttributes(builder, null, paragraph.Direction);
        if (options.IncludeRoundTripMetadata) {
            AppendListAttributes(builder, paragraph);
            AppendParagraphStyleAttributes(builder, paragraph);
            AppendParagraphRevisionAttributes(builder, paragraph);
            AppendParagraphControlAttributes(builder, paragraph);
            AppendParagraphFrameAttributes(builder, paragraph);
        }
        AppendParagraphStyle(builder, paragraph, document);
        builder.Append('>');
        AppendInlines(builder, paragraph.Inlines, options, document);
        builder.Append("</");
        builder.Append(tagName);
        builder.Append('>');
    }

    private static string GetParagraphTagName(RtfParagraph paragraph) {
        if (paragraph.ListKind != RtfListKind.None) {
            return "li";
        }

        if (paragraph.OutlineLevel.HasValue && paragraph.OutlineLevel.Value >= 0 && paragraph.OutlineLevel.Value <= 5) {
            return "h" + (paragraph.OutlineLevel.Value + 1).ToString(CultureInfo.InvariantCulture);
        }

        return "p";
    }

    private static void AppendParagraphStyle(StringBuilder builder, RtfParagraph paragraph, RtfDocument document) {
        if (!TryGetParagraphStyle(paragraph, document, out string? style)) {
            return;
        }

        builder.Append(" style=\"");
        builder.Append(EncodeAttribute(style!));
        builder.Append("\"");
    }

    private static bool TryGetParagraphStyle(RtfParagraph paragraph, RtfDocument document, out string? style) {
        var builder = new StringBuilder();
        if (TryGetColor(document, paragraph.BackgroundColorIndex, out RtfColor? background)) {
            builder.Append("background-color:");
            builder.Append(FormatColor(background!));
            builder.Append(';');
        }

        AppendParagraphShadingStyle(builder, paragraph, document);

        string? align = paragraph.Alignment == RtfTextAlignment.Left ? null : paragraph.Alignment.ToString().ToLowerInvariant();
        if (align != null) {
            builder.Append("text-align:");
            builder.Append(align);
            builder.Append(';');
        }

        if (paragraph.PageBreakBefore) {
            builder.Append("page-break-before:always;break-before:page;");
        }

        AppendTwipStyle(builder, "margin-left", paragraph.LeftIndentTwips);
        AppendTwipStyle(builder, "margin-right", paragraph.RightIndentTwips);
        AppendTwipStyle(builder, "text-indent", paragraph.FirstLineIndentTwips);
        AppendTwipStyle(builder, "margin-top", paragraph.SpaceBeforeTwips);
        AppendTwipStyle(builder, "margin-bottom", paragraph.SpaceAfterTwips);
        AppendLineHeightStyle(builder, paragraph);
        AppendParagraphBorderStyle(builder, "border-top", paragraph.TopBorder, document);
        AppendParagraphBorderStyle(builder, "border-left", paragraph.LeftBorder, document);
        AppendParagraphBorderStyle(builder, "border-bottom", paragraph.BottomBorder, document);
        AppendParagraphBorderStyle(builder, "border-right", paragraph.RightBorder, document);
        AppendLanguageDirectionStyle(builder, null, paragraph.Direction);

        style = builder.Length == 0 ? null : builder.ToString();
        return style != null;
    }

    private static void AppendTwipStyle(StringBuilder builder, string name, int? twips) {
        if (!twips.HasValue || twips.Value == 0) {
            return;
        }

        builder.Append(name);
        builder.Append(':');
        builder.Append(FormatPoints(twips.Value / 20d));
        builder.Append("pt;");
    }

    private static void AppendLineHeightStyle(StringBuilder builder, RtfParagraph paragraph) {
        if (!paragraph.LineSpacingTwips.HasValue || paragraph.LineSpacingTwips.Value == 0) {
            return;
        }

        builder.Append("line-height:");
        if (paragraph.LineSpacingMultiple == true) {
            builder.Append(FormatPoints(paragraph.LineSpacingTwips.Value / 240d));
        } else {
            builder.Append(FormatPoints(paragraph.LineSpacingTwips.Value / 20d));
            builder.Append("pt");
        }

        builder.Append(';');
    }

    private static void AppendParagraphBorderStyle(StringBuilder builder, string name, RtfParagraphBorder border, RtfDocument document) {
        if (!border.HasAnyValue) {
            return;
        }

        builder.Append(name);
        builder.Append(':');
        if (border.Width.HasValue) {
            builder.Append(FormatPoints(border.Width.Value / 20d));
            builder.Append("pt ");
        }

        builder.Append(FormatParagraphBorderStyle(border.Style));
        if (TryGetColor(document, border.ColorIndex, out RtfColor? color)) {
            builder.Append(' ');
            builder.Append(FormatColor(color!));
        }

        builder.Append(';');
    }

    private static void AppendInlines(StringBuilder builder, IReadOnlyList<IRtfInline> inlines, RtfToHtmlOptions options, RtfDocument document) {
        foreach (IRtfInline inline in inlines) {
            switch (inline) {
                case RtfRun run:
                    AppendRun(builder, run, options, document);
                    break;
                case RtfBreak rtfBreak:
                    AppendBreak(builder, rtfBreak.Kind, options.IncludeRoundTripMetadata);
                    break;
                case RtfField field:
                    AppendField(builder, field, options, document);
                    break;
                case RtfGeneratedText generatedText:
                    AppendGeneratedText(builder, generatedText, options.IncludeRoundTripMetadata);
                    AppendNote(builder, generatedText.Note, options, document);
                    break;
                case RtfImage image:
                    AppendImage(builder, image, options);
                    break;
                case RtfObject rtfObject:
                    AppendObject(builder, rtfObject, options, document, blockTag: false);
                    break;
                case RtfShape shape:
                    AppendShape(builder, shape, options, document, blockTag: false);
                    break;
                case RtfBookmarkMarker marker:
                    AppendBookmarkMarker(builder, marker, options.IncludeRoundTripMetadata);
                    break;
            }
        }
    }

    private static void AppendRun(StringBuilder builder, RtfRun run, RtfToHtmlOptions options, RtfDocument document) {
        bool revisionOpened = AppendRevisionStart(builder, run, document, options.IncludeRoundTripMetadata);
        int opened = 0;
        string? hyperlink = ResolveHtmlUrl(run.Hyperlink?.ToString(), options, "RtfHtmlHyperlinkRejected", "run.Hyperlink");
        if (hyperlink != null) {
            builder.Append("<a href=\"");
            builder.Append(EncodeAttribute(hyperlink));
            builder.Append("\">");
            opened++;
        }

        OpenRunStyle(builder, run, document, options.IncludeRoundTripMetadata, ref opened);
        OpenTag(builder, "strong", run.Bold, ref opened);
        OpenTag(builder, "em", run.Italic, ref opened);
        bool plainUnderline = run.Underline && !HasRichUnderline(run);
        OpenTag(builder, "u", plainUnderline, ref opened);
        bool plainStrike = (run.Strike || run.DoubleStrike) && !HasRichStrike(run);
        OpenTag(builder, "s", plainStrike, ref opened);
        OpenTag(builder, "sup", run.VerticalPosition == RtfVerticalPosition.Superscript, ref opened);
        OpenTag(builder, "sub", run.VerticalPosition == RtfVerticalPosition.Subscript, ref opened);

        builder.Append(Encode(run.Text));

        CloseTag(builder, "sub", run.VerticalPosition == RtfVerticalPosition.Subscript);
        CloseTag(builder, "sup", run.VerticalPosition == RtfVerticalPosition.Superscript);
        CloseTag(builder, "s", plainStrike);
        CloseTag(builder, "u", plainUnderline);
        CloseTag(builder, "em", run.Italic);
        CloseTag(builder, "strong", run.Bold);
        CloseRunStyle(builder, run, document);
        if (hyperlink != null) {
            builder.Append("</a>");
        }

        AppendNote(builder, run.Note, options, document);
        AppendRevisionEnd(builder, run, revisionOpened);
    }

    private static void OpenTag(StringBuilder builder, string tag, bool condition, ref int opened) {
        if (!condition) {
            return;
        }

        builder.Append('<');
        builder.Append(tag);
        builder.Append('>');
        opened++;
    }

    private static void CloseTag(StringBuilder builder, string tag, bool condition) {
        if (!condition) {
            return;
        }

        builder.Append("</");
        builder.Append(tag);
        builder.Append('>');
    }

    private static void AppendImage(StringBuilder builder, RtfImage image, RtfToHtmlOptions options) {
        string? source = ResolveImageSource(image, options);
        if (source == null) {
            return;
        }

        builder.Append("<img src=\"");
        builder.Append(EncodeAttribute(source));
        builder.Append('"');
        if (!string.IsNullOrWhiteSpace(image.Description)) {
            builder.Append(" alt=\"");
            builder.Append(EncodeAttribute(image.Description!));
            builder.Append('"');
        }

        AppendImageSize(builder, image);
        builder.Append('>');
    }

    private static string? ResolveImageSource(RtfImage image, RtfToHtmlOptions options) {
        if (options.ImageSourceResolver != null) {
            string? resolved;
            try {
                resolved = options.ImageSourceResolver(image);
            } catch (Exception exception) {
                options.AddDiagnostic(
                    "RtfHtmlImageSourceResolverFailed",
                    "RTF image source resolution failed.",
                    image.Format.ToString(),
                    exception);
                return null;
            }

            string? safeSource = ResolveHtmlUrl(resolved, options, "RtfHtmlImageSourceRejected", "ImageSourceResolver");
            if (safeSource != null) {
                return safeSource;
            }
        }

        if (!options.EmbedImagesAsDataUri) {
            options.AddDiagnostic(
                "RtfHtmlImageEmbeddingDisabled",
                "RTF image was skipped because no accepted image source was returned and data URI embedding is disabled.",
                image.Format.ToString());
            return null;
        }

        if (image.Data.Length == 0) {
            options.AddDiagnostic(
                "RtfHtmlImageDataMissing",
                "RTF image was skipped because it does not contain image data.",
                image.Format.ToString());
            return null;
        }

        if (options.MaxEmbeddedImageBytes < 0 || image.Data.Length > options.MaxEmbeddedImageBytes) {
            options.AddDiagnostic(
                "RtfHtmlImageEmbeddingLimitExceeded",
                "RTF image was skipped because it exceeds the configured data URI embedding limit.",
                image.Data.Length.ToString(CultureInfo.InvariantCulture),
                action: RtfConversionAction.Blocked);
            return null;
        }

        if (!TryGetImageMediaType(image.Format, out string? mediaType)) {
            options.AddDiagnostic(
                "RtfHtmlImageFormatUnsupported",
                "RTF image was skipped because the image format is not supported by the HTML writer.",
                image.Format.ToString());
            return null;
        }

        return "data:" + mediaType + ";base64," + Convert.ToBase64String(image.Data);
    }

    private static void AppendImageSize(StringBuilder builder, RtfImage image) {
        if (image.SourceWidth.HasValue) {
            builder.Append(" width=\"");
            builder.Append(image.SourceWidth.Value.ToString(CultureInfo.InvariantCulture));
            builder.Append('"');
        }

        if (image.SourceHeight.HasValue) {
            builder.Append(" height=\"");
            builder.Append(image.SourceHeight.Value.ToString(CultureInfo.InvariantCulture));
            builder.Append('"');
        }

        if (image.DesiredWidthTwips.HasValue || image.DesiredHeightTwips.HasValue) {
            builder.Append(" style=\"");
            if (image.DesiredWidthTwips.HasValue) {
                builder.Append("width:");
                builder.Append(FormatPoints(image.DesiredWidthTwips.Value / 20d));
                builder.Append("pt;");
            }

            if (image.DesiredHeightTwips.HasValue) {
                builder.Append("height:");
                builder.Append(FormatPoints(image.DesiredHeightTwips.Value / 20d));
                builder.Append("pt;");
            }

            builder.Append('"');
        }
    }

    private static void OpenList(StringBuilder builder, RtfListKind kind) {
        if (kind == RtfListKind.None) {
            return;
        }

        builder.Append(kind == RtfListKind.Decimal ? "<ol>" : "<ul>");
    }

    private static void CloseList(StringBuilder builder, RtfListKind kind) {
        if (kind == RtfListKind.None) {
            return;
        }

        builder.Append(kind == RtfListKind.Decimal ? "</ol>" : "</ul>");
    }

    private static bool TryGetImageMediaType(RtfImageFormat format, out string? mediaType) {
        switch (format) {
            case RtfImageFormat.Png:
                mediaType = "image/png";
                return true;
            case RtfImageFormat.Jpeg:
                mediaType = "image/jpeg";
                return true;
            default:
                mediaType = null;
                return false;
        }
    }

    private static void OpenRunStyle(StringBuilder builder, RtfRun run, RtfDocument document, bool includeRoundTripMetadata, ref int opened) {
        bool hasStyle = TryGetRunStyle(run, document, out string? style);
        if (!hasStyle && !run.StyleId.HasValue && FormatLanguageTag(run.LanguageId) == null && FormatTextDirection(run.Direction) == null) {
            return;
        }

        builder.Append("<span");
        if (includeRoundTripMetadata) {
            AppendRunStyleAttributes(builder, run);
        }
        AppendLanguageDirectionAttributes(builder, run.LanguageId, run.Direction);
        if (hasStyle) {
            builder.Append(" style=\"");
            builder.Append(EncodeAttribute(style!));
            builder.Append('"');
        }

        builder.Append('>');
        opened++;
    }

    private static void CloseRunStyle(StringBuilder builder, RtfRun run, RtfDocument document) {
        if (HasRunSpan(run, document)) {
            builder.Append("</span>");
        }
    }

    private static bool TryGetRunStyle(RtfRun run, RtfDocument document, out string? style) {
        var builder = new StringBuilder();
        if (TryGetFont(document, run.FontId, out RtfFont? font)) {
            builder.Append("font-family:");
            builder.Append(FormatFontFamily(font!.Name));
            builder.Append(';');
        }

        if (run.FontSize.HasValue) {
            builder.Append("font-size:");
            builder.Append(FormatPoints(run.FontSize.Value));
            builder.Append("pt;");
        }

        if (TryGetColor(document, run.ForegroundColorIndex, out RtfColor? foreground)) {
            builder.Append("color:");
            builder.Append(FormatColor(foreground!));
            builder.Append(';');
        }

        int? backgroundIndex = run.CharacterBackgroundColorIndex ?? run.HighlightColorIndex;
        if (TryGetColor(document, backgroundIndex, out RtfColor? background)) {
            builder.Append("background-color:");
            builder.Append(FormatColor(background!));
            builder.Append(';');
        }

        AppendCharacterShadingStyle(builder, run, document);

        AppendCharacterBorderStyle(builder, run.CharacterBorder, document);
        AppendTextDecorationStyle(builder, run, document);
        AppendCapsStyle(builder, run.CapsStyle);
        AppendCharacterEffectsStyle(builder, run);
        AppendCharacterMetricsStyle(builder, run);
        AppendLanguageDirectionStyle(builder, run);

        style = builder.Length == 0 ? null : builder.ToString();
        return style != null;
    }

    private static bool TryGetColor(RtfDocument document, int? index, out RtfColor? color) {
        if (!index.HasValue || index.Value <= 0 || index.Value > document.Colors.Count) {
            color = null;
            return false;
        }

        color = document.Colors[index.Value - 1];
        return true;
    }

    private static bool TryGetFont(RtfDocument document, int? id, out RtfFont? font) {
        if (!id.HasValue) {
            font = null;
            return false;
        }

        font = document.Fonts.FirstOrDefault(item => item.Id == id.Value);
        return font != null;
    }

    private static string FormatFontFamily(string fontFamily) {
        return "\"" + fontFamily.Replace("\\", "\\\\").Replace("\"", "\\\"") + "\"";
    }

    private static string FormatPoints(double points) {
        return points.ToString("0.###", CultureInfo.InvariantCulture);
    }

    private static string FormatTableWidth(int width, RtfTableWidthUnit unit) {
        switch (unit) {
            case RtfTableWidthUnit.Percent:
                return FormatPoints(width / 50d) + "%";
            case RtfTableWidthUnit.Auto:
                return "auto";
            default:
                return FormatPoints(width / 20d) + "pt";
        }
    }

    private static string FormatCellVerticalAlignment(RtfTableCellVerticalAlignment alignment) {
        switch (alignment) {
            case RtfTableCellVerticalAlignment.Center:
                return "middle";
            case RtfTableCellVerticalAlignment.Bottom:
                return "bottom";
            default:
                return "top";
        }
    }

    private static string FormatTableAlignment(RtfTableAlignment alignment) {
        switch (alignment) {
            case RtfTableAlignment.Center:
                return "center";
            case RtfTableAlignment.Right:
                return "right";
            default:
                return "left";
        }
    }

    private static string FormatCellBorderStyle(RtfTableCellBorderStyle style) {
        switch (style) {
            case RtfTableCellBorderStyle.Double:
                return "double";
            case RtfTableCellBorderStyle.Dotted:
                return "dotted";
            case RtfTableCellBorderStyle.Dashed:
                return "dashed";
            case RtfTableCellBorderStyle.None:
                return "none";
            default:
                return "solid";
        }
    }

    private static string FormatParagraphBorderStyle(RtfParagraphBorderStyle style) {
        switch (style) {
            case RtfParagraphBorderStyle.Double:
                return "double";
            case RtfParagraphBorderStyle.Dotted:
                return "dotted";
            case RtfParagraphBorderStyle.Dashed:
                return "dashed";
            case RtfParagraphBorderStyle.None:
                return "none";
            default:
                return "solid";
        }
    }

    private static string FormatColor(RtfColor color) {
        return "#" + color.Red.ToString("X2", CultureInfo.InvariantCulture) +
               color.Green.ToString("X2", CultureInfo.InvariantCulture) +
               color.Blue.ToString("X2", CultureInfo.InvariantCulture);
    }

    private static string Encode(string value) => WebUtility.HtmlEncode(value);

    private static string EncodeAttribute(string value) => WebUtility.HtmlEncode(value);

    private static string? ResolveHtmlUrl(string? rawUrl, RtfToHtmlOptions options, string diagnosticCode, string source) {
        if (string.IsNullOrWhiteSpace(rawUrl)) {
            return null;
        }

        string resolved = HtmlUrlPolicyEvaluator.ResolveUrl(rawUrl, null, options.GetUrlPolicy());
        if (resolved.Length > 0) {
            return resolved;
        }

        options.AddDiagnostic(diagnosticCode, "URL was omitted because it was rejected by the configured HTML URL policy.", source, action: RtfConversionAction.Blocked);
        return null;
    }
}
